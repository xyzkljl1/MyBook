using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Proxy;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using MimeKit;
using static System.Net.Mime.MediaTypeNames;
using static MailKit.Telemetry;

namespace MyBook
{
    // 从雅虎邮箱拉信用卡账单
    class MailUtil
    {
        private const string ICBCProvider = "ICBC";
        private const string IBKRReportSender = "donotreply@interactivebrokers.com";
        IProxyClient proxy;
        string username;
        string apppasswd;
        DatabaseUtil database;
        public MailUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            // 为了支持gbk编码
            System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            proxy = new Socks5Client(config["socksproxy"], Int32.Parse(config["socksproxy_port"]!));
            username = config["yahoo_user"]!;
            apppasswd = config["yahoo_pass"]!;
            this.database = database;
        }
        public Account GetICBCAccount(string name, CurrencyType currencyType)
        {
            return database.GetAccountByTypeAndId(ICBCProvider, name.Substring(0, 4), currencyType);
        }
        // 工行对账单，按卡号区分用途
        public async Task FetchICBCBill(DateTime date)
        {
            // 按月份搜索
            var monthText = date.ToString("yyyy-MM");
            if (database.IsStatementImported(ICBCProvider, monthText))
                return;

            var billText = await SearchBill("webmaster@icbc.com.cn", "中国工商银行客户对账单", date);
            if (!String.IsNullOrEmpty(billText))
            {
                Records records = new();
                try
                {
                    var tables = FormUtil.ReadFromHTML(billText);
                    if (tables.Count < 4 
                        || tables[0].Title != "需 还 款 明 细" 
                        || tables[1].Title != "本 期 交 易 汇 总"
                        || tables[2].Title != "人民币(本位币) 交 易 明 细"
                        || tables[3].Title != "外 币 交 易 明 细")
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid Tables");
                    foreach (var line in tables[1].Rows)
                    {
                        if (line.Count < 5 || line[0] == "合计")
                            continue;
                        var balance = Currency.Parse(line[4]);
                        var account = GetICBCAccount(line[0], balance.t);
                        //else
                        //    account.v = balance;
                    }

                    {
                        var table = tables[2];
                        if (table.Headers.Count!= 7
                            || table.Headers[0] != "卡号后四位"
                            || table.Headers[1] != "交易日"
                            || table.Headers[3] != "交易类型"
                            || table.Headers[4] != "商户名称/城市"
                            || table.Headers[6] != "记账金额/币种")
                            throw new MailParseException("Parse ICBC Bill Fail, Invalid Headers");
                        foreach (var line in table.Rows)
                        {
                            var record = new Record();
                            record.updateTime = DateTime.Now;
                            record.Account = GetICBCAccount(line[0], CurrencyType.RMB);
                            record.date = DateTime.ParseExact(line[1], "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            if (line[6].Contains("存入"))
                                record.isIn = true;
                            else if (line[6].Contains("支出"))
                                record.isIn = false;
                            else
                                throw new MailParseException("Parse ICBC Bill Fail");
                            record.Source = $"ICBC对账单邮件({monthText})/人民币明细/{DateTime.Now}/{string.Join(",", line)}";
                            record.DestAccount = line[4];
                            record.DescCurrency = Currency.Parse(line[5]);
                            record.CopyFrom(Currency.Parse(line[6]));
                            if (line[3]== "消费" || line[3] == "跨行消费" || line[3] == "境外消费")
                            {
                                record.Reason = record.Account.desc; // 工行按卡区分用途
                                records.Add(record);
                            }
                            else if (line[3] =="退款" || line[3] == "境外退货")
                            {
                                if (!record.isIn)
                                    throw new MailParseException("Parse ICBC Bill Fail, Invalid In");
                                //在同一个月内向前搜索对应的消费，尝试消除;比较DescCurrency因为退款是按交易金额退的
                                Record? destRecord = records.FindLast(destRecord => 
                                                            destRecord.DestAccount == record.DestAccount && destRecord.isIn == false
                                                            && destRecord.Account == record.Account && destRecord.DescCurrency == record.DescCurrency);
                                if (destRecord is not null)
                                    records.Remove(destRecord);
                                else // 不能消除则入账
                                {
                                    record.Reason = record.Account.desc; // 工行按卡区分用途
                                    records.Add(record);
                                }
                            }
                            else if (line[3] == "人民币自动转账还款" || line[3] == "自动购汇还款")
                            {
                                record.isInternal = true;
                                record.Reason = "信用卡还款";
                                records.Add(record);
                            }
                        }

                    }

                    database.SaveStatementRecordsOnce(ICBCProvider, monthText, records);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"parse mail fail :{e.Message}");
                }
            }
        }
        public async Task FetchIBKRReport(DateTime date)
        {
            var subjectDate = date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            var message = await SearchBill(
                IBKRReportSender,
                subjectDate,
                date,
                message => IsIBKRCustomActivityReportSubject(message.Subject, subjectDate));

            if (message is null)
            {
                Console.WriteLine($"Find no IBKR custom activity report {subjectDate}");
                return;
            }

            var bodyText = GetMessageText(message);
            var reportDate = ParseIBKRReportDate(bodyText);
            var reportAccountId = ParseIBKRAccountId(bodyText);
            if (reportDate is null || String.IsNullOrWhiteSpace(reportAccountId))
            {
                Console.WriteLine($"parse IBKR report mail fail: missing report date or account id, subject={message.Subject}");
                return;
            }

            var account = GetIBKRAccount(reportAccountId);

            var csvAttachments = ReadCsvAttachments(message);
            if (csvAttachments.Count == 0)
            {
                Console.WriteLine($"parse IBKR report mail fail: no csv attachment, subject={message.Subject}");
                return;
            }

            Console.WriteLine($"Find IBKR custom activity report: {reportDate.Value:yyyy-MM-dd} {reportAccountId} -> {account.name}, csv attachments={csvAttachments.Count}");
            foreach (var attachment in csvAttachments)
                Console.WriteLine($"Load IBKR report csv in memory: {attachment.FileName}, bytes={attachment.Content.Length}");

            // TODO: parse csvAttachments and save report content.
        }

        private static bool IsIBKRCustomActivityReportSubject(string? subject, string subjectDate)
        {
            if (String.IsNullOrWhiteSpace(subject) || !subject.Contains(subjectDate, StringComparison.OrdinalIgnoreCase))
                return false;

            return subject.Contains("Custom Activity", StringComparison.OrdinalIgnoreCase) ||
                subject.Contains("Activity Statement", StringComparison.OrdinalIgnoreCase) ||
                subject.Contains("Activity Report", StringComparison.OrdinalIgnoreCase) ||
                subject.Contains("自定义活动报表", StringComparison.OrdinalIgnoreCase);
        }

        private Account GetIBKRAccount(string reportAccountId)
        {
            var normalizedReportAccountId = reportAccountId.Trim().ToUpperInvariant();
            var visibleDigits = Regex.Match(normalizedReportAccountId, @"\d+$").Value;
            if (String.IsNullOrWhiteSpace(visibleDigits))
                throw new InvalidOperationException($"Invalid IBKR account id: {reportAccountId}");

            var candidates = database.GetAllAccounts()
                .Where(account => account.name.StartsWith("IBKR_", StringComparison.OrdinalIgnoreCase))
                .Where(account =>
                {
                    var accountId = account.name["IBKR_".Length..].Trim().ToUpperInvariant();
                    if (normalizedReportAccountId.Contains('*'))
                        return accountId.EndsWith(visibleDigits, StringComparison.OrdinalIgnoreCase);
                    return accountId.Equals(normalizedReportAccountId, StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            if (candidates.Count == 1)
                return candidates[0];

            var usdCandidates = candidates.Where(account => account._v_t == CurrencyType.USD).ToList();
            if (usdCandidates.Count == 1)
                return usdCandidates[0];

            if (candidates.Count > 1)
            {
                throw new InvalidOperationException(
                    $"Find multiple IBKR accounts for {reportAccountId}: {String.Join(",", candidates.Select(account => account.name))}");
            }

            throw new InvalidOperationException($"Account not found: IBKR/{reportAccountId}");
        }

        private static DateTime? ParseIBKRReportDate(string text)
        {
            var normalizedText = NormalizeMailText(text);
            var labeledMatch = Regex.Match(
                normalizedText,
                @"(?:Report\s*Date|Statement\s*Date|Activity\s*Date|Date|报表日期|日期)\s*[:：]?\s*(?<date>\d{1,2}/\d{1,2}/\d{4}|[A-Z][a-z]+\s+\d{1,2},\s+\d{4})",
                RegexOptions.IgnoreCase);
            if (labeledMatch.Success && TryParseIBKRDate(labeledMatch.Groups["date"].Value, out var labeledDate))
                return labeledDate;

            var match = Regex.Match(normalizedText, @"\b(?<date>\d{1,2}/\d{1,2}/\d{4}|[A-Z][a-z]+\s+\d{1,2},\s+\d{4})\b", RegexOptions.IgnoreCase);
            if (match.Success && TryParseIBKRDate(match.Groups["date"].Value, out var date))
                return date;

            return null;
        }

        private static bool TryParseIBKRDate(string text, out DateTime date)
        {
            return DateTime.TryParseExact(
                text.Trim(),
                ["M/d/yyyy", "MM/dd/yyyy", "MMMM d, yyyy", "MMM d, yyyy"],
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out date);
        }

        private static string? ParseIBKRAccountId(string text)
        {
            var normalizedText = NormalizeMailText(text);
            var labeledMatch = Regex.Match(
                normalizedText,
                @"(?:Account(?:\s*ID|\s*Number)?|账号|账户)\s*[:：]?\s*(?<account>U(?:\*+|\d*)\d{4,})",
                RegexOptions.IgnoreCase);
            if (labeledMatch.Success)
                return labeledMatch.Groups["account"].Value.ToUpperInvariant();

            var match = Regex.Match(normalizedText, @"\b(?<account>U(?:\*{2,}\d{4,}|\d{5,}))\b", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups["account"].Value.ToUpperInvariant() : null;
        }

        private static string GetMessageText(MimeMessage message)
        {
            if (!String.IsNullOrWhiteSpace(message.TextBody))
                return WebUtility.HtmlDecode(message.TextBody);

            if (String.IsNullOrWhiteSpace(message.HtmlBody))
                return "";

            var doc = new HtmlDocument();
            doc.LoadHtml(message.HtmlBody);
            return WebUtility.HtmlDecode(doc.DocumentNode.InnerText);
        }

        private static string NormalizeMailText(string text)
        {
            return Regex.Replace(text, @"\s+", " ").Trim();
        }

        private static List<InMemoryCsvAttachment> ReadCsvAttachments(MimeMessage message)
        {
            var attachments = new List<InMemoryCsvAttachment>();
            foreach (var attachment in message.Attachments)
            {
                if (!IsCsvAttachment(attachment) || attachment is not MimePart mimePart)
                    continue;

                using var memory = new MemoryStream();
                mimePart.Content.DecodeTo(memory);
                attachments.Add(new InMemoryCsvAttachment(
                    mimePart.FileName ?? "report.csv",
                    memory.ToArray()));
            }

            return attachments;
        }

        private static bool IsCsvAttachment(MimeEntity attachment)
        {
            var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name ?? "";
            return fileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ||
                attachment.ContentType.MediaSubtype.Equals("csv", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<string> SearchBill(string sender, string subject,DateTime date)
        {
            var message = await SearchBill(sender, subject, date, null);
            return message?.HtmlBody ?? message?.TextBody ?? "";
        }

        public async Task<MimeMessage?> SearchBill(string sender, string subject, DateTime date, Func<MimeMessage, bool>? messageFilter)
        {
            try
            {
                using (MailKit.Net.Imap.ImapClient client = new())
                {
                    client.ProxyClient = proxy;
                    //client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    await client.ConnectAsync("imap.mail.yahoo.com", 993, true);
                    // 邮箱->security->generate app password
                    await client.AuthenticateAsync(username, apppasswd);
                    await client.Inbox.OpenAsync(FolderAccess.ReadOnly);
                    //搜索date所在月份的邮件
                    var query = SearchQuery.FromContains(sender)
                        .And(SearchQuery.SubjectContains(subject))
                        .And(SearchQuery.SentSince(date.AddDays(1-date.Day)))
                        .And(SearchQuery.SentBefore(date.AddDays(1-date.Day).AddMonths(1).AddSeconds(-1)));
                    var uids = await client.Inbox.SearchAsync(query);
                    if (uids.Count>1)
                        Console.WriteLine($"Find multiple bills {sender} {subject} {date}");
                    foreach (var uid in uids)
                    {
                        var message = await client.Inbox.GetMessageAsync(uid);
                        if (messageFilter is not null && !messageFilter(message))
                            continue;

                        return message;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail fail :{e.Message}");
            }
            return null;
        }

        private sealed record InMemoryCsvAttachment(string FileName, byte[] Content);
    }
}
