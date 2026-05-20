using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Proxy;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using MimeKit;

namespace MyBook
{
    // 邮件账单获取的共同依赖与通用 IMAP/解析辅助逻辑。
    partial class MailUtil
    {
        private readonly IProxyClient proxy;
        private readonly string username;
        private readonly string apppasswd;
        private readonly DatabaseUtil database;

        public MailUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            // 为了支持 gbk 编码。
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            proxy = new Socks5Client(config["socksproxy"], Int32.Parse(config["socksproxy_port"]!));
            username = config["yahoo_user"]!;
            apppasswd = config["yahoo_pass"]!;
            this.database = database;
        }

        private static bool IsSameAccount(Account? left, Account? right)
        {
            if (left is null || right is null)
                return false;
            if (left.Id > 0 && right.Id > 0)
                return left.Id == right.Id;

            return left.name == right.name;
        }

        private DateTime GetNextMonthlyStatementDate(StatementImportProvider provider)
        {
            var latestTime = database.GetLatestStatementImportTime(provider);
            if (latestTime is null)
                throw new InvalidOperationException($"Missing statement import checkpoint for {provider}");

            return FirstDayOfMonth(latestTime.Value).AddMonths(1);
        }

        private DateTime GetNextDailyStatementDate(StatementImportProvider provider)
        {
            var latestTime = database.GetLatestStatementImportTime(provider);
            if (latestTime is null)
                throw new InvalidOperationException($"Missing statement import checkpoint for {provider}");

            return latestTime.Value.Date.AddDays(1);
        }

        private static DateTime FirstDayOfMonth(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
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

        private async Task<string> SearchBill(string sender, string subject, DateTime date)
        {
            var message = await SearchBill(sender, subject, date, null);
            return message?.HtmlBody ?? message?.TextBody ?? "";
        }

        private async Task<MimeMessage?> SearchBill(string sender, string subject, DateTime date, Func<MimeMessage, bool>? messageFilter)
        {
            try
            {
                using MailKit.Net.Imap.ImapClient client = new();
                client.ProxyClient = proxy;
                //client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                await client.ConnectAsync("imap.mail.yahoo.com", 993, true);
                // 邮箱->security->generate app password。
                await client.AuthenticateAsync(username, apppasswd);
                await client.Inbox.OpenAsync(FolderAccess.ReadOnly);

                // 搜索 date 所在月份的邮件。
                var query = SearchQuery.FromContains(sender)
                    .And(SearchQuery.SubjectContains(subject))
                    .And(SearchQuery.SentSince(date.AddDays(1 - date.Day)))
                    .And(SearchQuery.SentBefore(date.AddDays(1 - date.Day).AddMonths(1).AddSeconds(-1)));
                var uids = await client.Inbox.SearchAsync(query);
                if (uids.Count > 1)
                    Console.WriteLine($"Find multiple bills {sender} {subject} {date}");
                foreach (var uid in uids)
                {
                    var message = await client.Inbox.GetMessageAsync(uid);
                    if (messageFilter is not null && !messageFilter(message))
                        continue;

                    return message;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail fail :{e.Message}");
            }
            return null;
        }
    }
}
