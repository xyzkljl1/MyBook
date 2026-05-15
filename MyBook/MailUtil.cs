using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security.Authentication;
using System.Text;
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
        public Account? FindICBCAccount(string name, CurrencyType currencyType)
        {
            return database.GetOrAddAccount("ICBC", name.Substring(0, 4), currencyType.ToString());
        }
        // 工行对账单，所有信用卡视作同一账户
        public async Task SearchICBCBill(DateTime date)
        {
            // 按月份搜索
            var billText = await SearchBill("webmaster@icbc.com.cn", "中国工商银行客户对账单", date);
            var monthText = date.ToString("yyyy-MM");
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
                        var account = FindICBCAccount(line[0], balance.t);
                        if (account is null)
                            throw new MailParseException("Parse ICBC Bill Fail, Invalid Account");
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
                            record.Account = FindICBCAccount(line[0], CurrencyType.RMB);
                            if (record.Account is null)
                                throw new MailParseException("Parse ICBC Bill Fail, Invalid Account");
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

                    database.SaveRecords(records);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"parse mail fail :{e.Message}");
                }
            }
        }
        public async Task<string> SearchBill(string sender, string subject,DateTime date)
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
                        var message = client.Inbox.GetMessage(uid);
                        return message.HtmlBody;// ?? message.TextBody;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail fail :{e.Message}");
            }
            return "";
        }
    }
}
