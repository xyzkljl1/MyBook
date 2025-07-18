using System;
using System.Collections.Generic;
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
using static MailKit.Telemetry;

namespace MyBook
{
    // 从雅虎邮箱拉信用卡账单
    class MailUtil
    {
        IProxyClient proxy;
        string username;
        string apppasswd;
        public MailUtil(IConfigurationRoot config)
        {
            // 为了支持gbk编码
            System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            proxy = new Socks5Client(config["socksproxy"], Int32.Parse(config["socksproxy_port"]!));
            username = config["yahoo_user"]!;
            apppasswd = config["yahoo_pass"]!;
        }
        public async Task SearchICBCBill(DateTime date)
        {
            // 按月份搜索
            var text = await SearchBill("webmaster@icbc.com.cn", "中国工商银行客户对账单",date);
            if(!String.IsNullOrEmpty(text))
            {
                try
                {
                    var doc = new HtmlDocument();
                    doc.LoadHtml(text);
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
