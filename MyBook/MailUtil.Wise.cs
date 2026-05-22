using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;

namespace MyBook
{
    // Wise 邮件获取入口；当前只负责从邮箱定位邮件，解析和入库逻辑后续补充。
    partial class MailUtil
    {
        private const string WiseMailSender = "noreply@wise.com";
        private const int WiseMailClientTimeoutMilliseconds = 30000;
        private static readonly TimeSpan WiseMailClientTimeout = TimeSpan.FromMilliseconds(WiseMailClientTimeoutMilliseconds);

        public async Task FetchWiseReports()
        {
            await FetchWiseReports(DateTime.Today);
        }

        public async Task FetchWiseReports(DateTime date)
        {
            var reportMonth = FirstDayOfMonth(date);
            var messages = await SearchWiseMails(reportMonth);
            Console.WriteLine($"Fetch Wise mails: {reportMonth:yyyy-MM}, count={messages.Count}");
            foreach (var message in messages)
                ParseWiseMail(message);
        }

        private async Task<List<MimeMessage>> SearchWiseMails(DateTime date)
        {
            try
            {
                using ImapClient client = new();
                client.Timeout = WiseMailClientTimeoutMilliseconds;
                client.CheckCertificateRevocation = false;
                client.ProxyClient = null;

                Console.WriteLine($"mail connect direct Wise {date:yyyy-MM-dd}");
                await RunWiseMailOperation(token => client.ConnectAsync("imap.mail.yahoo.com", 993, true, token));
                await RunWiseMailOperation(token => client.AuthenticateAsync(username, apppasswd, token));
                await RunWiseMailOperation(token => client.Inbox.OpenAsync(FolderAccess.ReadOnly, token));

                var query = SearchQuery.FromContains(WiseMailSender)
                    .And(SearchQuery.SentSince(date.AddDays(1 - date.Day)))
                    .And(SearchQuery.SentBefore(date.AddDays(1 - date.Day).AddMonths(1).AddSeconds(-1)));
                var uids = await RunWiseMailOperation(token => client.Inbox.SearchAsync(query, token));
                var messages = new List<MimeMessage>();
                foreach (var uid in uids)
                {
                    var message = await RunWiseMailOperation(token => client.Inbox.GetMessageAsync(uid, token));
                    if (IsWiseMail(message))
                        messages.Add(message);
                }

                return messages.OrderBy(GetMailDate).ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch Wise mail fail :{e.Message}");
                throw new InvalidOperationException($"Fetch Wise mail failed: {e.Message}", e);
            }
        }

        private static bool IsWiseMail(MimeMessage message)
        {
            return message.From.Mailboxes.Any(mailbox =>
                String.Equals(mailbox.Address, WiseMailSender, StringComparison.OrdinalIgnoreCase));
        }

        private static void ParseWiseMail(MimeMessage message)
        {
            _ = message;
        }

        private static async Task RunWiseMailOperation(Func<CancellationToken, Task> operation)
        {
            using var cancellation = new CancellationTokenSource(WiseMailClientTimeoutMilliseconds);
            await operation(cancellation.Token).WaitAsync(WiseMailClientTimeout);
        }

        private static async Task<T> RunWiseMailOperation<T>(Func<CancellationToken, Task<T>> operation)
        {
            using var cancellation = new CancellationTokenSource(WiseMailClientTimeoutMilliseconds);
            return await operation(cancellation.Token).WaitAsync(WiseMailClientTimeout);
        }
    }
}
