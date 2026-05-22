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
    // OCBC mail discovery. Parsing is intentionally left blank until concrete
    // OCBC mail formats are mapped.
    partial class MailUtil
    {
        private const string OCBCMailSender = "notifications@ocbc.com";
        private static readonly DateTime OCBCSearchStartDate = new(2024, 1, 1);

        public async Task FetchOCBCReports()
        {
            await FetchOCBCReports(OCBCSearchStartDate, DateTime.Today.AddDays(1));
        }

        public async Task FetchOCBCReports(DateTime date)
        {
            var reportMonth = FirstDayOfMonth(date);
            await FetchOCBCReports(reportMonth, reportMonth.AddMonths(1));
        }

        private async Task FetchOCBCReports(DateTime since, DateTime before)
        {
            var messages = await SearchOCBCMails(since, before);
            Console.WriteLine($"Fetch OCBC mails: {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}, count={messages.Count}");

            foreach (var message in messages)
                ParseOCBCMail(message);

            Console.WriteLine("Fetch OCBC mails done");
        }

        private async Task<List<MimeMessage>> SearchOCBCMails(DateTime since, DateTime before)
        {
            try
            {
                using ImapClient client = new();
                client.Timeout = MailClientTimeoutMilliseconds;
                client.CheckCertificateRevocation = false;
                client.ProxyClient = null;

                Console.WriteLine($"mail connect direct OCBC {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}");
                await RunMailOperation(token => client.ConnectAsync("imap.mail.yahoo.com", 993, true, token));
                Console.WriteLine("OCBC mail connected");
                await RunMailOperation(token => client.AuthenticateAsync(username, apppasswd, token));
                Console.WriteLine("OCBC mail authenticated");
                await RunMailOperation(token => client.Inbox.OpenAsync(FolderAccess.ReadOnly, token));
                Console.WriteLine("OCBC mail inbox opened");

                var query = SearchQuery.FromContains(OCBCMailSender)
                    .And(SearchQuery.SentSince(since.Date))
                    .And(SearchQuery.SentBefore(before.Date));
                var uids = await RunMailOperation(token => client.Inbox.SearchAsync(query, token));
                Console.WriteLine($"OCBC mail search found {uids.Count}");

                var messages = new List<MimeMessage>();
                foreach (var uid in uids)
                {
                    var message = await RunMailOperation(token => client.Inbox.GetMessageAsync(uid, token));
                    if (IsOCBCMail(message))
                        messages.Add(message);
                }

                return messages
                    .OrderBy(GetOCBCMailDateTime)
                    .ThenBy(message => message.Subject, StringComparer.Ordinal)
                    .ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch OCBC mail fail :{e.Message}");
                throw new InvalidOperationException($"Fetch OCBC mail failed: {e.Message}", e);
            }
        }

        private static bool IsOCBCMail(MimeMessage message)
        {
            return message.From.Mailboxes.Any(mailbox =>
                String.Equals(mailbox.Address, OCBCMailSender, StringComparison.OrdinalIgnoreCase));
        }

        private static void ParseOCBCMail(MimeMessage message)
        {
            Console.WriteLine($"Skip OCBC mail parse placeholder: {GetOCBCMailDateTime(message):yyyy-MM-dd HH:mm:ss} {message.Subject}");
        }

        private static DateTime GetOCBCMailDateTime(MimeMessage message)
        {
            var localTime = message.Date.LocalDateTime;
            return localTime == default ? GetMailDate(message) : localTime;
        }
    }
}
