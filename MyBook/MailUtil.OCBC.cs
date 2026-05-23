using System;
using System.Collections.Generic;
using System.Threading.Tasks;
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
            var range = GetMonthRange(date);
            await FetchOCBCReports(range.Since, range.Before);
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
            var query = SearchQuery.FromContains(OCBCMailSender)
                .And(SearchQuery.SentSince(since.Date))
                .And(SearchQuery.SentBefore(before.Date));
            return await SearchMessages(
                $"OCBC {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}",
                query,
                IsOCBCMail,
                GetMailDateTime);
        }

        private static bool IsOCBCMail(MimeMessage message)
        {
            return IsFrom(message, OCBCMailSender);
        }

        private static void ParseOCBCMail(MimeMessage message)
        {
            Console.WriteLine($"Skip OCBC mail parse placeholder: {GetMailDateTime(message):yyyy-MM-dd HH:mm:ss} {message.Subject}");
        }
    }
}
