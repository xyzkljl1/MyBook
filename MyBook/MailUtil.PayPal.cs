using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MailKit.Search;
using MimeKit;

namespace MyBook
{
    // PayPal mail discovery only. My PayPal accounts currently do not have PayPal balance enabled
    // and are used only as a payment tool, so PayPal mails are not converted into Records.
    partial class MailUtil
    {
        private const string PayPalAccountPrefix = "PAYPAL";
        private static readonly string[] PayPalSenders = ["service@intl.paypal.com", "service@paypal.com"];
        private static readonly DateTime PayPalSearchStartDate = new(2024, 1, 1);

        public async Task FetchPayPalReports()
        {
            await FetchPayPalReports(PayPalSearchStartDate, DateTime.Today.AddDays(1));
        }

        public async Task FetchPayPalReports(DateTime date)
        {
            var range = GetMonthRange(date);
            await FetchPayPalReports(range.Since, range.Before);
        }

        private async Task FetchPayPalReports(DateTime since, DateTime before)
        {
            var accounts = GetPayPalAccounts();
            if (accounts.Count == 0)
            {
                Console.WriteLine("Fetch PayPal mails: no PayPal accounts");
                return;
            }

            foreach (var group in accounts
                         .GroupBy(account => account.email ?? "", StringComparer.OrdinalIgnoreCase)
                         .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
            {
                if (String.IsNullOrWhiteSpace(group.Key))
                    throw new InvalidOperationException($"Missing email for PayPal account {group.First().name}");

                var messages = await SearchPayPalMails(group.Key, since, before);
                Console.WriteLine($"Fetch PayPal mails {MaskEmail(group.Key)} {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}: count={messages.Count}");
                ParsePayPalMails(group.ToList(), messages);
            }
        }

        private List<Account> GetPayPalAccounts()
        {
            return database.GetAllAccounts()
                .Where(IsPayPalAccount)
                .OrderBy(account => account.name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsPayPalAccount(Account account)
        {
            return account.name.Equals(PayPalAccountPrefix, StringComparison.OrdinalIgnoreCase)
                || account.name.StartsWith($"{PayPalAccountPrefix}_", StringComparison.OrdinalIgnoreCase);
        }

        private async Task<List<MimeMessage>> SearchPayPalMails(string email, DateTime since, DateTime before)
        {
            var mailbox = CreateMailboxForEmail("PayPal", email);
            SearchQuery senderQuery = SearchQuery.FromContains(PayPalSenders[0]);
            foreach (var sender in PayPalSenders.Skip(1))
                senderQuery = senderQuery.Or(SearchQuery.FromContains(sender));

            var query = senderQuery
                .And(SearchQuery.SentSince(since.Date))
                .And(SearchQuery.SentBefore(before.Date));
            return await SearchMessagesFromMailbox(
                mailbox,
                $"PayPal {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}",
                query,
                IsPayPalSender,
                GetMailDateTime);
        }

        private static bool IsPayPalSender(MimeMessage message)
        {
            return PayPalSenders.Any(sender => IsFrom(message, sender));
        }

        private static void ParsePayPalMails(List<Account> accounts, List<MimeMessage> messages)
        {
            // Intentionally empty. These mails describe PayPal acting as a payment channel;
            // the real balance changes are imported from the linked card or bank statements.
        }
    }
}
