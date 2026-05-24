using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MailKit.Search;
using MimeKit;

namespace MyBook
{
    // PayPal mail discovery. PayPal accounts opt in by setting Account.email;
    // message parsing is intentionally left empty until the exact formats are known.
    partial class MailUtil
    {
        private const string PayPalAccountPrefix = "PAYPAL";
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

            foreach (var account in accounts)
            {
                var messages = await SearchPayPalMails(account, since, before);
                Console.WriteLine($"Fetch PayPal mails {account.name}: count={messages.Count}");
                ParsePayPalMails(account, messages);
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

        private async Task<List<MimeMessage>> SearchPayPalMails(Account account, DateTime since, DateTime before)
        {
            if (String.IsNullOrWhiteSpace(account.email))
                throw new InvalidOperationException($"Missing email for PayPal account {account.name}");

            var mailbox = CreateMailboxForEmail($"PayPal {account.name}", account.email);
            var query = SearchQuery.FromContains("paypal")
                .And(SearchQuery.SentSince(since.Date))
                .And(SearchQuery.SentBefore(before.Date));
            return await SearchMessagesFromMailbox(
                mailbox,
                $"PayPal {account.name} {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}",
                query,
                IsPayPalMail,
                GetMailDateTime);
        }

        private static bool IsPayPalMail(MimeMessage message)
        {
            return message.From.Mailboxes.Any(mailbox =>
                    mailbox.Address.Contains("paypal", StringComparison.OrdinalIgnoreCase))
                || message.Subject?.Contains("paypal", StringComparison.OrdinalIgnoreCase) == true;
        }

        private static void ParsePayPalMails(Account account, List<MimeMessage> messages)
        {
            foreach (var message in messages.OrderByDescending(GetMailDateTime).Take(3))
            {
                Console.WriteLine(
                    $"PayPal mail {account.name}: {GetMailDateTime(message):yyyy-MM-dd HH:mm:ss} {message.Subject}");
            }
        }
    }
}
