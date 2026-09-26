using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MailKit;
using MailKit.Search;
using MimeKit;
using MailKit.Net.Imap;
using System.Net;
using HtmlAgilityPack;

namespace MyBook
{
    partial class MailUtil
    {
        private static readonly string[] PayPalSenders = ["service@intl.paypal.com", "service@paypal.com"];
        internal async Task<List<CombinedUtil.PayPalMail>> ReadPayPalMailsAsync(
            Account account, DateTime since, CancellationToken cancellationToken)
        {
            var result = new List<CombinedUtil.PayPalMail>();
            try
            {
                if (String.IsNullOrWhiteSpace(account.email)) throw new InvalidOperationException();
                var mailbox = CreateMailboxForEmail("PayPal", account.email);
                cancellationToken.ThrowIfCancellationRequested();
                var query = SearchQuery.FromContains("paypal.com").And(SearchQuery.SentSince(since.Date));
                var messages = await RunWithMailSessionScope(() => SearchMessagesFromMailbox(mailbox, "PayPal", query,
                    IsPayPalSenderSummary, IsPayPalSender, GetMailDateTime,
                    allFolders: true, cancellationToken: cancellationToken)).ConfigureAwait(false);
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var message in messages)
                {
                    using (message)
                    {
                        var html = new HtmlDocument(); html.LoadHtml(message.HtmlBody ?? "");
                        var text = NormalizeMailText(WebUtility.HtmlDecode(message.TextBody ?? html.DocumentNode.InnerText));
                        var key = message.MessageId ?? CombinedUtil.PayPalHash(message.Subject + text);
                        if (seen.Add(key)) result.Add(new(account.Id, key, message.Date, message.Subject ?? "", text));
                    }
                }
            }
            catch (Exception e)
            {
                var cause = e.GetBaseException();
                var detail = cause is System.Net.Sockets.SocketException socket
                    ? socket.SocketErrorCode.ToString() : cause.GetType().Name;
                throw new MailParseException($"PayPal account {account.Id}: mail read failed; {detail}; protocol={(cause as ImapCommandException)?.Response.ToString() ?? "unavailable"}");
            }
            return result;
        }

        private static bool IsPayPalSenderSummary(IMessageSummary summary)
        {
            return PayPalSenders.Any(sender => SummaryIsFrom(summary, sender));
        }

        private static bool IsPayPalSender(MimeMessage message)
        {
            return PayPalSenders.Any(sender => IsFrom(message, sender));
        }

    }
}
