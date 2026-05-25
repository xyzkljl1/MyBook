using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using MimeKit;

namespace MyBook
{
    // 邮件账单获取的共同依赖与通用 IMAP/解析辅助逻辑。
    partial class MailUtil
    {
        private readonly string username;
        private readonly string apppasswd;
        private readonly DatabaseUtil database;
        private readonly IConfigurationRoot config;
        private const int MailClientTimeoutMilliseconds = 30000;
        private static readonly TimeSpan MailClientTimeout = TimeSpan.FromMilliseconds(MailClientTimeoutMilliseconds);
        private sealed record ImapMailbox(string Label, string Host, int Port, bool UseSsl, string Username, string Password, bool UseAllMail);

        public MailUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            // 为了支持 gbk 编码。
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            username = config["yahoo_user"]!;
            apppasswd = config["yahoo_pass"]!;
            this.database = database;
            this.config = config;
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
            var latestKey = database.GetLatestStatementImportKey(provider);
            if (DateTime.TryParseExact(latestKey, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out var latestKeyDate))
                return FirstDayOfMonth(latestKeyDate).AddMonths(1);

            var latestTime = database.GetLatestStatementImportTime(provider);
            if (latestTime is null)
                throw new InvalidOperationException($"Missing statement import checkpoint for {provider}");

            return FirstDayOfMonth(latestTime.Value).AddMonths(1);
        }

        private async Task FetchMonthlyStatements(
            StatementImportProvider provider,
            string displayName,
            Func<DateTime, Task<bool>> fetchStatement)
        {
            var month = GetNextMonthlyStatementDate(provider);
            var currentMonth = FirstDayOfMonth(DateTime.Today);
            while (month <= currentMonth)
            {
                var imported = await fetchStatement(month);
                if (!imported)
                {
                    if (DateTime.Today >= month.AddMonths(1))
                        throw new InvalidOperationException($"Missing {displayName} for {month:yyyy-MM}");
                    return;
                }

                month = month.AddMonths(1);
            }
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

        private static (DateTime Since, DateTime Before) GetMonthRange(DateTime date)
        {
            var since = FirstDayOfMonth(date);
            return (since, since.AddMonths(1));
        }

        private static DateTime GetMailDateTime(MimeMessage message)
        {
            var localTime = message.Date.LocalDateTime;
            return localTime == default ? default : localTime;
        }

        private static DateTime GetMailDate(MimeMessage message)
        {
            return GetMailDateTime(message).Date;
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

        private static bool IsFrom(MimeMessage message, string sender)
        {
            return message.From.Mailboxes.Any(mailbox =>
                String.Equals(mailbox.Address, sender, StringComparison.OrdinalIgnoreCase));
        }

        private static byte[] ReadMimePartBytes(MimePart mimePart)
        {
            using var memory = new MemoryStream();
            mimePart.Content.DecodeTo(memory);
            return memory.ToArray();
        }

        private static string GetAttachmentFileName(MimeEntity attachment)
        {
            return attachment.ContentDisposition?.FileName
                ?? attachment.ContentType.Name
                ?? "";
        }

        private static List<TAttachment> ReadMatchingAttachments<TAttachment>(
            MimeMessage message,
            Func<MimePart, string, TAttachment?> readAttachment)
            where TAttachment : class
        {
            var result = new List<TAttachment>();
            foreach (var attachment in message.Attachments.OfType<MimePart>())
            {
                var fileName = GetAttachmentFileName(attachment);
                var parsed = readAttachment(attachment, fileName);
                if (parsed is not null)
                    result.Add(parsed);
            }

            return result;
        }

        private static bool HasMatchingAttachment(
            MimeMessage message,
            Func<MimePart, string, bool> predicate)
        {
            return message.Attachments
                .OfType<MimePart>()
                .Any(attachment => predicate(attachment, GetAttachmentFileName(attachment)));
        }

        private async Task<MimeMessage?> SearchBill(string sender, string subject, DateTime date, Func<MimeMessage, bool>? messageFilter)
        {
            return (await SearchBills(sender, subject, date, messageFilter)).FirstOrDefault();
        }

        private async Task<List<MimeMessage>> SearchBills(string sender, string subject, DateTime date, Func<MimeMessage, bool>? messageFilter)
        {
            var range = GetMonthRange(date);
            var query = SearchQuery.FromContains(sender)
                .And(SearchQuery.SubjectContains(subject))
                .And(SearchQuery.SentSince(range.Since))
                .And(SearchQuery.SentBefore(range.Before.AddSeconds(-1)));
            var messages = await SearchMessages($"{subject} {date:yyyy-MM-dd}", query, messageFilter);
            if (messages.Count > 1)
                Console.WriteLine($"Find multiple bills {sender} {subject} {date}");

            return messages;
        }

        private async Task<List<MimeMessage>> SearchMessages(
            string label,
            SearchQuery query,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector = null)
        {
            try
            {
                return await SearchMessagesCore(CreateYahooMailbox(), label, query, messageFilter, orderDateSelector);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail fail {label}: {e.Message}");
                throw new InvalidOperationException($"Fetch mail failed: {label}: {e.Message}", e);
            }
        }

        private async Task<List<MimeMessage>> SearchMessagesFromMailbox(
            ImapMailbox mailbox,
            string label,
            SearchQuery query,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector = null)
        {
            try
            {
                return await SearchMessagesCore(mailbox, label, query, messageFilter, orderDateSelector);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail fail {label}: {e.Message}");
                throw new InvalidOperationException($"Fetch mail failed: {label}: {e.Message}", e);
            }
        }

        private async Task<List<MimeMessage>> SearchMessagesCore(
            ImapMailbox mailbox,
            string label,
            SearchQuery query,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector)
        {
            using MailKit.Net.Imap.ImapClient client = new();
            client.Timeout = MailClientTimeoutMilliseconds;
            client.CheckCertificateRevocation = false;
            client.ProxyClient = null;

            //client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            Console.WriteLine($"mail connect direct {mailbox.Label} {label}");
            await RunMailOperation(token => client.ConnectAsync(mailbox.Host, mailbox.Port, mailbox.UseSsl, token));
            Console.WriteLine("mail connected");
            await RunMailOperation(token => client.AuthenticateAsync(mailbox.Username, mailbox.Password, token));
            Console.WriteLine("mail authenticated");
            var folder = await GetMailSearchFolder(client, mailbox);
            await RunMailOperation(token => folder.OpenAsync(FolderAccess.ReadOnly, token));
            Console.WriteLine($"mail folder opened {folder.FullName}");

            var uids = await RunMailOperation(token => folder.SearchAsync(query, token));
            Console.WriteLine($"mail search {label} found {uids.Count}");
            var messages = new List<MimeMessage>();
            foreach (var uid in uids)
            {
                var message = await RunMailOperation(token => folder.GetMessageAsync(uid, token));
                if (messageFilter is not null && !messageFilter(message))
                    continue;

                messages.Add(message);
            }

            orderDateSelector ??= GetMailDate;
            return messages
                .OrderBy(orderDateSelector)
                .ThenBy(message => message.Subject, StringComparer.Ordinal)
                .ToList();
        }

        private ImapMailbox CreateYahooMailbox()
        {
            return new ImapMailbox("Yahoo", "imap.mail.yahoo.com", 993, true, username, apppasswd, false);
        }

        private ImapMailbox CreateMailboxForEmail(string label, string email)
        {
            if (IsConfiguredMailboxEmail(config["gmail_user"], email, "gmail.com"))
                return CreateGmailMailbox(label, email);

            if (IsConfiguredMailboxEmail(config["yahoo_user"], email, "yahoo.com"))
                return new ImapMailbox($"Yahoo {MaskEmail(email)}", "imap.mail.yahoo.com", 993, true, username, apppasswd, false);

            throw new InvalidOperationException($"Configured mail account mismatch for {label}: account email={MaskEmail(email)}");
        }

        private ImapMailbox CreateGmailMailbox(string label, string email)
        {
            var gmailUser = config["gmail_user"];
            var gmailPassword = config["gmail_app_pwd"];
            if (String.IsNullOrWhiteSpace(gmailUser) || String.IsNullOrWhiteSpace(gmailPassword))
                throw new InvalidOperationException("Missing gmail_user or gmail_app_pwd in config.json");
            if (!IsConfiguredMailboxEmail(gmailUser, email, "gmail.com"))
                throw new InvalidOperationException($"Configured Gmail account mismatch for {label}: account email={MaskEmail(email)}, gmail_user={MaskEmail(gmailUser)}");

            return new ImapMailbox($"Gmail {MaskEmail(email)}", "imap.gmail.com", 993, true, gmailUser, gmailPassword, true);
        }

        private static bool IsConfiguredMailboxEmail(string? configuredUser, string email, string defaultDomain)
        {
            if (String.IsNullOrWhiteSpace(configuredUser))
                return false;
            if (String.Equals(configuredUser, email, StringComparison.OrdinalIgnoreCase))
                return true;
            if (configuredUser.Contains('@', StringComparison.Ordinal))
                return false;

            return String.Equals($"{configuredUser}@{defaultDomain}", email, StringComparison.OrdinalIgnoreCase);
        }

        private static async Task<IMailFolder> GetMailSearchFolder(ImapClient client, ImapMailbox mailbox)
        {
            if (!mailbox.UseAllMail)
                return client.Inbox;

            try
            {
                return client.GetFolder(SpecialFolder.All);
            }
            catch
            {
                foreach (var folder in await ListMailFolders(client))
                {
                    if (folder.Attributes.HasFlag(FolderAttributes.All)
                        || String.Equals(folder.Name, "All Mail", StringComparison.OrdinalIgnoreCase)
                        || String.Equals(folder.Name, "所有邮件", StringComparison.OrdinalIgnoreCase))
                        return folder;
                }

                return client.Inbox;
            }
        }

        private static async Task<List<IMailFolder>> ListMailFolders(ImapClient client)
        {
            var result = new List<IMailFolder>();
            foreach (var ns in client.PersonalNamespaces)
            {
                var root = client.GetFolder(ns);
                await AddMailFolders(root, result);
            }

            return result;
        }

        private static async Task AddMailFolders(IMailFolder folder, List<IMailFolder> result)
        {
            result.Add(folder);
            foreach (var child in await folder.GetSubfoldersAsync(false))
                await AddMailFolders(child, result);
        }

        private static string MaskEmail(string? email)
        {
            if (String.IsNullOrWhiteSpace(email))
                return "";

            var at = email.IndexOf('@');
            if (at <= 1)
                return "***";

            return $"{email[0]}***{email[(at - 1)..]}";
        }

        private static async Task RunMailOperation(Func<CancellationToken, Task> operation)
        {
            using var cancellation = new CancellationTokenSource(MailClientTimeoutMilliseconds);
            await operation(cancellation.Token).WaitAsync(MailClientTimeout);
        }

        private static async Task<T> RunMailOperation<T>(Func<CancellationToken, Task<T>> operation)
        {
            using var cancellation = new CancellationTokenSource(MailClientTimeoutMilliseconds);
            return await operation(cancellation.Token).WaitAsync(MailClientTimeout);
        }
    }
}
