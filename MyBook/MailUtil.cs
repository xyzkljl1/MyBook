using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Extensions.Configuration;
using MimeKit;

namespace MyBook
{
    // 邮件账单获取的共同依赖与通用 IMAP/解析辅助逻辑。
    partial class MailUtil
    {
        private const string InitialReportDirectoryName = "initialReports";
        private readonly string username;
        private readonly string apppasswd;
        private readonly DatabaseUtil database;
        private readonly IConfigurationRoot config;
        private readonly AsyncLocal<MailSessionScope?> currentMailSessionScope = new();
        private const int MailClientTimeoutMilliseconds = 300000;
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

        public async Task RunWithMailSessionScope(Func<Task> action)
        {
            await UseMailSessionScope(async _ =>
            {
                await action().ConfigureAwait(false);
                return true;
            }).ConfigureAwait(false);
        }

        private async Task<T> RunWithMailSessionScope<T>(Func<Task<T>> action)
        {
            return await UseMailSessionScope(_ => action()).ConfigureAwait(false);
        }

        private async Task<T> UseMailSessionScope<T>(Func<MailSessionScope, Task<T>> action)
        {
            if (currentMailSessionScope.Value is not null)
                return await action(currentMailSessionScope.Value).ConfigureAwait(false);

            await using var scope = new MailSessionScope();
            var previousScope = currentMailSessionScope.Value;
            currentMailSessionScope.Value = scope;
            try
            {
                return await action(scope).ConfigureAwait(false);
            }
            finally
            {
                currentMailSessionScope.Value = previousScope;
            }
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
            await RunWithMailSessionScope(async () =>
            {
                var month = GetNextMonthlyStatementDate(provider);
                var currentMonth = FirstDayOfMonth(DateTime.Today);
                while (month <= currentMonth)
                {
                    var imported = await fetchStatement(month).ConfigureAwait(false);
                    if (!imported)
                    {
                        if (DateTime.Today >= month.AddMonths(1))
                            throw new InvalidOperationException($"Missing {displayName} for {month:yyyy-MM}");
                        return;
                    }

                    month = month.AddMonths(1);
                }
            }).ConfigureAwait(false);
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

        private static string? FindInitialReportsDirectory()
        {
            foreach (var root in EnumerateInitialReportSearchRoots())
            {
                var directory = Path.Combine(root, InitialReportDirectoryName);
                if (Directory.Exists(directory))
                    return directory;
            }

            return null;
        }

        private static IEnumerable<string> EnumerateInitialReportSearchRoots()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var start in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory })
            {
                var directory = new DirectoryInfo(start);
                while (directory is not null)
                {
                    if (seen.Add(directory.FullName))
                        yield return directory.FullName;
                    directory = directory.Parent;
                }
            }
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

        private static DateTime GetMailDateTime(MailAttachmentMessage message)
        {
            return message.MailDateTime;
        }

        private static DateTime GetMailDate(MimeMessage message)
        {
            return GetMailDateTime(message).Date;
        }

        private static DateTime GetMailDate(MailAttachmentMessage message)
        {
            return GetMailDateTime(message).Date;
        }

        private static DateTime GetSummaryDateTime(IMessageSummary summary)
        {
            return summary.Envelope?.Date?.LocalDateTime
                ?? summary.InternalDate?.LocalDateTime
                ?? default;
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

        private static bool IsToOrCc(MimeMessage message, string recipient)
        {
            return message.To.Mailboxes.Concat(message.Cc.Mailboxes).Any(mailbox =>
                String.Equals(mailbox.Address, recipient, StringComparison.OrdinalIgnoreCase));
        }

        private string GetConfiguredYahooAddress()
        {
            var yahooUser = username.Trim();
            return yahooUser.Contains('@', StringComparison.Ordinal)
                ? yahooUser
                : $"{yahooUser}@yahoo.com";
        }

        private bool IsSelfSentYahooMessage(MimeMessage message)
        {
            var yahooAddress = GetConfiguredYahooAddress();
            return IsFrom(message, yahooAddress) && IsToOrCc(message, yahooAddress);
        }

        private bool IsSelfSentYahooSummary(IMessageSummary summary)
        {
            var yahooAddress = GetConfiguredYahooAddress();
            return SummaryIsFrom(summary, yahooAddress) && SummaryIsToOrCc(summary, yahooAddress);
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

        private static string GetAttachmentFileName(BodyPartBasic attachment)
        {
            return attachment.FileName ?? "";
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

        private static List<TAttachment> ReadMatchingAttachments<TAttachment>(
            MailAttachmentMessage message,
            Func<MailAttachment, string, TAttachment?> readAttachment)
            where TAttachment : class
        {
            var result = new List<TAttachment>();
            foreach (var attachment in message.Attachments)
            {
                var parsed = readAttachment(attachment, attachment.FileName);
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

        private static bool SummaryHasMatchingAttachment(
            IMessageSummary summary,
            Func<string, bool> predicate)
        {
            if (summary.Body is null)
                return true;

            var hasFileName = false;
            foreach (var part in GetSummaryAttachmentParts(summary))
            {
                hasFileName = true;
                if (predicate(GetAttachmentFileName(part)))
                    return true;
            }

            return !hasFileName;
        }

        private static IEnumerable<BodyPartBasic> GetSummaryAttachmentParts(IMessageSummary summary)
        {
            if (summary.Body is null)
                yield break;

            foreach (var part in summary.BodyParts.OfType<BodyPartBasic>())
            {
                if (!String.IsNullOrWhiteSpace(GetAttachmentFileName(part)))
                    yield return part;
            }
        }

        private static bool SummaryIsFrom(IMessageSummary summary, string sender)
        {
            var envelope = summary.Envelope;
            if (envelope is null)
                return true;

            var from = envelope.From?.Mailboxes ?? Enumerable.Empty<MailboxAddress>();
            return from.Any(mailbox =>
                String.Equals(mailbox.Address, sender, StringComparison.OrdinalIgnoreCase));
        }

        private static bool SummaryIsToOrCc(IMessageSummary summary, string recipient)
        {
            var envelope = summary.Envelope;
            if (envelope is null)
                return true;

            var to = envelope.To?.Mailboxes ?? Enumerable.Empty<MailboxAddress>();
            var cc = envelope.Cc?.Mailboxes ?? Enumerable.Empty<MailboxAddress>();
            return to.Concat(cc).Any(mailbox =>
                String.Equals(mailbox.Address, recipient, StringComparison.OrdinalIgnoreCase));
        }

        private static bool SummarySubjectEquals(IMessageSummary summary, string subject)
        {
            var summarySubject = summary.Envelope?.Subject;
            return summarySubject is null
                || String.Equals(summarySubject.Trim(), subject, StringComparison.Ordinal);
        }

        private async Task<MimeMessage?> SearchBill(string sender, string subject, DateTime date, Func<MimeMessage, bool>? messageFilter)
        {
            return (await SearchBills(sender, subject, date, messageFilter)).FirstOrDefault();
        }

        private async Task<List<MimeMessage>> SearchBills(string sender, string subject, DateTime date, Func<MimeMessage, bool>? messageFilter)
        {
            return await SearchBills(sender, subject, date, null, messageFilter).ConfigureAwait(false);
        }

        private async Task<List<MimeMessage>> SearchBills(
            string sender,
            string subject,
            DateTime date,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<MimeMessage, bool>? messageFilter)
        {
            var range = GetMonthRange(date);
            var query = SearchQuery.FromContains(sender)
                .And(SearchQuery.SubjectContains(subject))
                .And(SearchQuery.SentSince(range.Since))
                .And(SearchQuery.SentBefore(range.Before.AddSeconds(-1)));
            var messages = await SearchMessages($"{subject} {date:yyyy-MM-dd}", query, summaryFilter, messageFilter);
            if (messages.Count > 1)
                Console.WriteLine($"Find multiple bills {sender} {subject} {date}");

            return messages;
        }

        private async Task<List<MailAttachmentMessage>> SearchBillAttachments(
            string sender,
            string subject,
            DateTime date,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<string, bool> attachmentFileNameFilter)
        {
            var range = GetMonthRange(date);
            var query = SearchQuery.FromContains(sender)
                .And(SearchQuery.SubjectContains(subject))
                .And(SearchQuery.SentSince(range.Since))
                .And(SearchQuery.SentBefore(range.Before.AddSeconds(-1)));
            var messages = await SearchAttachmentMessages(
                $"{subject} {date:yyyy-MM-dd}",
                query,
                summaryFilter,
                attachmentFileNameFilter,
                GetMailDateTime).ConfigureAwait(false);
            if (messages.Count > 1)
                Console.WriteLine($"Find multiple bills {sender} {subject} {date}");

            return messages;
        }

        private async Task<List<MimeMessage>> SearchSupplementalStatementMails(
            string label,
            string subject,
            DateTime statementMonth,
            Func<MimeMessage, bool>? messageFilter,
            Func<IMessageSummary, bool>? summaryFilter = null,
            Func<MimeMessage, DateTime>? orderDateSelector = null)
        {
            if (!ShouldSearchSupplementalStatementMail(statementMonth))
                return [];

            var yahooAddress = GetConfiguredYahooAddress();
            var query = SearchQuery.FromContains(yahooAddress)
                .And(SearchQuery.SubjectContains(subject))
                .And(SearchQuery.SentSince(statementMonth.Date));
            return await SearchMessages(
                $"{label} supplemental from {statementMonth:yyyy-MM-dd}",
                query,
                summary => IsSelfSentYahooSummary(summary) && (summaryFilter?.Invoke(summary) ?? true),
                message => IsSelfSentYahooMessage(message) && (messageFilter?.Invoke(message) ?? true),
                orderDateSelector);
        }

        private async Task<List<MailAttachmentMessage>> SearchSupplementalStatementAttachmentMails(
            string label,
            string subject,
            DateTime statementMonth,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<string, bool> attachmentFileNameFilter,
            Func<MailAttachmentMessage, DateTime>? orderDateSelector = null)
        {
            if (!ShouldSearchSupplementalStatementMail(statementMonth))
                return [];

            var yahooAddress = GetConfiguredYahooAddress();
            var query = SearchQuery.FromContains(yahooAddress)
                .And(SearchQuery.SubjectContains(subject))
                .And(SearchQuery.SentSince(statementMonth.Date));
            return await SearchAttachmentMessages(
                $"{label} supplemental from {statementMonth:yyyy-MM-dd}",
                query,
                summary => IsSelfSentYahooSummary(summary) && (summaryFilter?.Invoke(summary) ?? true),
                attachmentFileNameFilter,
                orderDateSelector).ConfigureAwait(false);
        }

        private static bool ShouldSearchSupplementalStatementMail(DateTime statementMonth)
        {
            var currentMonth = FirstDayOfMonth(DateTime.Today);
            return FirstDayOfMonth(statementMonth).AddMonths(2) <= currentMonth;
        }

        private async Task<List<MimeMessage>> SearchMessages(
            string label,
            SearchQuery query,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector = null)
        {
            return await SearchMessages(label, query, null, messageFilter, orderDateSelector).ConfigureAwait(false);
        }

        private async Task<List<MimeMessage>> SearchMessages(
            string label,
            SearchQuery query,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector = null)
        {
            try
            {
                return await SearchMessagesCore(CreateYahooMailbox(), label, query, summaryFilter, messageFilter, orderDateSelector);
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
            return await SearchMessagesFromMailbox(mailbox, label, query, null, messageFilter, orderDateSelector).ConfigureAwait(false);
        }

        private async Task<List<MimeMessage>> SearchMessagesFromMailbox(
            ImapMailbox mailbox,
            string label,
            SearchQuery query,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector = null)
        {
            try
            {
                return await SearchMessagesCore(mailbox, label, query, summaryFilter, messageFilter, orderDateSelector);
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
            Func<IMessageSummary, bool>? summaryFilter,
            Func<MimeMessage, bool>? messageFilter,
            Func<MimeMessage, DateTime>? orderDateSelector)
        {
            var uids = await UseMailFolderAsync(
                mailbox,
                label,
                folder => RunMailOperation(token => folder.SearchAsync(query, token))).ConfigureAwait(false);
            Console.WriteLine($"mail search {label} found {uids.Count}");
            var candidateUids = uids.AsEnumerable();
            if (summaryFilter is not null && uids.Count > 0)
            {
                var summaries = await UseMailFolderAsync(
                    mailbox,
                    $"{label} summaries",
                    folder => RunMailOperation(token => folder.FetchAsync(
                        uids,
                        MessageSummaryItems.Envelope
                            | MessageSummaryItems.BodyStructure
                            | MessageSummaryItems.UniqueId
                            | MessageSummaryItems.InternalDate,
                        token))).ConfigureAwait(false);
                var filteredSummaries = summaries
                    .Where(summaryFilter)
                    .ToList();
                Console.WriteLine($"mail summary filter {label} kept {filteredSummaries.Count}/{summaries.Count}");
                candidateUids = filteredSummaries.Select(summary => summary.UniqueId);
            }

            var messages = new List<MimeMessage>();
            foreach (var uid in candidateUids)
            {
                var message = await UseMailFolderAsync(
                    mailbox,
                    $"{label} uid={uid.Id}",
                    folder => RunMailOperation(token => folder.GetMessageAsync(uid, token))).ConfigureAwait(false);
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

        private async Task<List<MailAttachmentMessage>> SearchAttachmentMessages(
            string label,
            SearchQuery query,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<string, bool> attachmentFileNameFilter,
            Func<MailAttachmentMessage, DateTime>? orderDateSelector = null)
        {
            try
            {
                return await SearchAttachmentMessagesCore(
                    CreateYahooMailbox(),
                    label,
                    query,
                    summaryFilter,
                    attachmentFileNameFilter,
                    orderDateSelector).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail attachments fail {label}: {e.Message}");
                throw new InvalidOperationException($"Fetch mail attachments failed: {label}: {e.Message}", e);
            }
        }

        private async Task<List<MailAttachmentMessage>> SearchAttachmentMessagesFromMailbox(
            ImapMailbox mailbox,
            string label,
            SearchQuery query,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<string, bool> attachmentFileNameFilter,
            Func<MailAttachmentMessage, DateTime>? orderDateSelector = null)
        {
            try
            {
                return await SearchAttachmentMessagesCore(
                    mailbox,
                    label,
                    query,
                    summaryFilter,
                    attachmentFileNameFilter,
                    orderDateSelector).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch mail attachments fail {label}: {e.Message}");
                throw new InvalidOperationException($"Fetch mail attachments failed: {label}: {e.Message}", e);
            }
        }

        private async Task<List<MailAttachmentMessage>> SearchAttachmentMessagesCore(
            ImapMailbox mailbox,
            string label,
            SearchQuery query,
            Func<IMessageSummary, bool>? summaryFilter,
            Func<string, bool> attachmentFileNameFilter,
            Func<MailAttachmentMessage, DateTime>? orderDateSelector)
        {
            var uids = await UseMailFolderAsync(
                mailbox,
                label,
                folder => RunMailOperation(token => folder.SearchAsync(query, token))).ConfigureAwait(false);
            Console.WriteLine($"mail search {label} found {uids.Count}");
            if (uids.Count == 0)
                return [];

            var summaries = await UseMailFolderAsync(
                mailbox,
                $"{label} summaries",
                folder => RunMailOperation(token => folder.FetchAsync(
                    uids,
                    MessageSummaryItems.Envelope
                        | MessageSummaryItems.BodyStructure
                        | MessageSummaryItems.UniqueId
                        | MessageSummaryItems.InternalDate,
                    token))).ConfigureAwait(false);
            var filteredSummaries = summaries
                .Where(summary => summaryFilter?.Invoke(summary) ?? true)
                .ToList();
            if (summaryFilter is not null)
                Console.WriteLine($"mail summary filter {label} kept {filteredSummaries.Count}/{summaries.Count}");

            var attachmentMessages = new List<MailAttachmentMessage>();
            foreach (var summary in filteredSummaries)
            {
                var matchingParts = GetSummaryAttachmentParts(summary)
                    .Where(part => attachmentFileNameFilter(GetAttachmentFileName(part)))
                    .ToList();
                if (matchingParts.Count == 0)
                    continue;

                var attachments = new List<MailAttachment>();
                foreach (var part in matchingParts)
                {
                    var entity = await UseMailFolderAsync(
                        mailbox,
                        $"{label} uid={summary.UniqueId.Id} attachment={GetAttachmentFileName(part)}",
                        folder => RunMailOperation(token => folder.GetBodyPartAsync(summary.UniqueId, part, token))).ConfigureAwait(false);
                    if (entity is not MimePart mimePart)
                        continue;

                    var fileName = GetAttachmentFileName(mimePart);
                    if (String.IsNullOrWhiteSpace(fileName))
                        fileName = GetAttachmentFileName(part);
                    if (!attachmentFileNameFilter(fileName))
                        continue;

                    attachments.Add(new MailAttachment(fileName, ReadMimePartBytes(mimePart)));
                }

                if (attachments.Count == 0)
                    continue;

                attachmentMessages.Add(new MailAttachmentMessage(
                    summary.UniqueId.Id,
                    summary.Envelope?.Subject ?? "",
                    GetSummaryDateTime(summary),
                    attachments));
            }

            Console.WriteLine($"mail attachment filter {label} kept {attachmentMessages.Count}/{filteredSummaries.Count}");
            orderDateSelector ??= GetMailDateTime;
            return attachmentMessages
                .OrderBy(orderDateSelector)
                .ThenBy(message => message.Subject, StringComparer.Ordinal)
                .ThenBy(message => message.UniqueId)
                .ToList();
        }

        private async Task<T> UseMailFolderAsync<T>(
            ImapMailbox mailbox,
            string label,
            Func<IMailFolder, Task<T>> action)
        {
            return await UseMailSessionScope(scope =>
                scope.UseFolderAsync(mailbox, label, action)).ConfigureAwait(false);
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

        private sealed class MailSessionScope : IAsyncDisposable
        {
            private readonly Dictionary<ImapMailbox, MailSessionConnection> connections = new();

            public async Task<T> UseFolderAsync<T>(
                ImapMailbox mailbox,
                string label,
                Func<IMailFolder, Task<T>> action)
            {
                var connection = GetConnection(mailbox);
                for (var attempt = 0; ; attempt++)
                {
                    try
                    {
                        var folder = await EnsureOpenFolderAsync(connection, label).ConfigureAwait(false);
                        return await action(folder).ConfigureAwait(false);
                    }
                    catch (Exception exception) when (attempt == 0 && IsMailConnectionException(exception))
                    {
                        Console.WriteLine($"mail reconnect {mailbox.Label} {label}: {exception.Message}");
                        await ResetConnectionAsync(connection).ConfigureAwait(false);
                    }
                }
            }

            public async ValueTask DisposeAsync()
            {
                foreach (var connection in connections.Values)
                    await ResetConnectionAsync(connection).ConfigureAwait(false);
                connections.Clear();
            }

            private MailSessionConnection GetConnection(ImapMailbox mailbox)
            {
                if (!connections.TryGetValue(mailbox, out var connection))
                {
                    connection = new MailSessionConnection(mailbox);
                    connections.Add(mailbox, connection);
                }

                return connection;
            }

            private static async Task<IMailFolder> EnsureOpenFolderAsync(MailSessionConnection connection, string label)
            {
                if (connection.Client is not null
                    && connection.Client.IsConnected
                    && connection.Client.IsAuthenticated
                    && connection.Folder is not null
                    && connection.Folder.IsOpen)
                {
                    return connection.Folder;
                }

                await ResetConnectionAsync(connection).ConfigureAwait(false);
                var mailbox = connection.Mailbox;
                var client = new ImapClient
                {
                    Timeout = MailClientTimeoutMilliseconds,
                    CheckCertificateRevocation = false,
                    ProxyClient = null
                };

                try
                {
                    Console.WriteLine($"mail connect direct {mailbox.Label} {label}");
                    await RunMailOperation(token => client.ConnectAsync(mailbox.Host, mailbox.Port, mailbox.UseSsl, token))
                        .ConfigureAwait(false);
                    Console.WriteLine("mail connected");
                    await RunMailOperation(token => client.AuthenticateAsync(mailbox.Username, mailbox.Password, token))
                        .ConfigureAwait(false);
                    Console.WriteLine("mail authenticated");
                    var folder = await GetMailSearchFolder(client, mailbox).ConfigureAwait(false);
                    await RunMailOperation(token => folder.OpenAsync(FolderAccess.ReadOnly, token))
                        .ConfigureAwait(false);
                    Console.WriteLine($"mail folder opened {folder.FullName}");
                    connection.Client = client;
                    connection.Folder = folder;
                    return folder;
                }
                catch
                {
                    client.Dispose();
                    throw;
                }
            }

            private static async Task ResetConnectionAsync(MailSessionConnection connection)
            {
                var client = connection.Client;
                connection.Folder = null;
                connection.Client = null;
                if (client is null)
                    return;

                try
                {
                    if (client.IsConnected)
                        await client.DisconnectAsync(true).WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
                }
                catch
                {
                    // Broken IMAP sessions are discarded and recreated by the next operation.
                }
                finally
                {
                    client.Dispose();
                }
            }

            private static bool IsMailConnectionException(Exception exception)
            {
                return exception is IOException
                    or SocketException
                    or TimeoutException
                    or OperationCanceledException
                    or SslHandshakeException
                    or ServiceNotConnectedException
                    or ServiceNotAuthenticatedException
                    or ProtocolException
                    || exception.InnerException is not null && IsMailConnectionException(exception.InnerException);
            }
        }

        private sealed class MailSessionConnection(ImapMailbox mailbox)
        {
            public ImapMailbox Mailbox { get; } = mailbox;
            public ImapClient? Client { get; set; }
            public IMailFolder? Folder { get; set; }
        }

        private sealed record MailAttachmentMessage(
            uint UniqueId,
            string Subject,
            DateTime MailDateTime,
            IReadOnlyList<MailAttachment> Attachments);

        private sealed record MailAttachment(string FileName, byte[] Content);
    }
}
