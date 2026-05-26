using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using UglyToad.PdfPig;

namespace MyBook
{
    // ICBC history-detail PDF mail discovery and parsing.
    // These statements can arrive out of chronological order and their covered periods may overlap.
    partial class MailUtil
    {
        private const StatementImportProvider ICBCHistoryDetailProvider = StatementImportProvider.ICBCHistoryDetailMail;
        private const string ICBCHistoryDetailSender = "webmaster@icbc.com.cn";
        private const string ICBCHistoryDetailSubjectKeyword = "工商银行历史明细";
        private const string ICBCHistoryDetailRowCodePrefix = "ICBCHistory-";
        private static readonly Regex ICBCHistoryDetailSubjectRegex = new(
            @"^工商银行历史明细（申请单号：(?<applicationNo>\d+)）\s*$",
            RegexOptions.CultureInvariant);
        private static readonly Regex ICBCHistoryDetailHeaderRegex = new(
            @"卡号[:：]?\s*(?<cardNo>[\d*]+)\s*户名[:：]\s*(?<name>.+?)\s*起止日期[:：]\s*(?<start>\d{4}-\d{2}-\d{2})[\s\S]{0,40}?(?<end>\d{4}-\d{2}-\d{2})",
            RegexOptions.CultureInvariant);
        private static readonly Regex ICBCHistoryDetailOrderTimeRegex = new(
            @"下单时间[:：]\s*(?<time>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})",
            RegexOptions.CultureInvariant);

        public async Task<int> DebugDownloadICBCHistoryDetails(DateTime mailDate, string? directory)
        {
            var targetDirectory = GetICBCHistoryDetailDirectory(directory);
            Directory.CreateDirectory(targetDirectory);
            var messages = await SearchICBCHistoryDetailMessages(mailDate.Date);
            var savedCount = 0;
            foreach (var message in messages)
            {
                savedCount += SaveOpenableICBCHistoryDetailAttachments(message, targetDirectory, null).Count;
            }

            Console.WriteLine($"Saved ICBC history PDFs: {savedCount}");
            return savedCount;
        }

        public async Task<string> DebugDownloadLatestICBCHistoryDetail(string? directory)
        {
            var targetDirectory = GetICBCHistoryDetailDirectory(directory);
            Directory.CreateDirectory(targetDirectory);
            var mailbox = CreateYahooMailbox();
            using ImapClient client = new();
            client.Timeout = 120000;
            client.CheckCertificateRevocation = false;
            client.ProxyClient = null;

            Console.WriteLine($"mail connect direct {mailbox.Label} latest ICBC history detail");
            await RunMailOperation(token => client.ConnectAsync(mailbox.Host, mailbox.Port, mailbox.UseSsl, token));
            Console.WriteLine("mail connected");
            await RunMailOperation(token => client.AuthenticateAsync(mailbox.Username, mailbox.Password, token));
            Console.WriteLine("mail authenticated");
            var folder = await GetMailSearchFolder(client, mailbox);
            await RunMailOperation(token => folder.OpenAsync(FolderAccess.ReadOnly, token));
            Console.WriteLine($"mail folder opened {folder.FullName}");

            var query = SearchQuery.FromContains(ICBCHistoryDetailSender)
                .And(SearchQuery.SubjectContains(ICBCHistoryDetailSubjectKeyword))
                .And(SearchQuery.SentSince(DateTime.Today.AddYears(-2)));
            var uids = await RunMailOperation(token => folder.SearchAsync(query, token));
            Console.WriteLine($"mail search latest ICBC history detail found {uids.Count}");
            if (uids.Count == 0)
                throw new InvalidOperationException("No ICBC history detail mail found.");

            var summaries = await RunMailOperation(token => folder.FetchAsync(
                uids,
                MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId,
                token));
            var candidates = summaries
                .Where(IsICBCHistoryDetailSummary)
                .OrderByDescending(GetMessageSummaryDate)
                .ThenByDescending(summary => summary.UniqueId.Id)
                .ToList();
            foreach (var summary in candidates)
            {
                var message = await RunMailOperation(token => folder.GetMessageAsync(summary.UniqueId, token));
                var savedFiles = SaveOpenableICBCHistoryDetailAttachments(message, targetDirectory, 1);
                if (savedFiles.Count == 0)
                    continue;

                Console.WriteLine($"Saved latest ICBC history PDF: {savedFiles[0]}");
                return savedFiles[0];
            }

            throw new InvalidOperationException("No openable ICBC history detail PDF found.");
        }

        public int DebugFetchLocalICBCHistoryDetails(string? directory)
        {
            var targetDirectory = GetICBCHistoryDetailDirectory(directory);
            var files = Directory.GetFiles(targetDirectory, "*.pdf")
                .Where(IsLikelyLocalICBCHistoryDetailPdf)
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (files.Count == 0)
                throw new FileNotFoundException($"No ICBC history detail PDF found in {targetDirectory}");

            var importedCount = 0;
            foreach (var file in files)
            {
                var parsed = ParseICBCHistoryDetailPdfFile(file);
                if (ImportICBCHistoryDetail(parsed))
                    importedCount++;
            }

            Console.WriteLine($"Imported ICBC history detail statements: {importedCount}");
            return importedCount;
        }

        public async Task<int> FetchICBCHistoryDetails(DateTime since)
        {
            var mailbox = CreateYahooMailbox();
            using ImapClient client = new();
            client.Timeout = MailClientTimeoutMilliseconds;
            client.CheckCertificateRevocation = false;
            client.ProxyClient = null;

            Console.WriteLine($"mail connect direct {mailbox.Label} ICBC history detail since {since:yyyy-MM-dd}");
            await RunMailOperation(token => client.ConnectAsync(mailbox.Host, mailbox.Port, mailbox.UseSsl, token));
            Console.WriteLine("mail connected");
            await RunMailOperation(token => client.AuthenticateAsync(mailbox.Username, mailbox.Password, token));
            Console.WriteLine("mail authenticated");
            var folder = await GetMailSearchFolder(client, mailbox);
            await RunMailOperation(token => folder.OpenAsync(FolderAccess.ReadOnly, token));
            Console.WriteLine($"mail folder opened {folder.FullName}");

            var query = SearchQuery.FromContains(ICBCHistoryDetailSender)
                .And(SearchQuery.SubjectContains(ICBCHistoryDetailSubjectKeyword))
                .And(SearchQuery.SentSince(since.Date));
            var uids = await RunMailOperation(token => folder.SearchAsync(query, token));
            Console.WriteLine($"mail search ICBC history detail since {since:yyyy-MM-dd} found {uids.Count}");
            if (uids.Count == 0)
                return 0;

            var summaries = await RunMailOperation(token => folder.FetchAsync(
                uids,
                MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId,
                token));
            var candidates = summaries
                .Where(IsICBCHistoryDetailSummary)
                .OrderBy(GetMessageSummaryDate)
                .ThenBy(summary => summary.UniqueId.Id)
                .ToList();

            var importedCount = 0;
            var skippedUnreadableCount = 0;
            foreach (var summary in candidates)
            {
                var message = await RunMailOperation(token => folder.GetMessageAsync(summary.UniqueId, token));
                var parsedAttachments = ParseICBCHistoryDetailAttachments(message);
                skippedUnreadableCount += parsedAttachments.SkippedUnreadableCount;
                foreach (var attachment in parsedAttachments.Statements)
                {
                    var parsed = ParseICBCHistoryDetailPdfText(
                        attachment.Text,
                        attachment.FileName,
                        attachment.FileHash);
                    if (ImportICBCHistoryDetail(parsed))
                        importedCount++;
                }
            }

            Console.WriteLine($"Imported remote ICBC history detail statements: {importedCount}, skippedUnreadable={skippedUnreadableCount}");
            return importedCount;
        }

        private static bool IsICBCHistoryDetailSummary(IMessageSummary summary)
        {
            var subject = summary.Envelope?.Subject ?? "";
            return TryParseICBCHistoryDetailApplicationNo(subject, out _)
                && summary.Envelope?.From?.Mailboxes.Any(mailbox =>
                    String.Equals(mailbox.Address, ICBCHistoryDetailSender, StringComparison.OrdinalIgnoreCase)) == true;
        }

        private static DateTime GetMessageSummaryDate(IMessageSummary summary)
        {
            return summary.Envelope?.Date?.LocalDateTime ?? DateTime.MinValue;
        }

        private List<string> SaveOpenableICBCHistoryDetailAttachments(
            MimeMessage message,
            string targetDirectory,
            int? maxCount)
        {
            var applicationNo = TryParseICBCHistoryDetailApplicationNo(message.Subject ?? "", out var parsedApplicationNo)
                ? parsedApplicationNo
                : "";
            var attachments = ReadMatchingAttachments(message, (attachment, fileName) =>
            {
                if (!fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    return null;

                var bytes = ReadMimePartBytes(attachment);
                if (!TryReadICBCHistoryDetailPdfText(bytes, out _, out var error))
                {
                    Console.WriteLine($"Skip encrypted/unreadable ICBC history PDF: {fileName}; {error}");
                    return null;
                }

                var decodedFileName = LooksLikeGbkMojibake(fileName) ? DecodeLatin1AsGbk(fileName) : fileName;
                var safeName = String.IsNullOrWhiteSpace(decodedFileName)
                    ? $"工商银行历史明细（申请单号：{applicationNo}）.pdf"
                    : decodedFileName;
                return new ICBCHistoryDetailAttachment(safeName, bytes);
            });

            var savedFiles = new List<string>();
            foreach (var attachment in maxCount.HasValue ? attachments.Take(maxCount.Value) : attachments)
            {
                var filePath = GetUniqueFilePath(targetDirectory, SanitizeFileName(attachment.FileName));
                File.WriteAllBytes(filePath, attachment.Bytes);
                Console.WriteLine($"Saved ICBC history PDF: {filePath}");
                savedFiles.Add(filePath);
            }

            return savedFiles;
        }

        private ICBCHistoryDetailParsedAttachments ParseICBCHistoryDetailAttachments(MimeMessage message)
        {
            var statements = new List<ICBCHistoryDetailParsedAttachment>();
            var skippedUnreadableCount = 0;
            foreach (var attachment in ReadMatchingAttachments(message, (attachment, fileName) =>
            {
                if (!fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    return null;

                var bytes = ReadMimePartBytes(attachment);
                if (!TryReadICBCHistoryDetailPdfText(bytes, out var text, out var error))
                {
                    Console.WriteLine($"Skip encrypted/unreadable ICBC history PDF: {fileName}; {error}");
                    skippedUnreadableCount++;
                    return null;
                }

                var decodedFileName = LooksLikeGbkMojibake(fileName) ? DecodeLatin1AsGbk(fileName) : fileName;
                if (String.IsNullOrWhiteSpace(decodedFileName)
                    && TryParseICBCHistoryDetailApplicationNo(message.Subject ?? "", out var applicationNo))
                    decodedFileName = $"ICBCHistoryDetail-{applicationNo}.pdf";

                return new ICBCHistoryDetailParsedAttachment(decodedFileName, text, ComputeSha256(bytes));
            }))
            {
                statements.Add(attachment);
            }

            return new ICBCHistoryDetailParsedAttachments(statements, skippedUnreadableCount);
        }

        private static bool IsLikelyLocalICBCHistoryDetailPdf(string file)
        {
            var fileName = Path.GetFileName(file);
            return fileName.StartsWith("工商银行历史明细", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(fileName, @"^\d{18,}-\d{8}\.pdf$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private async Task<List<MimeMessage>> SearchICBCHistoryDetailMessages(DateTime mailDate)
        {
            var query = SearchQuery.FromContains(ICBCHistoryDetailSender)
                .And(SearchQuery.SubjectContains(ICBCHistoryDetailSubjectKeyword))
                .And(SearchQuery.SentSince(mailDate.Date))
                .And(SearchQuery.SentBefore(mailDate.Date.AddDays(1)));
            return await SearchMessages(
                $"ICBC history details {mailDate:yyyy-MM-dd}",
                query,
                IsICBCHistoryDetailMessage,
                GetMailDateTime);
        }

        private static bool IsICBCHistoryDetailMessage(MimeMessage message)
        {
            return IsFrom(message, ICBCHistoryDetailSender)
                && TryParseICBCHistoryDetailApplicationNo(message.Subject ?? "", out _)
                && HasMatchingAttachment(
                    message,
                    (_, fileName) => fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryParseICBCHistoryDetailApplicationNo(string subject, out string applicationNo)
        {
            applicationNo = "";
            var match = ICBCHistoryDetailSubjectRegex.Match(subject.Trim());
            if (!match.Success)
                return false;

            applicationNo = match.Groups["applicationNo"].Value;
            return true;
        }

        private ICBCHistoryDetailParsedStatement ParseICBCHistoryDetailPdfFile(string file)
        {
            var bytes = File.ReadAllBytes(file);
            if (!TryReadICBCHistoryDetailPdfText(bytes, out var text, out var error))
                throw new MailParseException($"Parse ICBC history detail PDF fail: {file}; {error}");

            return ParseICBCHistoryDetailPdfText(text, Path.GetFileName(file), ComputeSha256(bytes));
        }

        private static bool TryReadICBCHistoryDetailPdfText(byte[] pdfBytes, out string text, out string error)
        {
            text = "";
            error = "";
            try
            {
                using var document = PdfDocument.Open(pdfBytes);
                var builder = new StringBuilder();
                foreach (var page in document.GetPages())
                {
                    builder.AppendLine(page.Text);
                }

                text = builder.ToString();
                return true;
            }
            catch (Exception exception)
            {
                error = exception.Message;
                return false;
            }
        }

        private ICBCHistoryDetailParsedStatement ParseICBCHistoryDetailPdfText(
            string text,
            string fileName,
            string fileHash)
        {
            var normalizedText = NormalizeICBCHistoryDetailText(text);
            var headerMatch = ICBCHistoryDetailHeaderRegex.Match(normalizedText);
            if (!headerMatch.Success)
                throw new MailParseException($"Parse ICBC history detail PDF fail, missing header: {fileName}; text={TruncateForLog(NormalizeICBCHistoryText(normalizedText), 300)}");

            var cardNo = NormalizeICBCHistoryText(headerMatch.Groups["cardNo"].Value);
            var cardTail = GetICBCCardTail(cardNo);
            var account = database.GetAccountByTypeAndId(ICBCAccountType, cardTail);
            var postingAccount = database.GetPostingAccount(account);
            var isCredit = account.isCredit || postingAccount.isCredit;
            var startDate = DateTime.ParseExact(headerMatch.Groups["start"].Value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(headerMatch.Groups["end"].Value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            var applicationNo = ParseICBCHistoryDetailApplicationNoFromFileName(fileName);
            if (String.IsNullOrWhiteSpace(applicationNo))
                applicationNo = $"sha256-{fileHash[..16]}";
            var statementKey = BuildICBCHistoryDetailStatementKey(postingAccount.name, startDate, endDate, applicationNo);
            var importTime = ParseICBCHistoryDetailOrderTime(normalizedText) ?? endDate;
            var rows = ExtractICBCHistoryDetailRows(normalizedText)
                .Select(ParseICBCHistoryDetailRow)
                .ToList();
            if (rows.Count == 0)
                throw new MailParseException($"Parse ICBC history detail PDF fail, no transaction rows: {fileName}");

            var records = new Records();
            var candidates = new List<ICBCHistoryDetailCandidate>();
            var debitCandidates = new List<ICBCHistoryDetailDebitCandidate>();
            var rowCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in rows)
            {
                var rowCode = BuildICBCHistoryDetailRowCode(postingAccount.name, row);
                if (!rowCodes.Add(rowCode))
                    continue;

                if (isCredit)
                {
                    candidates.Add(BuildICBCHistoryDetailCandidate(
                        candidates.Count,
                        row,
                        postingAccount,
                        statementKey,
                        rowCode));
                }
                else
                {
                    var record = BuildICBCHistoryDetailRecord(row, account, postingAccount, statementKey, rowCode);
                    records.Add(record);
                    debitCandidates.Add(new ICBCHistoryDetailDebitCandidate(
                        debitCandidates.Count,
                        row,
                        record,
                        rowCode));
                }
            }

            return new ICBCHistoryDetailParsedStatement(
                fileName,
                statementKey,
                postingAccount.name,
                isCredit,
                startDate,
                endDate,
                importTime,
                records,
                candidates,
                debitCandidates,
                [
                    new AccountInternalId
                    {
                        Account = account,
                        cardNo = cardNo,
                        desc = "ICBC history detail header card no",
                        sourceText = $"ICBC history detail file={fileName}; header cardNo={cardNo}"
                    }
                ]);
        }

        private Record BuildICBCHistoryDetailRecord(
            ICBCHistoryDetailRow row,
            Account cardAccount,
            Account postingAccount,
            string statementKey,
            string rowCode)
        {
            var record = new Record
            {
                Account = postingAccount,
                date = row.PostingDate,
                postingDate = row.PostingDate,
                updateTime = DateTime.Now,
                Source = BuildICBCHistoryDetailSource(row, rowCode, statementKey),
                DestAccount = BuildICBCHistoryDetailDestAccount(row),
                Reason = row.Summary
            };
            record.CopyFrom(row.Amount);
            record.DescCurrency = row.DescCurrency;

            var internalCounterparty = database.FindAccountByInternalCardNoText(
                ICBCAccountType,
                $"ICBC history detail statement={statementKey}; postingDate={row.PostingDate:yyyy-MM-dd HH:mm:ss}; row={row.RawText}",
                row.CounterpartyName,
                row.CounterpartyAccount,
                record.DestAccount);
            if (internalCounterparty is not null && !IsSameAccount(database.GetPostingAccount(internalCounterparty), postingAccount))
            {
                record.DestAccount = database.GetPostingAccount(internalCounterparty).name;
                record.isInternal = true;
            }

            if (IsICBCHistoryDetailSelfTransfer(row, cardAccount))
                record.isInternal = true;

            return record;
        }

        private ICBCHistoryDetailCandidate BuildICBCHistoryDetailCandidate(
            int index,
            ICBCHistoryDetailRow row,
            Account postingAccount,
            string statementKey,
            string rowCode)
        {
            var destAccount = BuildICBCHistoryDetailDestAccount(row);
            var internalCounterparty = database.FindAccountByInternalCardNoText(
                ICBCAccountType,
                $"ICBC history detail statement={statementKey}; postingDate={row.PostingDate:yyyy-MM-dd HH:mm:ss}; row={row.RawText}",
                row.CounterpartyName,
                row.CounterpartyAccount,
                destAccount);
            if (internalCounterparty is not null && !IsSameAccount(database.GetPostingAccount(internalCounterparty), postingAccount))
                destAccount = database.GetPostingAccount(internalCounterparty).name;

            return new ICBCHistoryDetailCandidate(
                index,
                row.PostingDate,
                row.Amount,
                row.DescCurrency,
                destAccount,
                $"code={rowCode}; ICBC鍘嗗彶鏄庣粏; statementKey={statementKey}; row={row.RawText}",
                rowCode);
        }

        private bool ImportICBCHistoryDetail(ICBCHistoryDetailParsedStatement parsed)
        {
            if (database.IsStatementKeyImported(ICBCHistoryDetailProvider, parsed.StatementKey))
            {
                Console.WriteLine($"Skip imported ICBC history detail {parsed.StatementKey}");
                return false;
            }

            return parsed.IsCredit
                ? ImportICBCCreditHistoryDetail(parsed)
                : ImportICBCDebitHistoryDetail(parsed);
        }

        private bool ImportICBCDebitHistoryDetail(ICBCHistoryDetailParsedStatement parsed)
        {
            var stats = new ICBCHistoryDetailImportStats();
            var account = database.GetAccountByName(parsed.AccountName);
            var candidates = parsed.DebitCandidates;
            ValidateICBCHistoryDetailCandidateOrder(parsed.StatementKey, candidates);
            if (!TryValidateICBCHistoryDetailBalanceChain(
                    parsed.StatementKey,
                    candidates,
                    out var balanceChainMismatch))
            {
                var marked = database.MarkStatementProcessedOnce(
                    ICBCHistoryDetailProvider,
                    parsed.ImportTime,
                    parsed.StatementKey,
                    parsed.InternalCardNos);
                stats.Ignore("incomplete balance chain", candidates.Count);
                Console.WriteLine(
                    marked
                        ? $"Skip incomplete ICBC history detail {parsed.StatementKey}, markedProcessed=true, imported=0, ignored={stats.FormatIgnored()}, {balanceChainMismatch}"
                        : $"Skip imported incomplete ICBC history detail {parsed.StatementKey}, imported=0, ignored={stats.FormatIgnored()}");
                return false;
            }

            var existingRecords = database.GetStatementRecords(ICBCHistoryDetailProvider, account)
                .Where(record => !DatabaseUtil.IsInitializationRecord(record))
                .OrderBy(GetICBCHistoryDetailRecordPostingDate)
                .ThenBy(record => record.Id)
                .ToList();
            var newCandidates = SelectSequentialICBCHistoryDetailCandidates(
                parsed,
                candidates,
                existingRecords,
                stats);
            if (newCandidates.Count == 0)
            {
                Console.WriteLine($"Skip ICBC history detail {parsed.StatementKey}, imported=0, ignored={stats.FormatIgnored()}");
                return false;
            }

            var records = newCandidates.Select(candidate => candidate.Record).ToList();
            var beginningBalances = BuildICBCHistoryDetailBeginningBalances(account, newCandidates);
            var endingBalances = BuildICBCHistoryDetailEndingBalances(account, newCandidates);
            var saved = database.SaveStatementRecordsOnce(
                ICBCHistoryDetailProvider,
                parsed.ImportTime,
                records,
                endingBalances,
                parsed.StatementKey,
                beginningBalances,
                afterSaveInTransaction: _ => database.EnsureAccountInternalCardNos(parsed.InternalCardNos),
                forceValidateBeginningBalances: existingRecords.Count > 0);
            stats.Imported = saved ? records.Count : 0;
            Console.WriteLine(saved
                ? $"Import ICBC history detail {parsed.StatementKey}, imported={stats.Imported}, overlap={stats.Overlap}, balanceUpdates={endingBalances.Count}, ignored={stats.FormatIgnored()}"
                : $"Skip imported ICBC history detail {parsed.StatementKey}, ignored={stats.FormatIgnored()}");
            return saved;
        }

        private List<ICBCHistoryDetailDebitCandidate> SelectSequentialICBCHistoryDetailCandidates(
            ICBCHistoryDetailParsedStatement parsed,
            List<ICBCHistoryDetailDebitCandidate> candidates,
            List<Record> existingRecords,
            ICBCHistoryDetailImportStats stats)
        {
            if (existingRecords.Count == 0)
                return candidates;

            var firstCandidate = candidates[0];
            var existingStart = GetICBCHistoryDetailRecordPostingDate(existingRecords[0]);
            var existingEnd = GetICBCHistoryDetailRecordPostingDate(existingRecords[^1]);
            var candidatesStart = firstCandidate.Row.PostingDate;
            var candidatesEnd = candidates[^1].Row.PostingDate;
            var overlapCandidates = new List<int>();
            for (var existingIndex = 0; existingIndex < existingRecords.Count; existingIndex++)
            {
                if (!IsICBCHistoryDetailRecordMatch(firstCandidate, existingRecords[existingIndex]))
                    continue;

                var overlapLength = Math.Min(candidates.Count, existingRecords.Count - existingIndex);
                var matches = true;
                for (var candidateIndex = 0; candidateIndex < overlapLength; candidateIndex++)
                {
                    if (IsICBCHistoryDetailRecordMatch(
                            candidates[candidateIndex],
                            existingRecords[existingIndex + candidateIndex]))
                        continue;

                    matches = false;
                    break;
                }

                if (!matches)
                    continue;

                overlapCandidates.Add(overlapLength);
            }

            if (overlapCandidates.Count == 0)
            {
                if (candidatesEnd < existingStart || candidatesStart > existingEnd)
                {
                    stats.Ignore("out-of-order or non-overlapping", candidates.Count);
                    return [];
                }

                throw new InvalidOperationException(
                    $"ICBC history detail overlaps existing records but the starting transaction does not match: {parsed.StatementKey}");
            }

            if (overlapCandidates.Count > 1)
            {
                throw new InvalidOperationException(
                    $"ICBC history detail overlaps existing records ambiguously: {parsed.StatementKey}");
            }

            var bestOverlapLength = overlapCandidates[0];
            stats.Overlap = bestOverlapLength;
            if (bestOverlapLength == candidates.Count)
            {
                stats.Ignore("old or duplicate", candidates.Count);
                return [];
            }

            return candidates.Skip(bestOverlapLength).ToList();
        }

        private static void ValidateICBCHistoryDetailCandidateOrder(
            string statementKey,
            List<ICBCHistoryDetailDebitCandidate> candidates)
        {
            for (var index = 1; index < candidates.Count; index++)
            {
                if (candidates[index - 1].Row.PostingDate <= candidates[index].Row.PostingDate)
                    continue;

                throw new MailParseException(
                    $"ICBC history detail records are not ordered by posting time: {statementKey}; previous={candidates[index - 1].Row.PostingDate:yyyy-MM-dd HH:mm:ss}; current={candidates[index].Row.PostingDate:yyyy-MM-dd HH:mm:ss}");
            }
        }

        private static bool TryValidateICBCHistoryDetailBalanceChain(
            string statementKey,
            List<ICBCHistoryDetailDebitCandidate> candidates,
            out string mismatch)
        {
            mismatch = "";
            foreach (var group in candidates.GroupBy(candidate => candidate.Row.Amount.t))
            {
                var groupCandidates = group.ToList();
                if (groupCandidates.Any(candidate => !candidate.Row.Balance.HasValue))
                {
                    mismatch = $"balance is required: currency={group.Key}";
                    return false;
                }

                var previous = groupCandidates[0].Row.Balance!.Value - groupCandidates[0].Row.Amount.v;
                foreach (var candidate in groupCandidates)
                {
                    var expected = previous + candidate.Row.Amount.v;
                    var actual = candidate.Row.Balance!.Value;
                    if (expected != actual)
                    {
                        mismatch =
                            $"balance chain mismatch: currency={group.Key}; postingDate={candidate.Row.PostingDate:yyyy-MM-dd HH:mm:ss}; expected={expected}; actual={actual}; row={candidate.Row.RawText}";
                        return false;
                    }

                    previous = actual;
                }
            }

            return true;
        }

        private static List<AccountBalance> BuildICBCHistoryDetailBeginningBalances(
            Account account,
            List<ICBCHistoryDetailDebitCandidate> candidates)
        {
            return candidates
                .GroupBy(candidate => candidate.Row.Amount.t)
                .Select(group =>
                {
                    var first = group.First();
                    return new AccountBalance(
                        account,
                        new Currency(first.Row.Balance!.Value - first.Row.Amount.v, group.Key));
                })
                .ToList();
        }

        private static List<AccountBalance> BuildICBCHistoryDetailEndingBalances(
            Account account,
            List<ICBCHistoryDetailDebitCandidate> candidates)
        {
            return candidates
                .GroupBy(candidate => candidate.Row.Amount.t)
                .Select(group =>
                {
                    var last = group.Last();
                    return new AccountBalance(
                        account,
                        new Currency(last.Row.Balance!.Value, group.Key));
                })
                .ToList();
        }

        private static bool IsICBCHistoryDetailRecordMatch(
            ICBCHistoryDetailDebitCandidate candidate,
            Record record)
        {
            if (GetICBCHistoryDetailRecordPostingDate(record) != candidate.Row.PostingDate
                || record.v != candidate.Row.Amount.v
                || record.t != candidate.Row.Amount.t
                || record.DescCurrency != candidate.Row.DescCurrency
                || !String.Equals(record.Reason, candidate.Record.Reason, StringComparison.Ordinal)
                || !String.Equals(record.DestAccount, candidate.Record.DestAccount, StringComparison.Ordinal))
                return false;

            if (!TryGetICBCHistoryDetailRecordBalance(record, out var existingBalance))
                return true;

            return candidate.Row.Balance.HasValue
                && existingBalance.t == candidate.Row.Amount.t
                && existingBalance.v == candidate.Row.Balance.Value;
        }

        private static bool TryGetICBCHistoryDetailRecordBalance(Record record, out Currency balance)
        {
            if (TryParseICBCHistoryDetailSourceBalance(record.Source, out balance))
                return true;

            try
            {
                var rawRow = ExtractICBCHistoryDetailRawRow(record.Source);
                if (String.Equals(rawRow, record.Source, StringComparison.Ordinal))
                    return false;

                var row = ParseICBCHistoryDetailRow(rawRow);
                if (!row.Balance.HasValue)
                    return false;

                balance = new Currency(row.Balance.Value, row.Amount.t);
                return true;
            }
            catch
            {
                balance = null!;
                return false;
            }
        }

        private static bool TryParseICBCHistoryDetailSourceBalance(string source, out Currency balance)
        {
            var match = Regex.Match(
                source ?? "",
                @"(?:^|;\s*)historyBalance=(?<currency>[A-Za-z0-9_]+):(?<amount>[+-]?\d[\d,]*(?:\.\d+)?)",
                RegexOptions.CultureInvariant);
            if (!match.Success
                || !Enum.TryParse<CurrencyType>(match.Groups["currency"].Value, out var currencyType))
            {
                balance = null!;
                return false;
            }

            balance = new Currency(ParseInvariantDecimal(match.Groups["amount"].Value), currencyType);
            return true;
        }

        private static string BuildICBCHistoryDetailSource(
            ICBCHistoryDetailRow row,
            string rowCode,
            string statementKey)
        {
            var parts = new List<string>
            {
                $"code={rowCode}",
                "ICBCHistoryDetail",
                $"statementKey={statementKey}"
            };
            if (row.Balance.HasValue)
                parts.Add($"historyBalance={row.Amount.t}:{row.Balance.Value.ToString(CultureInfo.InvariantCulture)}");

            parts.Add($"row={row.RawText}");
            return LimitICBCHistoryDetailRecordText(String.Join("; ", parts));
        }

        private static DateTime GetICBCHistoryDetailRecordPostingDate(Record record)
        {
            if (!record.postingDate.HasValue)
                throw new InvalidOperationException($"ICBC history detail record missing postingDate: {record.Id}");

            return record.postingDate.Value;
        }

        private bool ImportICBCCreditHistoryDetail(ICBCHistoryDetailParsedStatement parsed)
        {
            var stats = new ICBCHistoryDetailImportStats();
            var existingCodes = database.GetRecordSourceCodes(ICBCHistoryDetailRowCodePrefix);
            var candidates = parsed.Candidates
                .Where(candidate =>
                {
                    if (!existingCodes.Contains(candidate.RowCode))
                        return true;

                    stats.Ignore("duplicate", 1);
                    return false;
                })
                .ToList();
            if (candidates.Count == 0)
            {
                Console.WriteLine($"Skip ICBC credit history detail {parsed.StatementKey}, ignored={stats.FormatIgnored()}");
                return false;
            }

            var account = database.GetAccountByName(parsed.AccountName);
            var monthlyRecords = database.GetStatementRecords(
                ICBCProvider,
                account,
                parsed.StartDate.Date.AddMonths(-1),
                parsed.EndDate.Date.AddDays(1))
                .Where(record => !DatabaseUtil.IsInitializationRecord(record))
                .ToList();
            if (monthlyRecords.Count == 0)
            {
                stats.Ignore("no monthly bill", candidates.Count);
                Console.WriteLine($"Skip ICBC credit history detail {parsed.StatementKey}, imported=0, ignored={stats.FormatIgnored()}");
                return false;
            }

            var periods = BuildICBCMonthlyStatementPeriods()
                .Where(period => period.EndDate >= parsed.StartDate.Date && period.StartDate <= parsed.EndDate.Date)
                .ToList();
            var usedCandidateIndexes = new HashSet<int>();
            var supplements = new List<RecordSourceSupplement>();
            foreach (var period in periods)
            {
                var periodAccountMonthlyRecords = monthlyRecords
                    .Where(record => record._statementImport_Id == period.StatementImport.Id)
                    .ToList();
                var periodStart = period.StartDate;
                if (periodStart == DateTime.MinValue.Date && periodAccountMonthlyRecords.Count > 0)
                    periodStart = periodAccountMonthlyRecords.Min(GetICBCMonthlyHistoryMatchDate);

                var overlapStart = MaxDate(parsed.StartDate.Date, periodStart);
                var overlapEnd = MinDate(parsed.EndDate.Date, period.EndDate);
                if (overlapStart > overlapEnd)
                    continue;

                var periodMonthlyRecords = periodAccountMonthlyRecords
                    .Where(record => GetICBCMonthlyHistoryMatchDate(record) >= overlapStart
                        && GetICBCMonthlyHistoryMatchDate(record) <= overlapEnd)
                    .OrderBy(GetICBCMonthlyHistoryMatchDate)
                    .ThenBy(record => record.date)
                    .ThenBy(record => record.Id)
                    .ToList();
                var periodCandidates = candidates
                    .Where(candidate => !usedCandidateIndexes.Contains(candidate.Index)
                        && candidate.PostingDate.Date >= overlapStart
                        && candidate.PostingDate.Date <= overlapEnd)
                    .OrderBy(candidate => candidate.PostingDate)
                    .ThenBy(candidate => candidate.Index)
                    .ToList();
                if (periodCandidates.Count == 0)
                {
                    if (periodMonthlyRecords.Count > 0 && overlapEnd < parsed.EndDate.Date)
                        stats.MissingMonthlyRows += periodMonthlyRecords.Count;
                    continue;
                }

                foreach (var candidate in periodCandidates)
                    usedCandidateIndexes.Add(candidate.Index);

                if (periodMonthlyRecords.Count == 0)
                {
                    stats.Ignore("no monthly record in period", periodCandidates.Count);
                    continue;
                }

                if (!TryBuildICBCHistoryDetailSupplements(
                        parsed,
                        periodMonthlyRecords,
                        periodCandidates,
                        supplements,
                        out var mismatchReason,
                        out var missingMonthlyRows))
                {
                    stats.Ignore(mismatchReason, periodCandidates.Count);
                    stats.MissingMonthlyRows += missingMonthlyRows;
                }
            }

            var noMonthlyCount = candidates.Count(candidate => !usedCandidateIndexes.Contains(candidate.Index));
            stats.Ignore("no overlapping monthly bill", noMonthlyCount);
            stats.Imported = supplements.Count;
            if (supplements.Count == 0)
            {
                Console.WriteLine($"Skip ICBC credit history detail {parsed.StatementKey}, imported=0, ignored={stats.FormatIgnored()}, missingMonthlyRows={stats.MissingMonthlyRows}");
                return false;
            }

            var saved = database.SaveRecordSourceSupplementsOnce(
                ICBCHistoryDetailProvider,
                parsed.ImportTime,
                parsed.StatementKey,
                supplements,
                parsed.InternalCardNos);
            Console.WriteLine(saved
                ? $"Import ICBC credit history detail {parsed.StatementKey}, imported={stats.Imported}, ignored={stats.FormatIgnored()}, missingMonthlyRows={stats.MissingMonthlyRows}"
                : $"Skip imported ICBC credit history detail {parsed.StatementKey}");
            return saved;
        }

        private List<ICBCMonthlyStatementPeriod> BuildICBCMonthlyStatementPeriods()
        {
            var imports = database.GetStatementImports(ICBCProvider)
                .Select(statementImport => new
                {
                    StatementImport = statementImport,
                    EndDate = TryParseICBCMonthlyStatementEndDate(statementImport.statementKey)
                })
                .Where(item => item.EndDate.HasValue)
                .OrderBy(item => item.EndDate!.Value)
                .ThenBy(item => item.StatementImport.Id)
                .ToList();
            var periods = new List<ICBCMonthlyStatementPeriod>();
            DateTime? previousEndDate = null;
            foreach (var import in imports)
            {
                var endDate = import.EndDate!.Value;
                var startDate = previousEndDate?.AddDays(1) ?? DateTime.MinValue.Date;
                periods.Add(new ICBCMonthlyStatementPeriod(import.StatementImport, startDate, endDate));
                previousEndDate = endDate;
            }

            return periods;
        }

        private bool TryBuildICBCHistoryDetailSupplements(
            ICBCHistoryDetailParsedStatement parsed,
            List<Record> monthlyRecords,
            List<ICBCHistoryDetailCandidate> candidates,
            List<RecordSourceSupplement> supplements,
            out string mismatchReason,
            out int missingMonthlyRows)
        {
            mismatchReason = "";
            missingMonthlyRows = 0;
            var pending = new List<RecordSourceSupplement>();
            var dates = monthlyRecords
                .Select(GetICBCMonthlyHistoryMatchDate)
                .Union(candidates.Select(candidate => candidate.PostingDate.Date))
                .OrderBy(date => date)
                .ToList();
            foreach (var date in dates)
            {
                var dayMonthlyRecords = monthlyRecords
                    .Where(record => GetICBCMonthlyHistoryMatchDate(record) == date)
                    .OrderBy(GetICBCMonthlyHistoryMatchDate)
                    .ThenBy(record => record.date)
                    .ThenBy(record => record.Id)
                    .ToList();
                var dayCandidates = candidates
                    .Where(candidate => candidate.PostingDate.Date == date)
                    .OrderBy(candidate => candidate.PostingDate)
                    .ThenBy(candidate => candidate.Index)
                    .ToList();
                var isLastDetailDay = date == parsed.EndDate.Date;
                if (!isLastDetailDay && dayMonthlyRecords.Count != dayCandidates.Count)
                {
                    mismatchReason = "monthly segment mismatch";
                    missingMonthlyRows += Math.Max(0, dayMonthlyRecords.Count - dayCandidates.Count);
                    return false;
                }

                if (isLastDetailDay && dayCandidates.Count > dayMonthlyRecords.Count)
                {
                    mismatchReason = "monthly segment mismatch";
                    return false;
                }

                if (isLastDetailDay)
                    missingMonthlyRows += dayMonthlyRecords.Count - dayCandidates.Count;

                var unmatchedMonthlyRecords = dayMonthlyRecords.ToList();
                foreach (var dayCandidate in dayCandidates)
                {
                    var monthlyIndex = unmatchedMonthlyRecords.FindIndex(monthlyRecord =>
                        IsICBCHistoryMonthlyRecordMatch(dayCandidate, monthlyRecord));
                    if (monthlyIndex < 0)
                    {
                        mismatchReason = "monthly segment mismatch";
                        return false;
                    }

                    var monthlyRecord = unmatchedMonthlyRecords[monthlyIndex];
                    unmatchedMonthlyRecords.RemoveAt(monthlyIndex);
                    pending.Add(new RecordSourceSupplement(
                        monthlyRecord.Id,
                        dayCandidate.RowCode,
                        BuildICBCHistoryDetailSupplementSource(parsed, dayCandidate)));
                }
            }

            supplements.AddRange(pending);
            return true;
        }

        private static bool IsICBCHistoryMonthlyRecordMatch(ICBCHistoryDetailCandidate historyRecord, Record monthlyRecord)
        {
            if (historyRecord.PostingDate.Date != GetICBCMonthlyHistoryMatchDate(monthlyRecord)
                || historyRecord.Amount.v != monthlyRecord.v
                || historyRecord.Amount.t != monthlyRecord.t
                || historyRecord.DescCurrency != monthlyRecord.DescCurrency)
                return false;

            var historyDest = NormalizeICBCHistoryComparableText(historyRecord.DestAccount);
            var monthlyDest = NormalizeICBCHistoryComparableText(monthlyRecord.DestAccount);
            return String.IsNullOrWhiteSpace(historyDest)
                || String.IsNullOrWhiteSpace(monthlyDest)
                || String.Equals(historyDest, monthlyDest, StringComparison.Ordinal);
        }

        private static DateTime GetICBCMonthlyHistoryMatchDate(Record record)
        {
            return (record.postingDate ?? record.date).Date;
        }

        private static string BuildICBCHistoryDetailSupplementSource(
            ICBCHistoryDetailParsedStatement parsed,
            ICBCHistoryDetailCandidate candidate)
        {
            return LimitICBCHistoryDetailRecordText(
                $"code={candidate.RowCode}; supplementedBy=ICBCHistoryDetail; statementKey={parsed.StatementKey}; row={ExtractICBCHistoryDetailRawRow(candidate.Source)}");
        }

        private static string ExtractICBCHistoryDetailRawRow(string source)
        {
            var match = Regex.Match(source, @"(?:^|;\s*)row=(?<row>.+)$", RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["row"].Value.Trim() : source;
        }

        private static string NormalizeICBCHistoryComparableText(string value)
        {
            return Regex.Replace(value ?? "", @"\s+", "").Trim();
        }

        private static DateTime? TryParseICBCMonthlyStatementEndDate(string statementKey)
        {
            return DateTime.TryParseExact(
                statementKey,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date)
                ? date.Date
                : null;
        }

        private static DateTime MaxDate(DateTime left, DateTime right)
        {
            return left >= right ? left : right;
        }

        private static DateTime MinDate(DateTime left, DateTime right)
        {
            return left <= right ? left : right;
        }

        private static string LimitICBCHistoryDetailRecordText(string text)
        {
            const int maxLength = 1024;
            var normalized = NormalizeICBCHistoryText(text);
            return normalized.Length <= maxLength ? normalized : normalized[..maxLength];
        }

        private static string NormalizeICBCHistoryDetailText(string text)
        {
            if (LooksLikeGbkMojibake(text))
                text = DecodeLatin1AsGbk(text);

            text = text.Replace("\r\n", "\n").Replace('\r', '\n');
            text = Regex.Replace(text, @"(\d{4}-\d{2}-\d{2})(\d{2}:\d{2}:\d{2})", "$1 $2", RegexOptions.CultureInvariant);
            text = Regex.Replace(text, @"(?<!\n)(?=\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})", "\n", RegexOptions.CultureInvariant);
            text = Regex.Replace(text, @"(?<!\n)(?=本页(?:支出|交易|收入))", "\n", RegexOptions.CultureInvariant);
            text = Regex.Replace(text, @"(?<!\n)(?=下单时间)", "\n", RegexOptions.CultureInvariant);
            text = Regex.Replace(text, @"(?<!\n)(?=共\s*\d+\s*页)", "\n", RegexOptions.CultureInvariant);
            text = Regex.Replace(text, @"(?<!\n)(?=中国工商银行\s*\d+)", "\n", RegexOptions.CultureInvariant);
            return text;
        }

        private static bool LooksLikeGbkMojibake(string text)
        {
            return text.Contains("ÖÐ¹ú", StringComparison.Ordinal)
                || text.Contains("¹¤ÉÌ", StringComparison.Ordinal)
                || text.Contains("ÈËÃñ±Ò", StringComparison.Ordinal);
        }

        private static string DecodeLatin1AsGbk(string text)
        {
            var encoding = Encoding.GetEncoding("GB18030");
            var result = new StringBuilder();
            var bytes = new List<byte>();
            foreach (var ch in text)
            {
                if (ch <= 255)
                {
                    bytes.Add((byte)ch);
                    continue;
                }

                AppendDecodedBytes(result, bytes, encoding);
                result.Append(ch);
            }

            AppendDecodedBytes(result, bytes, encoding);
            return result.ToString();
        }

        private static void AppendDecodedBytes(StringBuilder result, List<byte> bytes, Encoding encoding)
        {
            if (bytes.Count == 0)
                return;

            result.Append(encoding.GetString(bytes.ToArray()));
            bytes.Clear();
        }

        private static List<string> ExtractICBCHistoryDetailRows(string text)
        {
            var rows = new List<string>();
            var current = new StringBuilder();
            foreach (var rawLine in text.Split('\n'))
            {
                var line = NormalizeICBCHistoryText(rawLine);
                if (String.IsNullOrWhiteSpace(line) || IsICBCHistoryDetailNonDataLine(line))
                    continue;

                if (IsICBCHistoryDetailRowStart(line))
                {
                    AddICBCHistoryDetailRow(rows, current);
                    current.Clear();
                    current.Append(line);
                    continue;
                }

                if (current.Length > 0)
                    current.Append(' ').Append(line);
            }

            AddICBCHistoryDetailRow(rows, current);
            return rows;
        }

        private static void AddICBCHistoryDetailRow(List<string> rows, StringBuilder current)
        {
            if (current.Length == 0)
                return;

            var row = Regex.Replace(
                NormalizeICBCHistoryText(current.ToString()),
                @"\s*中国工商银行\s+\d+.*$",
                "",
                RegexOptions.CultureInvariant);
            if (!String.IsNullOrWhiteSpace(row))
                rows.Add(row);
        }

        private static bool IsICBCHistoryDetailRowStart(string line)
        {
            return Regex.IsMatch(line, @"^\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2}\b)?", RegexOptions.CultureInvariant);
        }

        private static bool IsICBCHistoryDetailNonDataLine(string line)
        {
            return line.StartsWith("请扫描二维码", StringComparison.Ordinal)
                || line.StartsWith("识别明细真伪", StringComparison.Ordinal)
                || line.StartsWith("中国工商银行借记账户历史明细", StringComparison.Ordinal)
                || line.StartsWith("中国工商银行信用卡历史明细", StringComparison.Ordinal)
                || line.StartsWith("卡号", StringComparison.Ordinal)
                || line.StartsWith("交易日期 ", StringComparison.Ordinal)
                || line.StartsWith("入账日期 ", StringComparison.Ordinal)
                || line.StartsWith("本页", StringComparison.Ordinal)
                || line.StartsWith("下单时间", StringComparison.Ordinal)
                || Regex.IsMatch(line, @"^第\s*\d+\s*页$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^共\s*\d+\s*页$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^共\s*\d+\s*页\s+第\s*\d+\s*页$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}.*(?:第\s*\d+\s*页|中国工商银行)", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+[A-F0-9]{8,}", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^中国工商银行\s+\d+", RegexOptions.CultureInvariant);
        }

        private static ICBCHistoryDetailRow ParseICBCHistoryDetailRow(string rowText)
        {
            var match = Regex.Match(
                rowText,
                @"^(?<date>\d{4}-\d{2}-\d{2})\s+(?<time>\d{2}:\d{2}:\d{2})\s*(?<rest>.+)$",
                RegexOptions.CultureInvariant);
            if (!match.Success)
                throw new MailParseException($"Parse ICBC history detail row fail: {rowText}");

            var tokens = Regex.Split(match.Groups["rest"].Value, @"\s+")
                .Where(token => !String.IsNullOrWhiteSpace(token))
                .ToList();
            if (tokens.Count < 6 || tokens.FindIndex(IsICBCHistoryDetailCurrencyToken) < 0)
                return ParseCompactICBCHistoryDetailRow(match.Groups["date"].Value, match.Groups["time"].Value, match.Groups["rest"].Value, rowText);

            if (IsICBCHistoryCreditCardRow(tokens))
                return ParseICBCHistoryCreditCardRow(match.Groups["date"].Value, match.Groups["time"].Value, tokens, rowText);

            return ParseICBCHistoryDebitAccountRow(match.Groups["date"].Value, match.Groups["time"].Value, tokens, rowText);
        }

        private static ICBCHistoryDetailRow ParseCompactICBCHistoryDetailRow(
            string dateText,
            string timeText,
            string restText,
            string rowText)
        {
            restText = NormalizeICBCHistoryText(restText);
            if (TryParseCompactICBCHistoryCreditCardRow(dateText, timeText, restText, rowText, out var creditRow))
                return creditRow;
            if (TryParseCompactICBCHistoryDebitAccountRow(dateText, timeText, restText, rowText, out var debitRow))
                return debitRow;

            throw new MailParseException($"Parse compact ICBC history detail row fail: {rowText}");
        }

        private static bool TryParseCompactICBCHistoryCreditCardRow(
            string dateText,
            string timeText,
            string restText,
            string rowText,
            out ICBCHistoryDetailRow row)
        {
            row = default!;
            var currencyPattern = GetICBCHistoryCurrencyPattern();
            var match = Regex.Match(
                restText,
                $@"^(?<cardNo>[\d*]+)(?<direction>[借贷])(?<transactionCurrency>{currencyPattern})(?<transactionAmount>\d[\d,]*\.\d{{2}})(?<postingCurrency>{currencyPattern})(?<postingAmount>\d[\d,]*\.\d{{2}})(?:(?<balance>[+-]?\d[\d,]*\.\d{{2}}))?(?<tail>.+)$",
                RegexOptions.CultureInvariant);
            if (!match.Success)
                return false;

            var tail = match.Groups["tail"].Value;
            if (!TrySplitICBCHistoryKnownPrefix(tail, GetICBCHistoryCreditSummaries(), out var summary, out var place))
            {
                if (tail.Length < 2)
                    throw new MailParseException($"Parse compact ICBC credit history row fail, unknown summary: {rowText}");

                summary = tail[..2];
                place = tail[2..];
            }

            var sign = match.Groups["direction"].Value switch
            {
                "借" => -1,
                "贷" => 1,
                _ => throw new MailParseException($"Parse compact ICBC credit history row fail, unknown direction: {rowText}")
            };
            row = new ICBCHistoryDetailRow(
                ParseICBCHistoryDetailDateTime(dateText, timeText),
                match.Groups["cardNo"].Value,
                new Currency(
                    Math.Abs(ParseInvariantDecimal(match.Groups["postingAmount"].Value)) * sign,
                    ParseICBCHistoryDetailCurrencyType(match.Groups["postingCurrency"].Value)),
                new Currency(
                    Math.Abs(ParseInvariantDecimal(match.Groups["transactionAmount"].Value)) * sign,
                    ParseICBCHistoryDetailCurrencyType(match.Groups["transactionCurrency"].Value)),
                match.Groups["balance"].Success ? ParseInvariantDecimal(match.Groups["balance"].Value) : null,
                summary,
                place,
                "",
                "",
                rowText,
                true);
            return true;
        }

        private static bool TryParseCompactICBCHistoryDebitAccountRow(
            string dateText,
            string timeText,
            string restText,
            string rowText,
            out ICBCHistoryDetailRow row)
        {
            row = default!;
            var currencyPattern = GetICBCHistoryCurrencyPattern();
            var match = Regex.Match(
                restText,
                $@"^(?<accountNo>[\d*]+)(?<deposit>活期|定期|通知)(?<serial>\d{{5}})(?<currency>{currencyPattern})(?<cash>钞|汇)(?<body>.+)$",
                RegexOptions.CultureInvariant);
            if (!match.Success)
                return false;

            var bodyMatch = Regex.Match(
                match.Groups["body"].Value,
                @"^(?<summary>.+?)(?<region>\d{4})(?<amount>[+-]\d[\d,]*\.\d{2})(?<balance>[+-]?\d[\d,]*\.\d{2})(?<tail>.*)$",
                RegexOptions.CultureInvariant);
            if (!bodyMatch.Success)
                throw new MailParseException($"Parse compact ICBC debit history row fail, missing amount/balance: {rowText}");

            var tail = ParseCompactICBCHistoryDetailTail(bodyMatch.Groups["tail"].Value);
            row = new ICBCHistoryDetailRow(
                ParseICBCHistoryDetailDateTime(dateText, timeText),
                match.Groups["accountNo"].Value,
                new Currency(
                    ParseInvariantDecimal(bodyMatch.Groups["amount"].Value),
                    ParseICBCHistoryDetailCurrencyType(match.Groups["currency"].Value)),
                null,
                ParseInvariantDecimal(bodyMatch.Groups["balance"].Value),
                bodyMatch.Groups["summary"].Value,
                tail.CounterpartyName,
                tail.CounterpartyAccount,
                tail.Channel,
                rowText,
                false);
            return true;
        }

        private static ICBCHistoryDetailRow ParseICBCHistoryDebitAccountRow(
            string dateText,
            string timeText,
            List<string> tokens,
            string rowText)
        {
            var currencyIndex = tokens.FindIndex(IsICBCHistoryDetailCurrencyToken);
            if (currencyIndex < 0)
                throw new MailParseException($"Parse ICBC history detail row fail, missing currency: {rowText}");

            var amountIndex = tokens.FindIndex(currencyIndex + 1, IsSignedMoneyToken);
            if (amountIndex < 0)
                throw new MailParseException($"Parse ICBC history detail row fail, missing signed amount: {rowText}");

            var balanceIndex = amountIndex + 1 < tokens.Count && IsMoneyToken(tokens[amountIndex + 1])
                ? amountIndex + 1
                : -1;
            var regionIndex = amountIndex > 0 && Regex.IsMatch(tokens[amountIndex - 1], @"^\d{4}$", RegexOptions.CultureInvariant)
                ? amountIndex - 1
                : -1;
            var summaryStart = currencyIndex + 1;
            if (summaryStart < tokens.Count && tokens[summaryStart] is "钞" or "汇")
                summaryStart++;
            var summaryEnd = regionIndex >= 0 ? regionIndex : amountIndex;
            if (summaryStart >= summaryEnd)
                throw new MailParseException($"Parse ICBC history detail row fail, missing summary: {rowText}");

            var tailStart = balanceIndex >= 0 ? balanceIndex + 1 : amountIndex + 1;
            var tailTokens = tailStart < tokens.Count ? tokens.Skip(tailStart).ToList() : [];
            var tail = ParseICBCHistoryDetailTail(tailTokens);
            var date = ParseICBCHistoryDetailDateTime(dateText, timeText);
            var currencyType = ParseICBCHistoryDetailCurrencyType(tokens[currencyIndex]);
            return new ICBCHistoryDetailRow(
                date,
                tokens[0],
                new Currency(ParseInvariantDecimal(tokens[amountIndex]), currencyType),
                null,
                balanceIndex >= 0 ? ParseInvariantDecimal(tokens[balanceIndex]) : null,
                String.Join("", tokens.Skip(summaryStart).Take(summaryEnd - summaryStart)),
                tail.CounterpartyName,
                tail.CounterpartyAccount,
                tail.Channel,
                rowText,
                false);
        }

        private static bool IsICBCHistoryCreditCardRow(List<string> tokens)
        {
            return tokens.Count >= 8
                && tokens[1] is "借" or "贷"
                && IsICBCHistoryDetailCurrencyToken(tokens[2])
                && IsMoneyToken(tokens[3])
                && IsICBCHistoryDetailCurrencyToken(tokens[4])
                && IsMoneyToken(tokens[5])
                && IsMoneyToken(tokens[6]);
        }

        private static ICBCHistoryDetailRow ParseICBCHistoryCreditCardRow(
            string dateText,
            string timeText,
            List<string> tokens,
            string rowText)
        {
            var sign = tokens[1] switch
            {
                "借" => -1,
                "贷" => 1,
                _ => throw new MailParseException($"Parse ICBC credit history row fail, unknown direction: {rowText}")
            };
            var transactionCurrency = ParseICBCHistoryDetailCurrencyType(tokens[2]);
            var postingCurrency = ParseICBCHistoryDetailCurrencyType(tokens[4]);
            var postingAmount = Math.Abs(ParseInvariantDecimal(tokens[5])) * sign;
            var transactionAmount = Math.Abs(ParseInvariantDecimal(tokens[3])) * sign;
            var place = String.Join("", tokens.Skip(8));
            return new ICBCHistoryDetailRow(
                ParseICBCHistoryDetailDateTime(dateText, timeText),
                tokens[0],
                new Currency(postingAmount, postingCurrency),
                new Currency(transactionAmount, transactionCurrency),
                ParseInvariantDecimal(tokens[6]),
                tokens[7],
                place,
                "",
                "",
                rowText,
                true);
        }

        private static ICBCHistoryDetailTail ParseICBCHistoryDetailTail(List<string> tailTokens)
        {
            if (tailTokens.Count == 0)
                return new ICBCHistoryDetailTail("", "", "");

            var accountIndex = -1;
            for (var i = tailTokens.Count - 1; i >= 0; i--)
            {
                if (IsCounterpartyAccountToken(tailTokens[i]))
                {
                    accountIndex = i;
                    break;
                }
            }

            if (accountIndex < 0)
                return new ICBCHistoryDetailTail(String.Join("", tailTokens), "", "");

            var name = NormalizeICBCHistoryEmptyText(String.Join("", tailTokens.Take(accountIndex)));
            var account = NormalizeICBCHistoryEmptyText(tailTokens[accountIndex]);
            var channel = String.Join("", tailTokens.Skip(accountIndex + 1));
            return new ICBCHistoryDetailTail(name, account, channel);
        }

        private static ICBCHistoryDetailTail ParseCompactICBCHistoryDetailTail(string tailText)
        {
            tailText = NormalizeICBCHistoryText(tailText);
            if (String.IsNullOrWhiteSpace(tailText))
                return new ICBCHistoryDetailTail("", "", "");

            var channel = "";
            foreach (var knownChannel in GetICBCHistoryChannels())
            {
                if (!tailText.EndsWith(knownChannel, StringComparison.Ordinal))
                    continue;

                channel = knownChannel;
                tailText = tailText[..^knownChannel.Length];
                break;
            }

            if (tailText.EndsWith("（空）", StringComparison.Ordinal))
            {
                var name = NormalizeICBCHistoryEmptyText(tailText[..^3]);
                return new ICBCHistoryDetailTail(name, "", channel);
            }

            var match = Regex.Match(tailText, @"^(?<name>.*?)(?<account>[\d*]{4,})$", RegexOptions.CultureInvariant);
            if (!match.Success)
                return new ICBCHistoryDetailTail(NormalizeICBCHistoryEmptyText(tailText), "", channel);

            return new ICBCHistoryDetailTail(
                NormalizeICBCHistoryEmptyText(match.Groups["name"].Value),
                NormalizeICBCHistoryEmptyText(match.Groups["account"].Value),
                channel);
        }

        private static bool TrySplitICBCHistoryKnownPrefix(
            string text,
            IEnumerable<string> knownPrefixes,
            out string prefix,
            out string rest)
        {
            foreach (var knownPrefix in knownPrefixes.OrderByDescending(item => item.Length))
            {
                if (!text.StartsWith(knownPrefix, StringComparison.Ordinal))
                    continue;

                prefix = knownPrefix;
                rest = text[knownPrefix.Length..];
                return true;
            }

            prefix = "";
            rest = "";
            return false;
        }

        private static bool IsCounterpartyAccountToken(string token)
        {
            return IsICBCHistoryEmptyToken(token)
                || token.Contains('*', StringComparison.Ordinal)
                || Regex.IsMatch(token, @"\d{4,}", RegexOptions.CultureInvariant);
        }

        private static bool IsICBCHistoryDetailCurrencyToken(string token)
        {
            return GetICBCHistoryCurrencyTokens().Contains(token);
        }

        private static string GetICBCHistoryCurrencyPattern()
        {
            return String.Join("|", GetICBCHistoryCurrencyTokens().Select(Regex.Escape));
        }

        private static string[] GetICBCHistoryCurrencyTokens()
        {
            return ["新加坡元", "人民币", "美元", "日元", "港币", "港元", "RMB", "CNY", "USD", "JPY", "SGD", "HKD"];
        }

        private static string[] GetICBCHistoryChannels()
        {
            return ["中间业务后台方式", "快捷支付", "手机银行", "网上银行", "批量业务", "柜面", "其他"];
        }

        private static string[] GetICBCHistoryCreditSummaries()
        {
            return ["自动还款", "消费", "退款", "还款", "转账", "转帐", "利息", "费用", "调整"];
        }

        private static CurrencyType ParseICBCHistoryDetailCurrencyType(string token)
        {
            return token switch
            {
                "人民币" or "RMB" or "CNY" => CurrencyType.RMB,
                "美元" or "USD" => CurrencyType.USD,
                "日元" or "JPY" => CurrencyType.JPY,
                "新加坡元" or "SGD" => CurrencyType.SGD,
                "港币" or "港元" or "HKD" => CurrencyType.HKD,
                _ => throw new MailParseException($"Unsupported ICBC history detail currency: {token}")
            };
        }

        private static bool IsSignedMoneyToken(string token)
        {
            return Regex.IsMatch(token, @"^[+-]\d[\d,]*(?:\.\d+)?$", RegexOptions.CultureInvariant);
        }

        private static bool IsMoneyToken(string token)
        {
            return Regex.IsMatch(token, @"^[+-]?\d[\d,]*(?:\.\d+)?$", RegexOptions.CultureInvariant);
        }

        private static decimal ParseInvariantDecimal(string token)
        {
            return Decimal.Parse(
                token.Replace(",", "", StringComparison.Ordinal),
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture);
        }

        private static DateTime ParseICBCHistoryDetailDateTime(string dateText, string timeText)
        {
            return DateTime.ParseExact(
                $"{dateText} {timeText}",
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture);
        }

        private static string BuildICBCHistoryDetailDestAccount(ICBCHistoryDetailRow row)
        {
            if (String.IsNullOrWhiteSpace(row.CounterpartyName))
                return row.CounterpartyAccount;
            if (String.IsNullOrWhiteSpace(row.CounterpartyAccount))
                return row.CounterpartyName;
            return $"{row.CounterpartyName} {row.CounterpartyAccount}";
        }

        private bool IsICBCHistoryDetailSelfTransfer(ICBCHistoryDetailRow row, Account cardAccount)
        {
            if (String.IsNullOrWhiteSpace(row.CounterpartyAccount))
                return false;

            var cardTail = GetICBCCardTail(row.CounterpartyAccount);
            if (String.IsNullOrWhiteSpace(cardTail))
                return false;

            try
            {
                var counterparty = database.GetAccountByTypeAndId(ICBCAccountType, cardTail);
                return !IsSameAccount(database.GetPostingAccount(counterparty), database.GetPostingAccount(cardAccount));
            }
            catch
            {
                return false;
            }
        }

        private static string BuildICBCHistoryDetailStatementKey(
            string accountName,
            DateTime startDate,
            DateTime endDate,
            string applicationNo)
        {
            return $"{accountName}|{startDate:yyyyMMdd}|{endDate:yyyyMMdd}|{applicationNo}";
        }

        private static (string AccountName, DateTime StartDate, DateTime EndDate, string ApplicationNo)? TryParseICBCHistoryDetailStatementKey(string statementKey)
        {
            var parts = statementKey.Split('|');
            if (parts.Length != 4)
                return null;
            if (!DateTime.TryParseExact(parts[1], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var startDate)
                || !DateTime.TryParseExact(parts[2], "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var endDate))
                return null;

            return (parts[0], startDate, endDate, parts[3]);
        }

        private static string BuildICBCHistoryDetailRowCode(string accountName, ICBCHistoryDetailRow row)
        {
            var canonical = String.Join(
                "|",
                accountName,
                row.PostingDate.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                row.Amount.t,
                row.Amount.v.ToString(CultureInfo.InvariantCulture),
                row.DescCurrency?.t.ToString() ?? "",
                row.DescCurrency?.v.ToString(CultureInfo.InvariantCulture) ?? "",
                row.Balance?.ToString(CultureInfo.InvariantCulture) ?? "",
                row.Summary,
                row.CounterpartyName,
                row.CounterpartyAccount,
                row.Channel);
            return ICBCHistoryDetailRowCodePrefix + ComputeSha256(Encoding.UTF8.GetBytes(canonical))[..24];
        }

        private static string ExtractICBCHistoryDetailRowCode(string source)
        {
            var match = Regex.Match(source, @"(?:^|;\s*)code=(?<code>[^;]+)", RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["code"].Value.Trim() : "";
        }

        private static string ParseICBCHistoryDetailApplicationNoFromFileName(string fileName)
        {
            var match = Regex.Match(fileName, @"申请单号[:：]?(?<applicationNo>\d+)", RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["applicationNo"].Value : "";
        }

        private static DateTime? ParseICBCHistoryDetailOrderTime(string text)
        {
            var match = ICBCHistoryDetailOrderTimeRegex.Match(text);
            if (!match.Success)
                return null;

            return DateTime.ParseExact(match.Groups["time"].Value, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static string GetICBCCardTail(string cardNo)
        {
            var digits = Regex.Replace(cardNo, @"\D", "");
            return digits.Length <= 4 ? digits : digits[^4..];
        }

        private static string NormalizeICBCHistoryText(string value)
        {
            return Regex.Replace(value, @"\s+", " ").Trim();
        }

        private static string TruncateForLog(string value, int maxLength)
        {
            return value.Length <= maxLength ? value : value[..maxLength];
        }

        private static string NormalizeICBCHistoryEmptyText(string value)
        {
            return IsICBCHistoryEmptyToken(value) ? "" : value;
        }

        private static bool IsICBCHistoryEmptyToken(string value)
        {
            return value is "（空）" or "(空)" or "空";
        }

        private static string GetICBCHistoryDetailDirectory(string? directory)
        {
            return String.IsNullOrWhiteSpace(directory)
                ? Directory.GetCurrentDirectory()
                : Path.GetFullPath(directory);
        }

        private static string SanitizeFileName(string fileName)
        {
            var invalidChars = Path.GetInvalidFileNameChars().ToHashSet();
            var sanitized = new string(fileName.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray());
            return String.IsNullOrWhiteSpace(sanitized) ? "ICBCHistoryDetail.pdf" : sanitized;
        }

        private static string GetUniqueFilePath(string directory, string fileName)
        {
            var path = Path.Combine(directory, fileName);
            if (!File.Exists(path))
                return path;

            var name = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            for (var index = 1; ; index++)
            {
                path = Path.Combine(directory, $"{name}.{index}{extension}");
                if (!File.Exists(path))
                    return path;
            }
        }

        private static string ComputeSha256(byte[] bytes)
        {
            return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        }

        private sealed record ICBCHistoryDetailAttachment(string FileName, byte[] Bytes);
        private sealed record ICBCHistoryDetailParsedAttachment(string FileName, string Text, string FileHash);
        private sealed record ICBCHistoryDetailParsedAttachments(
            List<ICBCHistoryDetailParsedAttachment> Statements,
            int SkippedUnreadableCount);
        private sealed record ICBCHistoryDetailTail(string CounterpartyName, string CounterpartyAccount, string Channel);
        private sealed record ICBCHistoryDetailRow(
            DateTime PostingDate,
            string AccountNo,
            Currency Amount,
            Currency? DescCurrency,
            decimal? Balance,
            string Summary,
            string CounterpartyName,
            string CounterpartyAccount,
            string Channel,
            string RawText,
            bool UseCardAccount);
        private sealed record ICBCHistoryDetailCandidate(
            int Index,
            DateTime PostingDate,
            Currency Amount,
            Currency? DescCurrency,
            string DestAccount,
            string Source,
            string RowCode);
        private sealed record ICBCHistoryDetailDebitCandidate(
            int Index,
            ICBCHistoryDetailRow Row,
            Record Record,
            string RowCode);
        private sealed record ICBCMonthlyStatementPeriod(
            StatementImport StatementImport,
            DateTime StartDate,
            DateTime EndDate);
        private sealed record ICBCHistoryDetailParsedStatement(
            string FileName,
            string StatementKey,
            string AccountName,
            bool IsCredit,
            DateTime StartDate,
            DateTime EndDate,
            DateTime ImportTime,
            Records Records,
            List<ICBCHistoryDetailCandidate> Candidates,
            List<ICBCHistoryDetailDebitCandidate> DebitCandidates,
            List<AccountInternalId> InternalCardNos);

        private sealed class ICBCHistoryDetailImportStats
        {
            private readonly Dictionary<string, int> ignored = new(StringComparer.OrdinalIgnoreCase);

            public int Imported { get; set; }
            public int Overlap { get; set; }
            public int MissingMonthlyRows { get; set; }

            public void Ignore(string reason, int count)
            {
                if (count <= 0)
                    return;

                ignored[reason] = ignored.TryGetValue(reason, out var current)
                    ? current + count
                    : count;
            }

            public string FormatIgnored()
            {
                if (ignored.Count == 0)
                    return "none";

                return String.Join(
                    ", ",
                    ignored
                        .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(item => $"{item.Key}:{item.Value}"));
            }
        }
    }
}
