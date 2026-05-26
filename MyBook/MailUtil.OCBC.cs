using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit.Search;
using MimeKit;
using UglyToad.PdfPig;

namespace MyBook
{
    // OCBC monthly combined e-statement discovery and parsing.
    // The PDF statement is the source of truth for OCBC balances.
    partial class MailUtil
    {
        private const StatementImportProvider OCBCProvider = StatementImportProvider.OCBCStatementMail;
        private const string OCBCMailSender = "documents@ocbc.com";
        private const string OCBCAccountType = "OCBC";
        private const string OCBCStatementPasswordConfigKey = "ocbc_statement_passwords";
        private static readonly Regex OCBCStatementSubjectRegex = new(
            @"^OCBC:\s*Your Combined e-Statement for (?<month>[A-Za-z]{3})\s+(?<year>\d{4})(?:\s+is attached)?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        public async Task FetchOCBCReports()
        {
            await FetchMonthlyStatements(OCBCProvider, "OCBC statement", FetchOCBCReport);
        }

        public async Task FetchOCBCReports(DateTime date)
        {
            await FetchOCBCReport(date);
        }

        private async Task<bool> FetchOCBCReport(DateTime date)
        {
            var statementMonth = FirstDayOfMonth(date);
            var subject = BuildOCBCStatementSubject(statementMonth);
            var message = await SearchOCBCStatementMail(statementMonth, subject);
            if (message is null)
                return false;

            ImportOCBCStatement(statementMonth, message);
            return true;
        }

        private void ImportOCBCStatement(DateTime statementMonth, MimeMessage message)
        {
            var accounts = GetOCBCStatementAccounts();
            var attachment = ReadOCBCStatementPdfAttachment(message);
            var text = ReadOCBCStatementPdfText(attachment);
            var parsed = ParseOCBCStatement(statementMonth, GetMailDate(message), text, accounts);
            if (database.IsStatementKeyImported(OCBCProvider, parsed.StatementKey))
            {
                Console.WriteLine($"Skip imported OCBC statement {parsed.StatementKey}");
                return;
            }

            var saved = database.SaveStatementRecordsOnce(
                OCBCProvider,
                parsed.ImportTime,
                parsed.Records,
                parsed.EndingBalances,
                parsed.StatementKey,
                parsed.BeginningBalances,
                internalCardNos: parsed.InternalCardNos);

            Console.WriteLine(saved
                ? $"Import OCBC statement {parsed.StatementKey}, records={parsed.Records.Count}"
                : $"Skip imported OCBC statement {parsed.StatementKey}");
        }

        private async Task<MimeMessage?> SearchOCBCStatementMail(DateTime statementMonth, string subject)
        {
            var query = SearchQuery.FromContains(OCBCMailSender)
                .And(SearchQuery.SentSince(statementMonth.Date))
                .And(SearchQuery.SentBefore(statementMonth.AddMonths(2).Date));
            var messages = await SearchMessages(
                $"OCBC statement {statementMonth:yyyy-MM}",
                query,
                message => IsOCBCStatementMail(message, statementMonth),
                GetMailDateTime);
            return messages.FirstOrDefault();
        }

        private static bool IsOCBCStatementMail(MimeMessage message, DateTime statementMonth)
        {
            if (!IsFrom(message, OCBCMailSender))
                return false;

            if (!TryParseOCBCStatementSubjectMonth(message.Subject ?? "", out var subjectMonth))
                return false;

            return subjectMonth == FirstDayOfMonth(statementMonth);
        }

        private static DateTime ParseOCBCStatementSubjectMonth(string subject)
        {
            if (!TryParseOCBCStatementSubjectMonth(subject, out var subjectMonth))
                throw new MailParseException($"Invalid OCBC statement subject: {subject}");

            return subjectMonth;
        }

        private static bool TryParseOCBCStatementSubjectMonth(string subject, out DateTime subjectMonth)
        {
            subjectMonth = default;
            var match = OCBCStatementSubjectRegex.Match(subject);
            if (!match.Success)
                return false;

            return DateTime.TryParseExact(
                $"{match.Groups["month"].Value} {match.Groups["year"].Value}",
                "MMM yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out subjectMonth);
        }

        private static string BuildOCBCStatementSubject(DateTime statementMonth)
        {
            return $"OCBC: Your Combined e-Statement for {statementMonth:MMM yyyy}";
        }

        private List<Account> GetOCBCStatementAccounts()
        {
            var accounts = database.GetAllAccounts()
                .Where(account => account.name.StartsWith($"{OCBCAccountType}_", StringComparison.OrdinalIgnoreCase))
                .OrderBy(account => account.name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (accounts.Count != 3)
                throw new InvalidOperationException($"OCBC statement import expects exactly 3 accounts, actual={accounts.Count}");

            return accounts;
        }

        private byte[] ReadOCBCStatementPdfAttachment(MimeMessage message)
        {
            var pdfAttachments = ReadMatchingAttachments(message, (attachment, fileName) =>
                fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)
                    ? ReadMimePartBytes(attachment)
                    : null);
            if (pdfAttachments.Count != 1)
                throw new MailParseException($"OCBC statement mail should contain exactly one PDF attachment: {message.Subject}");

            return pdfAttachments[0];
        }

        private string ReadOCBCStatementPdfText(byte[] pdfBytes)
        {
            var passwords = GetOCBCStatementPasswords();
            Exception? lastException = null;
            foreach (var password in passwords)
            {
                try
                {
                    using var document = PdfDocument.Open(pdfBytes, new ParsingOptions { Password = password });
                    var text = new StringBuilder();
                    foreach (var page in document.GetPages())
                    {
                        var lines = page.GetWords()
                            .GroupBy(word => Math.Round(word.BoundingBox.Bottom / 2.0) * 2.0)
                            .OrderByDescending(group => group.Key)
                            .Select(group => String.Join(" ", group
                                .OrderBy(word => word.BoundingBox.Left)
                                .Select(word => word.Text)));
                        foreach (var line in lines)
                            text.AppendLine(line);
                    }

                    return text.ToString();
                }
                catch (Exception e)
                {
                    lastException = e;
                }
            }

            throw new MailParseException($"Parse OCBC statement PDF fail, no configured password could open it: {lastException?.Message}");
        }

        private List<string> GetOCBCStatementPasswords()
        {
            var passwords = config.GetSection(OCBCStatementPasswordConfigKey)
                .GetChildren()
                .Select(section => section.Value ?? "")
                .ToList();
            if (passwords.Count == 0)
            {
                passwords = config.GetChildren()
                    .Where(section => String.Equals(
                        section.Key.Trim(),
                        OCBCStatementPasswordConfigKey,
                        StringComparison.OrdinalIgnoreCase))
                    .SelectMany(section => section.GetChildren())
                    .Select(section => section.Value ?? "")
                    .ToList();
            }

            if (passwords.Count == 0)
                throw new InvalidOperationException($"Missing config array: {OCBCStatementPasswordConfigKey}");

            return passwords;
        }

        private OCBCParsedStatement ParseOCBCStatement(
            DateTime statementMonth,
            DateTime importTime,
            string text,
            List<Account> accounts)
        {
            var normalizedText = NormalizeOCBCPdfText(text);
            var statementEndDate = ParseOCBCStatementEndDate(normalizedText, statementMonth);
            var statementKey = statementEndDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var lines = normalizedText.Split('\n')
                .Select(line => line.Trim())
                .Where(line => !String.IsNullOrWhiteSpace(line))
                .ToList();
            var words = Regex.Split(normalizedText, @"\s+")
                .Where(word => !String.IsNullOrWhiteSpace(word))
                .ToList();
            var accountInfos = accounts.Select(account => new OCBCAccountInfo(account, GetOCBCAccountTail(account))).ToList();
            var parsedBalances = ParseOCBCStatement(lines, words, accountInfos, statementEndDate, statementKey);
            return BuildOCBCStatementImport(importTime, statementKey, accountInfos, parsedBalances, []);
        }

        private OCBCParsedStatement BuildOCBCStatementImport(
            DateTime importTime,
            string statementKey,
            List<OCBCAccountInfo> accounts,
            OCBCParsedBalances parsedBalances,
            List<string> ownCounterpartyMarkers)
        {
            var beginningBalances = new List<AccountBalance>();
            var endingBalances = new List<AccountBalance>();
            var records = new Records();
            var now = DateTime.Now;
            var fxCounterparties = BuildOCBCFxCounterpartyMap(parsedBalances.Transactions);
            foreach (var transaction in parsedBalances.Transactions)
                records.Add(CreateOCBCRecord(transaction, statementKey, accounts, fxCounterparties, ownCounterpartyMarkers, now));

            foreach (var accountInfo in accounts)
            {
                var currency = GetOCBCAccountCurrency(accountInfo);
                var key = (accountInfo.Account.name, currency);
                var currentBalance = database.GetAccountBalance(accountInfo.Account, currency).v;
                var hasEnding = parsedBalances.Ending.TryGetValue(key, out var parsedEnding);
                var ending = hasEnding ? parsedEnding : 0;
                var beginning = parsedBalances.Beginning.TryGetValue(key, out var parsedBeginning)
                    ? parsedBeginning
                    : hasEnding && parsedBalances.PresentAccounts.Contains(accountInfo.Account.name)
                        ? ending
                        : currentBalance;

                beginningBalances.Add(new AccountBalance(accountInfo.Account, new Currency(beginning, currency)));
                endingBalances.Add(new AccountBalance(accountInfo.Account, new Currency(ending, currency)));
            }

            return new OCBCParsedStatement(importTime, statementKey, records, beginningBalances, endingBalances, parsedBalances.InternalCardNos);
        }

        private static OCBCParsedBalances ParseOCBCStatement(
            List<string> lines,
            List<string> words,
            List<OCBCAccountInfo> accounts,
            DateTime statementEndDate,
            string statementKey)
        {
            var beginning = new Dictionary<(string AccountName, CurrencyType Currency), decimal>();
            var ending = new Dictionary<(string AccountName, CurrencyType Currency), decimal>();
            var transactions = new List<OCBCParsedTransaction>();
            var internalCardNos = new List<AccountInternalId>();
            var presentAccounts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddOCBCSummaryEndingBalances(words, accounts, presentAccounts, ending, internalCardNos, statementKey);
            var sections = FindOCBCAccountSections(lines, accounts);
            foreach (var group in sections.GroupBy(section => section.Account.Account.name, StringComparer.OrdinalIgnoreCase))
            {
                var accountInfo = group.First().Account;
                internalCardNos.AddRange(group
                    .Select(section => section.AccountNumber)
                    .Where(accountNumber => !String.IsNullOrWhiteSpace(accountNumber))
                    .Select(accountNumber => new AccountInternalId
                    {
                        Account = accountInfo.Account,
                        cardNo = accountNumber,
                        currencyType = GetOCBCAccountCurrency(accountInfo),
                        sourceText = $"OCBC statement {statementKey}; section accountNumber={accountNumber}; account={accountInfo.Account.name}"
                    }));
                var sectionLines = group.SelectMany(section => section.Lines).ToList();
                var currency = GetOCBCAccountCurrency(accountInfo);
                var hasBeginningBalance = TryReadOCBCSectionBalance(sectionLines, "B/F", out var beginningBalance);
                var hasEndingBalance = TryReadOCBCSectionBalance(sectionLines, "C/F", out var endingBalance);
                if (!hasBeginningBalance && !hasEndingBalance)
                    continue;

                presentAccounts.Add(accountInfo.Account.name);
                if (!hasBeginningBalance)
                    throw new MailParseException($"Parse OCBC Statement Fail, missing beginning balance for {accountInfo.Account.name} in {statementKey}");

                if (!hasEndingBalance)
                    throw new MailParseException($"Parse OCBC Statement Fail, missing ending balance for {accountInfo.Account.name} in {statementKey}");

                var key = (accountInfo.Account.name, currency);
                if (beginning.ContainsKey(key))
                    throw new MailParseException($"Parse OCBC Statement Fail, duplicate account section for {accountInfo.Account.name} in {statementKey}");

                var sectionTransactions = ParseOCBCSectionTransactions(
                    new OCBCAccountSection(accountInfo, "", sectionLines),
                    statementEndDate,
                    statementKey,
                    beginningBalance,
                    endingBalance);
                transactions.AddRange(sectionTransactions);
                beginning[key] = beginningBalance;
                ending[key] = endingBalance;
            }

            foreach (var accountInfo in accounts.Where(account => presentAccounts.Contains(account.Account.name)))
            {
                var accountEnding = ending.Keys
                    .Any(key => String.Equals(key.AccountName, accountInfo.Account.name, StringComparison.OrdinalIgnoreCase));
                if (!accountEnding)
                    throw new MailParseException($"Parse OCBC Statement Fail, missing ending balance for {accountInfo.Account.name}");
            }

            return new OCBCParsedBalances(presentAccounts, beginning, ending, transactions, internalCardNos);
        }

        private static void AddOCBCSummaryEndingBalances(
            List<string> words,
            List<OCBCAccountInfo> accounts,
            HashSet<string> presentAccounts,
            Dictionary<(string AccountName, CurrencyType Currency), decimal> ending,
            List<AccountInternalId> internalCardNos,
            string statementKey)
        {
            var summaryEnd = FindOCBCSummaryEnd(words);
            for (var i = 0; i < summaryEnd; i++)
            {
                if (!TryParseOCBCSummaryAccountNumber(words[i], out var accountNumber))
                    continue;

                foreach (var account in accounts)
                {
                    if (!accountNumber.EndsWith(account.Tail, StringComparison.Ordinal))
                        continue;

                    var amountIndex = Enumerable.Range(i + 1, Math.Min(6, summaryEnd - i - 1))
                        .FirstOrDefault(index => IsOCBCAmountWord(words[index]));
                    if (amountIndex <= i)
                        throw new MailParseException($"Parse OCBC Statement Fail, missing summary balance for {account.Account.name}");

                    presentAccounts.Add(account.Account.name);
                    internalCardNos.Add(new AccountInternalId
                    {
                        Account = account.Account,
                        cardNo = accountNumber,
                        currencyType = GetOCBCAccountCurrency(account),
                        sourceText = $"OCBC statement {statementKey}; summary accountNumber={accountNumber}; account={account.Account.name}"
                    });
                    ending[(account.Account.name, GetOCBCAccountCurrency(account))] = ParseOCBCDecimal(words[amountIndex]);
                }
            }
        }

        private static bool TryParseOCBCSummaryAccountNumber(string word, out string accountNumber)
        {
            accountNumber = "";
            var token = word.Trim();
            // OCBC 摘要里的账号是纯数字标识。不要去掉标点后再匹配，否则 670.01 这类金额会被误读成账号。
            if (!Regex.IsMatch(token, @"^\d{8,}$", RegexOptions.CultureInvariant))
                return false;

            accountNumber = token;
            return true;
        }

        private static int FindOCBCSummaryEnd(List<string> words)
        {
            for (var i = 0; i + 1 < words.Count; i++)
            {
                if (String.Equals(words[i], "TRANSACTION", StringComparison.OrdinalIgnoreCase)
                    && String.Equals(words[i + 1], "CODE", StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return words.Count;
        }

        private static List<OCBCAccountSection> FindOCBCAccountSections(List<string> lines, List<OCBCAccountInfo> accounts)
        {
            var starts = new List<(int Index, OCBCAccountInfo Account, string AccountNumber)>();
            for (var i = 0; i < lines.Count; i++)
            {
                var match = Regex.Match(
                    lines[i],
                    @"\bAccount\s+No\.?\s+(?<accountNumber>\d+)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!match.Success)
                    continue;

                var accountNumber = Regex.Replace(match.Groups["accountNumber"].Value, @"\D", "");
                foreach (var account in accounts)
                {
                    if (accountNumber.EndsWith(account.Tail, StringComparison.Ordinal))
                        starts.Add((i, account, accountNumber));
                }
            }

            var sections = new List<OCBCAccountSection>();
            for (var i = 0; i < starts.Count; i++)
            {
                var end = i + 1 < starts.Count ? starts[i + 1].Index : lines.Count;
                sections.Add(new OCBCAccountSection(
                    starts[i].Account,
                    starts[i].AccountNumber,
                    lines.Skip(starts[i].Index).Take(end - starts[i].Index).ToList()));
            }

            return sections;
        }

        private static CurrencyType GetOCBCAccountCurrency(OCBCAccountInfo accountInfo)
        {
            return accountInfo.Tail == "1201" ? CurrencyType.USD : CurrencyType.SGD;
        }

        private static List<OCBCParsedTransaction> ParseOCBCSectionTransactions(
            OCBCAccountSection section,
            DateTime statementEndDate,
            string statementKey,
            decimal beginningBalance,
            decimal endingBalance)
        {
            var beginIndex = section.Lines.FindIndex(line => IsOCBCBalanceLine(line, "B/F"));
            var endIndex = section.Lines.FindIndex(line => IsOCBCBalanceLine(line, "C/F"));
            if (beginIndex < 0 || endIndex < 0 || endIndex <= beginIndex)
                throw new MailParseException($"Parse OCBC Statement Fail, invalid balance markers for {section.Account.Account.name} in {statementKey}");

            var currentBalance = beginningBalance;
            var pendingDescriptionLines = new List<string>();
            OCBCParsedTransaction? currentTransaction = null;
            var transactions = new List<OCBCParsedTransaction>();
            var transactionStartIndex = beginIndex + 1;
            if (!OCBCBalanceLineHasInlineAmount(section.Lines[beginIndex])
                && transactionStartIndex < endIndex
                && IsOCBCAmountWord(section.Lines[transactionStartIndex].Trim()))
            {
                transactionStartIndex++;
            }

            for (var i = transactionStartIndex; i < endIndex; i++)
            {
                var line = section.Lines[i].Trim();
                if (String.IsNullOrWhiteSpace(line))
                    continue;
                if (IsOCBCBoilerplateLine(line))
                    continue;

                if (TryParseOCBCTransactionLine(line, statementEndDate, out var transactionLine))
                {
                    var delta = transactionLine.Balance - currentBalance;
                    var unsignedAmount = transactionLine.Amount;
                    if (unsignedAmount < 0)
                        throw new MailParseException($"Parse OCBC Statement Fail, negative unsigned transaction amount: {line}");

                    decimal signedAmount;
                    if (delta == unsignedAmount)
                        signedAmount = unsignedAmount;
                    else if (delta == -unsignedAmount)
                        signedAmount = -unsignedAmount;
                    else
                        throw new MailParseException(
                            $"Parse OCBC Statement Fail, running balance mismatch for {section.Account.Account.name} in {statementKey}: previous={currentBalance}, lineAmount={unsignedAmount}, lineBalance={transactionLine.Balance}, line={line}");

                    var descriptionLines = new List<string>(pendingDescriptionLines);
                    pendingDescriptionLines.Clear();
                    if (!String.IsNullOrWhiteSpace(transactionLine.InlineDescription))
                        descriptionLines.Add(transactionLine.InlineDescription);

                    currentTransaction = new OCBCParsedTransaction(
                        section.Account,
                        GetOCBCAccountCurrency(section.Account),
                        transactionLine.TransactionDate,
                        transactionLine.ValueDate,
                        signedAmount,
                        transactionLine.Balance,
                        descriptionLines,
                        line);
                    transactions.Add(currentTransaction);
                    currentBalance = transactionLine.Balance;
                    continue;
                }

                if (IsOCBCPendingTransactionTitle(section.Lines, i, statementEndDate))
                {
                    pendingDescriptionLines.Add(line);
                    currentTransaction = null;
                    continue;
                }

                if (currentTransaction is not null)
                {
                    if (ShouldIgnoreOCBCContinuation(currentTransaction))
                        continue;

                    currentTransaction.DescriptionLines.Add(line);
                    continue;
                }

                pendingDescriptionLines.Add(line);
            }

            if (pendingDescriptionLines.Count > 0)
                throw new MailParseException(
                    $"Parse OCBC Statement Fail, description without transaction for {section.Account.Account.name} in {statementKey}: {String.Join(" / ", pendingDescriptionLines)}");

            if (currentBalance != endingBalance)
                throw new MailParseException(
                    $"Parse OCBC Statement Fail, section balance mismatch for {section.Account.Account.name} in {statementKey}: parsed={currentBalance}, ending={endingBalance}");

            return transactions;
        }

        private static bool IsOCBCPendingTransactionTitle(List<string> lines, int index, DateTime statementEndDate)
        {
            if (!IsKnownOCBCTransactionTitle(lines[index]))
                return false;

            for (var i = index + 1; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                if (String.IsNullOrWhiteSpace(line))
                    continue;

                return TryParseOCBCTransactionLine(line, statementEndDate, out _);
            }

            return false;
        }

        private static bool IsKnownOCBCTransactionTitle(string line)
        {
            var normalized = NormalizeOCBCTransactionText(line);
            return normalized is "FUND TRANSFER"
                or "INTEREST CREDIT"
                or "BONUS INTEREST"
                or "CALL A/C TT DEP"
                or "COMM/COMM IN LIEU"
                or "BANK CHARGES"
                or "TRAN CHARGE"
                or "CCY CONVERSION FEE"
                or "SERVICE CHARGE"
                or "DEBIT PURCHASE"
                or "MEPS RECEIPTS"
                or "CREDIT ADVICE";
        }

        private static bool ShouldIgnoreOCBCContinuation(OCBCParsedTransaction transaction)
        {
            var mainDescription = GetOCBCMainDescription(transaction);
            return mainDescription.Equals("INTEREST CREDIT", StringComparison.OrdinalIgnoreCase)
                || mainDescription.Equals("BONUS INTEREST", StringComparison.OrdinalIgnoreCase)
                || mainDescription.Equals("SERVICE CHARGE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCBoilerplateLine(string line)
        {
            var normalized = NormalizeOCBCTransactionText(line);
            if (normalized.Length == 1)
                return true;

            return normalized.StartsWith("Deposit Insurance Scheme", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Singapore dollar deposits", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("and separately insured", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Foreign currency deposits", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("OCBC Bank", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("65 Chulia Street", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("STATEMENT OF ACCOUNT", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Account No.", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Transaction Value", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Date Date Description", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(normalized, @"^\d{1,2}\s+[A-Z]{3}\s+\d{4}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                || normalized.Equals("LI JINGLUN", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(normalized, @"^[A-Za-z](?:\s+\d+)?$", RegexOptions.CultureInvariant)
                || (normalized.Length <= 5 && normalized.Contains(' ', StringComparison.Ordinal))
                || normalized.StartsWith("Page ", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("For enquiries", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Please check this statement", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("CHECK YOUR STATEMENT", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Total Withdrawals/Deposits", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Total Interest Paid", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("Average Balance", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryParseOCBCTransactionLine(string line, DateTime statementEndDate, out OCBCTransactionLine transactionLine)
        {
            transactionLine = new OCBCTransactionLine(default, default, "", 0, 0);
            var match = Regex.Match(
                line,
                @"^(?<td>\d{1,2})\s+(?<tm>[A-Za-z]{3})\s+(?<vd>\d{1,2})\s+(?<vm>[A-Za-z]{3})(?:\s+(?<desc>.*?))?\s+(?<amount>\(?-?\d[\d,]*(?:\.\d+)?\)?)\s+(?<balance>\(?-?\d[\d,]*(?:\.\d+)?\)?)\s*$",
                RegexOptions.CultureInvariant);
            if (!match.Success)
                return false;

            transactionLine = new OCBCTransactionLine(
                ParseOCBCLineDate(match.Groups["td"].Value, match.Groups["tm"].Value, statementEndDate),
                ParseOCBCLineDate(match.Groups["vd"].Value, match.Groups["vm"].Value, statementEndDate),
                NormalizeOCBCTransactionText(match.Groups["desc"].Value),
                ParseOCBCDecimal(match.Groups["amount"].Value),
                ParseOCBCDecimal(match.Groups["balance"].Value));
            return true;
        }

        private static DateTime ParseOCBCLineDate(string dayText, string monthText, DateTime statementEndDate)
        {
            if (!DateTime.TryParseExact(
                    monthText,
                    "MMM",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var parsedMonth))
            {
                throw new MailParseException($"Invalid OCBC transaction month: {monthText}");
            }

            var day = Int32.Parse(dayText, CultureInfo.InvariantCulture);
            var candidate = new DateTime(statementEndDate.Year, parsedMonth.Month, day);
            if (candidate < statementEndDate.AddMonths(-11))
                candidate = candidate.AddYears(1);
            else if (candidate > statementEndDate.AddMonths(1))
                candidate = candidate.AddYears(-1);

            return candidate.Date;
        }

        private static bool TryReadOCBCSectionBalance(List<string> lines, string balanceKind, out decimal balance)
        {
            balance = 0;
            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var match = Regex.Match(
                    line,
                    $@"^\s*BALANCE\s+{Regex.Escape(balanceKind)}(?:\s+(?<amount>\(?-?\d[\d,]*(?:\.\d+)?\)?))?\s*$",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!match.Success)
                    continue;

                if (match.Groups["amount"].Success)
                {
                    balance = ParseOCBCDecimal(match.Groups["amount"].Value);
                    return true;
                }

                var nextAmountLine = lines
                    .Skip(i + 1)
                    .FirstOrDefault(candidate => !String.IsNullOrWhiteSpace(candidate));
                if (nextAmountLine is not null && IsOCBCAmountWord(nextAmountLine))
                {
                    balance = ParseOCBCDecimal(nextAmountLine);
                    return true;
                }
            }

            return false;
        }

        private static bool IsOCBCBalanceLine(string line, string balanceKind)
        {
            return Regex.IsMatch(
                line,
                $@"^\s*BALANCE\s+{Regex.Escape(balanceKind)}(?:\s+\(?-?\d[\d,]*(?:\.\d+)?\)?)?\s*$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool OCBCBalanceLineHasInlineAmount(string line)
        {
            return Regex.IsMatch(
                line,
                @"^\s*BALANCE\s+(?:B/F|C/F)\s+\(?-?\d[\d,]*(?:\.\d+)?\)?\s*$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsOCBCAmountWord(string value)
        {
            return Regex.IsMatch(value, @"^\(?-?\d[\d,]*(?:\.\d+)?\)?$", RegexOptions.CultureInvariant);
        }

        private static decimal ParseOCBCDecimal(string amountText)
        {
            var normalized = amountText.Trim();
            var negative = normalized.StartsWith("(", StringComparison.Ordinal) && normalized.EndsWith(")", StringComparison.Ordinal);
            normalized = normalized.Trim('(', ')').Replace(",", "");
            var amount = Decimal.Parse(normalized, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            return negative ? -Math.Abs(amount) : amount;
        }

        private static Dictionary<OCBCParsedTransaction, string> BuildOCBCFxCounterpartyMap(List<OCBCParsedTransaction> transactions)
        {
            var result = new Dictionary<OCBCParsedTransaction, string>();
            foreach (var group in transactions
                         .Where(IsOCBCFxTransaction)
                         .GroupBy(transaction => transaction.Reference, StringComparer.OrdinalIgnoreCase))
            {
                if (String.IsNullOrWhiteSpace(group.Key))
                    throw new MailParseException($"Parse OCBC Statement Fail, FX transaction missing reference: {group.First().DescriptionText}");

                var list = group.ToList();
                if (list.Count != 2 || list.Select(transaction => transaction.Currency).Distinct().Count() != 2)
                    throw new MailParseException($"Parse OCBC Statement Fail, unpaired FX transaction reference {group.Key}: count={list.Count}");

                foreach (var transaction in list)
                    result[transaction] = list.First(other => !ReferenceEquals(other, transaction)).AccountInfo.Account.name;
            }

            return result;
        }

        private Record CreateOCBCRecord(
            OCBCParsedTransaction transaction,
            string statementKey,
            List<OCBCAccountInfo> accounts,
            Dictionary<OCBCParsedTransaction, string> fxCounterparties,
            List<string> ownCounterpartyMarkers,
            DateTime updateTime)
        {
            var classification = ClassifyOCBCTransaction(transaction, statementKey, accounts, fxCounterparties, ownCounterpartyMarkers);
            return new Record
            {
                Account = transaction.AccountInfo.Account,
                date = transaction.ValueDate,
                updateTime = updateTime,
                DestAccount = classification.DestAccount,
                isInternal = classification.IsInternal,
                Source = LimitOCBCRecordText(BuildOCBCSource(statementKey, transaction)),
                Reason = classification.Reason,
                v = transaction.Amount,
                t = transaction.Currency
            };
        }

        private OCBCTransactionClassification ClassifyOCBCTransaction(
            OCBCParsedTransaction transaction,
            string statementKey,
            List<OCBCAccountInfo> accounts,
            Dictionary<OCBCParsedTransaction, string> fxCounterparties,
            List<string> ownCounterpartyMarkers)
        {
            if (IsOCBCFeeTransaction(transaction))
                return new OCBCTransactionClassification("手续费", ExtractOCBCCounterparty(transaction, accounts), false);

            if (IsOCBCInterestTransaction(transaction))
                return new OCBCTransactionClassification("利息", "OCBC", false);

            if (IsOCBCFxTransaction(transaction))
            {
                var destAccount = fxCounterparties.TryGetValue(transaction, out var pairedAccount)
                    ? pairedAccount
                    : ExtractOCBCFxDescription(transaction);
                return new OCBCTransactionClassification("换汇", destAccount, true);
            }

            var counterparty = ExtractOCBCCounterparty(transaction, accounts);
            var isInternal = IsOCBCOwnCounterparty(counterparty, transaction.DescriptionText, ownCounterpartyMarkers);
            var internalCounterparty = database.FindAccountByInternalCardNoText(
                null,
                $"{BuildOCBCSource(statementKey, transaction)}; raw={transaction.RawLine}",
                counterparty,
                transaction.DescriptionText);
            if (internalCounterparty is not null && !IsSameAccount(database.GetPostingAccount(internalCounterparty), transaction.AccountInfo.Account))
            {
                counterparty = internalCounterparty.name;
                isInternal = true;
            }

            if (IsOCBCTransferTransaction(transaction))
                return new OCBCTransactionClassification("转账", counterparty, isInternal);

            if (IsOCBCTelegraphicDepositTransaction(transaction))
                return new OCBCTransactionClassification("电汇入账", counterparty, isInternal);

            if (IsOCBCPurchaseTransaction(transaction))
                return new OCBCTransactionClassification("消费", counterparty, isInternal);

            if (IsOCBCReceiptTransaction(transaction))
                return new OCBCTransactionClassification("转账入账", counterparty, isInternal);

            if (IsOCBCCreditAdviceTransaction(transaction))
                return new OCBCTransactionClassification(
                    transaction.DescriptionText.Contains("Welcome", StringComparison.OrdinalIgnoreCase) ? "奖励" : "入账",
                    counterparty,
                    isInternal);

            return new OCBCTransactionClassification(GetOCBCMainDescription(transaction), counterparty, isInternal);
        }

        private static string BuildOCBCSource(string statementKey, OCBCParsedTransaction transaction)
        {
            return $"OCBC statement {statementKey}/{transaction.AccountInfo.Account.name}/{transaction.Currency}: transactionDate={transaction.TransactionDate:yyyy-MM-dd}; valueDate={transaction.ValueDate:yyyy-MM-dd}; {transaction.DescriptionText}";
        }

        private static string ExtractOCBCCounterparty(OCBCParsedTransaction transaction, List<OCBCAccountInfo> accounts)
        {
            var lines = transaction.DescriptionLines
                .Select(NormalizeOCBCTransactionText)
                .Where(line => !String.IsNullOrWhiteSpace(line))
                .ToList();
            var ownAccount = ExtractOCBCOwnAccountCounterparty(lines, accounts, transaction.AccountInfo);
            if (!String.IsNullOrWhiteSpace(ownAccount))
                return ownAccount;

            var toLine = lines.FirstOrDefault(line => line.StartsWith("to ", StringComparison.OrdinalIgnoreCase));
            if (!String.IsNullOrWhiteSpace(toLine))
            {
                var reference = lines.FirstOrDefault(line => line.StartsWith("OTHR - ", StringComparison.OrdinalIgnoreCase));
                return String.IsNullOrWhiteSpace(reference)
                    ? toLine[3..].Trim()
                    : $"{toLine[3..].Trim()} / {reference[7..].Trim()}";
            }

            var wiseLine = lines.FirstOrDefault(line => line.Contains("Wise", StringComparison.OrdinalIgnoreCase));
            if (!String.IsNullOrWhiteSpace(wiseLine))
                return ExtractOCBCMerchantLine(wiseLine, @"xx-\d+\s+Wise");

            var paypalLine = lines.FirstOrDefault(line => line.Contains("PAYPAL", StringComparison.OrdinalIgnoreCase));
            if (!String.IsNullOrWhiteSpace(paypalLine))
                return ExtractOCBCMerchantLine(paypalLine, @"xx-\d+\s+PAYPAL.*");

            var namedLines = lines
                .Where(line => !IsOCBCDescriptionMetadata(line))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (namedLines.Count > 0)
                return String.Join(" / ", namedLines);

            return GetOCBCMainDescription(transaction);
        }

        private static string ExtractOCBCMerchantLine(string line, string pattern)
        {
            var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? NormalizeOCBCTransactionText(match.Value) : line;
        }

        private static string ExtractOCBCOwnAccountCounterparty(List<string> lines, List<OCBCAccountInfo> accounts, OCBCAccountInfo currentAccount)
        {
            foreach (var line in lines)
            {
                var accountNumber = Regex.Replace(line, @"\D", "");
                if (accountNumber.Length < 4)
                    continue;

                var account = accounts
                    .FirstOrDefault(item => accountNumber.EndsWith(item.Tail, StringComparison.Ordinal)
                        && !String.Equals(item.Account.name, currentAccount.Account.name, StringComparison.OrdinalIgnoreCase));
                if (account is not null)
                    return account.Account.name;
            }

            return "";
        }

        private static string ExtractOCBCFxDescription(OCBCParsedTransaction transaction)
        {
            var fxLine = transaction.DescriptionLines
                .FirstOrDefault(line => Regex.IsMatch(line, @"\b[A-Z]{3}\s+to\s+[A-Z]{3}\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
            return String.IsNullOrWhiteSpace(fxLine)
                ? GetOCBCMainDescription(transaction)
                : NormalizeOCBCTransactionText(fxLine);
        }

        private static bool IsOCBCDescriptionMetadata(string line)
        {
            if (IsKnownOCBCTransactionTitle(line))
                return true;

            return line.StartsWith("*", StringComparison.Ordinal)
                || line.StartsWith("Sys ", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("via ", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("OTHR - ", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(line, @"^\d+$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^\d{1,2}/\d{1,2}/\d{2,4}(?:\s+[A-Za-z])?$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"^[A-Z]{3}\s+\d[\d,]*(?:\.\d+)?$", RegexOptions.CultureInvariant)
                || Regex.IsMatch(line, @"\b[A-Z]{3}\s+to\s+[A-Z]{3}\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                || line.Contains("FX Transaction", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(line, @"^[A-Z0-9]{8,}$", RegexOptions.CultureInvariant);
        }

        private static bool IsOCBCOwnCounterparty(string counterparty, string descriptionText, List<string> ownCounterpartyMarkers)
        {
            var normalizedCounterparty = NormalizeOCBCNameForComparison(counterparty + " " + descriptionText);
            if (normalizedCounterparty.Contains("WISE", StringComparison.Ordinal)
                || normalizedCounterparty.Contains("OWNACCOUNT", StringComparison.Ordinal)
                || Regex.IsMatch(counterparty, @"^OCBC_\d{4}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return true;

            return ownCounterpartyMarkers
                .Select(NormalizeOCBCNameForComparison)
                .Where(marker => !String.IsNullOrWhiteSpace(marker))
                .Any(marker => normalizedCounterparty.Contains(marker, StringComparison.Ordinal));
        }

        private static string NormalizeOCBCNameForComparison(string value)
        {
            return Regex.Replace(value, @"\s+", "").ToUpperInvariant();
        }

        private static bool IsOCBCFeeTransaction(OCBCParsedTransaction transaction)
        {
            var text = transaction.DescriptionText;
            return text.Contains("COMM/COMM IN LIEU", StringComparison.OrdinalIgnoreCase)
                || text.Contains("BANK CHARGES", StringComparison.OrdinalIgnoreCase)
                || text.Contains("TRAN CHARGE", StringComparison.OrdinalIgnoreCase)
                || text.Contains("CCY CONVERSION FEE", StringComparison.OrdinalIgnoreCase)
                || text.Contains("SERVICE CHARGE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCInterestTransaction(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionText.Contains("INTEREST CREDIT", StringComparison.OrdinalIgnoreCase)
                || transaction.DescriptionText.Contains("BONUS INTEREST", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCPurchaseTransaction(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionText.Contains("DEBIT PURCHASE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCReceiptTransaction(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionText.Contains("MEPS RECEIPTS", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCCreditAdviceTransaction(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionText.Contains("CREDIT ADVICE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCFxTransaction(OCBCParsedTransaction transaction)
        {
            var text = transaction.DescriptionText;
            return text.Contains("FX Transaction", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(text, @"\b[A-Z]{3}\s+to\s+[A-Z]{3}\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsOCBCTransferTransaction(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionText.Contains("FUND TRANSFER", StringComparison.OrdinalIgnoreCase)
                || transaction.DescriptionText.Contains("TRANSFER", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOCBCTelegraphicDepositTransaction(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionText.Contains("CALL A/C TT DEP", StringComparison.OrdinalIgnoreCase)
                || transaction.DescriptionText.Contains("TT DEP", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetOCBCMainDescription(OCBCParsedTransaction transaction)
        {
            return transaction.DescriptionLines.Count == 0
                ? "未分类"
                : NormalizeOCBCTransactionText(transaction.DescriptionLines[0]);
        }

        private static string LimitOCBCRecordText(string text)
        {
            const int maxLength = 1024;
            var normalized = NormalizeOCBCTransactionText(text);
            return normalized.Length <= maxLength ? normalized : normalized[..maxLength];
        }

        private static string NormalizeOCBCTransactionText(string text)
        {
            return Regex.Replace(text ?? "", @"\s+", " ").Trim();
        }

        private static string ExtractOCBCReference(IEnumerable<string> lines)
        {
            foreach (var line in lines.Reverse().Select(NormalizeOCBCTransactionText))
            {
                var othr = Regex.Match(line, @"\bOTHR\s*-\s*(?<reference>[A-Z0-9]{8,})\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (othr.Success)
                    return othr.Groups["reference"].Value;

                var reference = Regex.Match(line, @"\b(?<reference>(?:FTS)?[A-Z0-9]*\d[A-Z0-9]{7,})\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (reference.Success && !line.StartsWith("*", StringComparison.Ordinal))
                    return reference.Groups["reference"].Value;
            }

            return "";
        }

        private static DateTime ParseOCBCStatementEndDate(string text, DateTime statementMonth)
        {
            var datePatterns = new[]
            {
                @"Statement\s+Date\s*:?\s*(?<date>\d{1,2}\s+[A-Za-z]{3}\s+\d{4})",
                @"Statement\s+Period\s*:?.*?(?:to|-)\s*(?<date>\d{1,2}\s+[A-Za-z]{3}\s+\d{4})",
                @"From\s+\d{1,2}\s+[A-Za-z]{3}\s+\d{4}\s+(?:to|-)\s*(?<date>\d{1,2}\s+[A-Za-z]{3}\s+\d{4})"
            };
            foreach (var pattern in datePatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);
                if (match.Success
                    && DateTime.TryParseExact(match.Groups["date"].Value.Trim(), "d MMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                {
                    return date.Date;
                }
            }

            return statementMonth.AddMonths(1).AddDays(-1);
        }

        private static string NormalizeOCBCPdfText(string text)
        {
            var normalized = text
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
            normalized = Regex.Replace(normalized, @"[ \t]+", " ");
            normalized = Regex.Replace(normalized, @"\n{2,}", "\n");
            return normalized.Trim();
        }

        private static string GetOCBCAccountTail(Account account)
        {
            var separator = account.name.LastIndexOf('_');
            var tail = separator >= 0 ? account.name[(separator + 1)..] : account.name;
            if (!Regex.IsMatch(tail, @"^\d{4}$", RegexOptions.CultureInvariant))
                throw new InvalidOperationException($"Invalid OCBC account name, expected OCBC_1234: {account.name}");

            return tail;
        }

        private sealed record OCBCAccountInfo(Account Account, string Tail);

        private sealed record OCBCAccountSection(OCBCAccountInfo Account, string AccountNumber, List<string> Lines);

        private sealed record OCBCParsedBalances(
            HashSet<string> PresentAccounts,
            Dictionary<(string AccountName, CurrencyType Currency), decimal> Beginning,
            Dictionary<(string AccountName, CurrencyType Currency), decimal> Ending,
            List<OCBCParsedTransaction> Transactions,
            List<AccountInternalId> InternalCardNos);

        private sealed record OCBCTransactionLine(
            DateTime TransactionDate,
            DateTime ValueDate,
            string InlineDescription,
            decimal Amount,
            decimal Balance);

        private sealed record OCBCTransactionClassification(
            string Reason,
            string DestAccount,
            bool IsInternal);

        private sealed class OCBCParsedTransaction
        {
            public OCBCParsedTransaction(
                OCBCAccountInfo accountInfo,
                CurrencyType currency,
                DateTime transactionDate,
                DateTime valueDate,
                decimal amount,
                decimal runningBalance,
                List<string> descriptionLines,
                string rawLine)
            {
                AccountInfo = accountInfo;
                Currency = currency;
                TransactionDate = transactionDate;
                ValueDate = valueDate;
                Amount = amount;
                RunningBalance = runningBalance;
                DescriptionLines = descriptionLines;
                RawLine = rawLine;
            }

            public OCBCAccountInfo AccountInfo { get; }
            public CurrencyType Currency { get; }
            public DateTime TransactionDate { get; }
            public DateTime ValueDate { get; }
            public decimal Amount { get; }
            public decimal RunningBalance { get; }
            public List<string> DescriptionLines { get; }
            public string RawLine { get; }
            public string DescriptionText => NormalizeOCBCTransactionText(String.Join(" / ", DescriptionLines));
            public string Reference => ExtractOCBCReference(DescriptionLines);
        }

        private sealed record OCBCParsedStatement(
            DateTime ImportTime,
            string StatementKey,
            Records Records,
            List<AccountBalance> BeginningBalances,
            List<AccountBalance> EndingBalances,
            List<AccountInternalId> InternalCardNos);
    }
}
