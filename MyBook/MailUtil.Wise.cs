using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using MailKit.Search;
using MimeKit;

namespace MyBook
{
    // Wise monthly XML statement discovery and parsing. XML statements contain exact
    // opening/closing balances, so imports validate Records against statement balances.
    partial class MailUtil
    {
        private const StatementImportProvider WiseProvider = StatementImportProvider.WiseMail;
        private const string WiseMailSender = "noreply@wise.com";
        private const string WiseAccountName = "WISE";
        private static readonly XNamespace WiseCamtNamespace = "urn:iso:std:iso:20022:tech:xsd:camt.053.001.10";
        private static readonly Regex WiseStatementFileNameRegex = new(
            @"^statement_(?<statementId>[^_]+)_(?<currency>[A-Z]{3})_(?<from>\d{4}-\d{2}-\d{2})_(?<to>\d{4}-\d{2}-\d{2})\.xml$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex WiseConvertedRegex = new(
            @"^Converted\s+(?<fromAmount>[\d,]+(?:\.\d+)?)\s+(?<fromCurrency>[A-Z]{3})\s+to\s+(?<toAmount>[\d,]+(?:\.\d+)?)\s+(?<toCurrency>[A-Z]{3})(?:\s+\(fee:\s+(?<feeAmount>[\d,]+(?:\.\d+)?)\s+(?<feeCurrency>[A-Z]{3})\))?$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex WiseSentMoneyRegex = new(
            @"^Sent money to\s+(?<counterparty>.+?)(?:\s+\(fee:\s+(?<feeAmount>[\d,]+(?:\.\d+)?)\s+(?<feeCurrency>[A-Z]{3})\))?$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex WiseReceivedMoneyRegex = new(
            @"^Received money from\s+(?<counterparty>.+?)(?:\s+with reference.*)?$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex WiseCardTransactionRegex = new(
            @"^Card transaction of\s+(?<amount>[\d,]+(?:\.\d+)?)\s+(?<currency>[A-Z]{3})\s+issued by\s+(?<merchant>.+)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex WiseDirectDebitRegex = new(
            @"^Paid to\s+(?<counterparty>.+)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public async Task FetchWiseReports()
        {
            await FetchMonthlyStatements(WiseProvider, "Wise statement", FetchWiseReport);
        }

        public async Task FetchWiseReports(DateTime date)
        {
            await FetchWiseReport(date);
        }

        public void DebugFetchLocalWiseReports(string? directory = null)
        {
            var files = FindLocalWiseStatementXmlFiles(directory);
            var attachments = files
                .Select(file => new InMemoryWiseStatementAttachment(
                    Path.GetFileName(file),
                    null,
                    File.ReadAllBytes(file)))
                .ToList();
            var parsed = ParseWiseStatementAttachments(attachments)
                .OrderBy(statement => statement.StatementStartDate)
                .ThenBy(statement => statement.StatementEndDate)
                .ToList();
            if (parsed.Count != 1)
                throw new MailParseException($"Local Wise XML test expects exactly one statement group, actual={parsed.Count}");

            var saved = SaveWiseParsedStatement(parsed[0]);
            PrintWiseParsedStatementSummary("Local Wise XML", parsed[0], saved);
        }

        private async Task<bool> FetchWiseReport(DateTime date)
        {
            var statementMonth = FirstDayOfMonth(date);
            var messages = await SearchWiseStatementMails(statementMonth);
            var attachments = messages
                .SelectMany(ReadWiseStatementAttachments)
                .ToList();
            if (attachments.Count == 0)
                return false;

            var statements = ParseWiseStatementAttachments(attachments)
                .Where(statement => FirstDayOfMonth(statement.StatementEndDate) == statementMonth)
                .OrderBy(statement => statement.StatementEndDate)
                .ToList();
            if (statements.Count == 0)
                return false;
            if (statements.Count > 1)
                throw new MailParseException($"Found multiple Wise XML statements for {statementMonth:yyyy-MM}");

            var saved = SaveWiseParsedStatement(statements[0]);
            PrintWiseParsedStatementSummary($"Wise XML {statementMonth:yyyy-MM}", statements[0], saved);
            return true;
        }

        private async Task<List<MimeMessage>> SearchWiseStatementMails(DateTime statementMonth)
        {
            var query = SearchQuery.FromContains(WiseMailSender)
                .And(SearchQuery.SentSince(statementMonth.Date))
                .And(SearchQuery.SentBefore(statementMonth.AddMonths(2).Date));
            return await SearchMessages(
                $"Wise XML statement {statementMonth:yyyy-MM}",
                query,
                IsWiseStatementMail,
                GetMailDateTime);
        }

        private static bool IsWiseStatementMail(MimeMessage message)
        {
            return IsFrom(message, WiseMailSender)
                && HasMatchingAttachment(message, (_, fileName) => IsWiseXmlStatementAttachment(fileName));
        }

        private static bool IsWiseXmlStatementAttachment(string fileName)
        {
            return fileName.StartsWith("statement_", StringComparison.OrdinalIgnoreCase)
                && Path.GetExtension(fileName).Equals(".xml", StringComparison.OrdinalIgnoreCase);
        }

        private static List<InMemoryWiseStatementAttachment> ReadWiseStatementAttachments(MimeMessage message)
        {
            return ReadMatchingAttachments(message, (attachment, fileName) =>
                IsWiseXmlStatementAttachment(fileName)
                    ? new InMemoryWiseStatementAttachment(
                        fileName,
                        GetMailDate(message),
                        ReadMimePartBytes(attachment))
                    : null);
        }

        private List<string> FindLocalWiseStatementXmlFiles(string? directory)
        {
            var directories = new[]
                {
                    directory,
                    Path.Combine(Directory.GetCurrentDirectory(), "archive"),
                    Path.Combine(Directory.GetCurrentDirectory(), "..", "archive"),
                    Directory.GetCurrentDirectory()
                }
                .Where(value => !String.IsNullOrWhiteSpace(value))
                .Select(value => Path.GetFullPath(value!))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(Directory.Exists)
                .ToList();

            var files = directories
                .SelectMany(dir => Directory.GetFiles(dir, "statement_*.xml"))
                .Where(file => WiseStatementFileNameRegex.IsMatch(Path.GetFileName(file)))
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (files.Count == 0)
                throw new FileNotFoundException("No local Wise XML statement found.");

            return files;
        }

        private List<WiseParsedStatement> ParseWiseStatementAttachments(List<InMemoryWiseStatementAttachment> attachments)
        {
            var account = database.GetAccountByName(WiseAccountName);
            var currencyStatements = attachments
                .Select(attachment => ParseWiseCurrencyStatement(attachment, account))
                .ToList();
            if (currencyStatements.Count == 0)
                throw new MailParseException("Wise XML statement contains no currency statements.");

            return currencyStatements
                .GroupBy(statement => new
                {
                    statement.StatementStartDate,
                    statement.StatementEndDate,
                    statement.CustomerId
                })
                .Select(group => BuildWiseParsedStatement(group.ToList()))
                .ToList();
        }

        private WiseParsedStatement BuildWiseParsedStatement(List<WiseCurrencyStatement> statements)
        {
            var first = statements[0];
            var duplicateCurrency = statements
                .GroupBy(statement => statement.Currency)
                .FirstOrDefault(group => group.Count() > 1);
            if (duplicateCurrency is not null)
                throw new MailParseException($"Duplicate Wise XML currency statement: {duplicateCurrency.Key}");

            var statementKey = first.StatementEndDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var importTime = statements
                .Select(statement => statement.ImportTime)
                .Where(time => time.HasValue)
                .DefaultIfEmpty(first.StatementEndDate)
                .Max()!.Value
                .Date;

            return new WiseParsedStatement(
                first.StatementStartDate,
                first.StatementEndDate,
                importTime,
                statementKey,
                first.Account,
                statements.SelectMany(statement => statement.Records)
                    .OrderBy(record => record.date)
                    .ThenBy(record => record.Source, StringComparer.Ordinal)
                    .ToRecords(),
                statements.Select(statement => new AccountBalance(first.Account, statement.BeginningBalance)).ToList(),
                statements.Select(statement => new AccountBalance(first.Account, statement.EndingBalance)).ToList(),
                statements.SelectMany(statement => statement.InternalCardNos).ToList());
        }

        private WiseCurrencyStatement ParseWiseCurrencyStatement(
            InMemoryWiseStatementAttachment attachment,
            Account account)
        {
            var document = XDocument.Parse(ReadWiseXmlText(attachment.Content), LoadOptions.None);
            var statement = RequireElement(
                RequireElement(document.Root, WiseCamtNamespace + "BkToCstmrStmt", attachment.FileName),
                WiseCamtNamespace + "Stmt",
                attachment.FileName);
            var statementId = RequireText(statement, WiseCamtNamespace + "Id", attachment.FileName);
            var customerId = OptionalText(
                statement,
                WiseCamtNamespace + "Acct",
                WiseCamtNamespace + "Ownr",
                WiseCamtNamespace + "Id",
                WiseCamtNamespace + "PrvtId",
                WiseCamtNamespace + "Othr",
                WiseCamtNamespace + "Id");
            var currencyCode = RequireText(
                statement,
                [WiseCamtNamespace + "Acct", WiseCamtNamespace + "Ccy"],
                attachment.FileName);
            var currency = ParseWiseCurrencyType(currencyCode);
            ValidateWiseStatementFileName(attachment.FileName, currencyCode);

            var statementStart = ParseWiseXmlDateTime(RequireText(
                statement,
                [WiseCamtNamespace + "FrToDt", WiseCamtNamespace + "FrDtTm"],
                attachment.FileName)).Date;
            var statementEnd = GetWiseStatementEndDate(ParseWiseXmlDateTime(RequireText(
                statement,
                [WiseCamtNamespace + "FrToDt", WiseCamtNamespace + "ToDtTm"],
                attachment.FileName)));
            var balances = ParseWiseBalances(statement, currency, currencyCode, attachment.FileName);
            var records = statement.Elements(WiseCamtNamespace + "Ntry")
                .Select(entry => ParseWiseRecord(entry, account, currency, currencyCode, attachment.FileName))
                .ToRecords();
            ValidateWiseTransactionSummary(statement, records, attachment.FileName);
            ValidateWiseCurrencyStatementBalance(balances.Beginning, balances.Ending, records, attachment.FileName);

            var internalCardNos = ParseWiseInternalCardNos(
                statement,
                attachment.FileName,
                statementId,
                statementEnd,
                account,
                currency);
            return new WiseCurrencyStatement(
                statementStart,
                statementEnd,
                attachment.ImportTime,
                customerId,
                account,
                currency,
                balances.Beginning,
                balances.Ending,
                records,
                internalCardNos);
        }

        private static string ReadWiseXmlText(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var reader = new StreamReader(stream, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }

        private static void ValidateWiseStatementFileName(string fileName, string currencyCode)
        {
            var match = WiseStatementFileNameRegex.Match(fileName);
            if (!match.Success)
                return;

            if (!String.Equals(match.Groups["currency"].Value, currencyCode, StringComparison.OrdinalIgnoreCase))
                throw new MailParseException($"Wise XML file currency mismatch: {fileName}, xml={currencyCode}");
        }

        private static (Currency Beginning, Currency Ending) ParseWiseBalances(
            XElement statement,
            CurrencyType currency,
            string currencyCode,
            string context)
        {
            var balances = new Dictionary<string, Currency>(StringComparer.OrdinalIgnoreCase);
            foreach (var balance in statement.Elements(WiseCamtNamespace + "Bal"))
            {
                var code = RequireText(
                    balance,
                    [WiseCamtNamespace + "Tp", WiseCamtNamespace + "CdOrPrtry", WiseCamtNamespace + "Cd"],
                    context);
                var amount = ParseWiseSignedAmount(balance, currency, currencyCode, context);
                if (!balances.TryAdd(code, amount))
                    throw new MailParseException($"Duplicate Wise balance type {code}: {context}");
            }

            if (!balances.TryGetValue("OPBD", out var beginning))
                throw new MailParseException($"Missing Wise opening balance: {context}");
            if (!balances.TryGetValue("CLBD", out var ending))
                throw new MailParseException($"Missing Wise closing balance: {context}");

            return (beginning, ending);
        }

        private Record ParseWiseRecord(
            XElement entry,
            Account account,
            CurrencyType statementCurrency,
            string statementCurrencyCode,
            string fileName)
        {
            var status = RequireText(entry, [WiseCamtNamespace + "Sts", WiseCamtNamespace + "Cd"], fileName);
            if (!String.Equals(status, "BOOK", StringComparison.OrdinalIgnoreCase))
                throw new MailParseException($"Unsupported Wise entry status {status}: {fileName}");

            var amount = ParseWiseSignedAmount(entry, statementCurrency, statementCurrencyCode, fileName);
            var time = ParseWiseXmlDateTime(RequireText(
                entry,
                [WiseCamtNamespace + "BookgDt", WiseCamtNamespace + "DtTm"],
                fileName)).DateTime;
            var code = RequireText(
                entry,
                [WiseCamtNamespace + "BkTxCd", WiseCamtNamespace + "Prtry", WiseCamtNamespace + "Cd"],
                fileName);
            var addtlInfo = RequireText(entry, WiseCamtNamespace + "AddtlNtryInf", $"{fileName} {code}");
            var raw = new WiseXmlEntry(
                fileName,
                code,
                OptionalText(entry, WiseCamtNamespace + "NtryRef"),
                time,
                amount,
                addtlInfo,
                OptionalText(
                    entry,
                    WiseCamtNamespace + "NtryDtls",
                    WiseCamtNamespace + "TxDtls",
                    WiseCamtNamespace + "Refs",
                    WiseCamtNamespace + "EndToEndId"),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "NtryDtls",
                    WiseCamtNamespace + "TxDtls",
                    WiseCamtNamespace + "Refs",
                    WiseCamtNamespace + "TxId"),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "NtryDtls",
                    WiseCamtNamespace + "TxDtls",
                    WiseCamtNamespace + "RltdPties",
                    WiseCamtNamespace + "Dbtr",
                    WiseCamtNamespace + "Pty",
                    WiseCamtNamespace + "Nm"),
                OptionalFirstText(
                    entry,
                    [
                        WiseCamtNamespace + "NtryDtls",
                        WiseCamtNamespace + "TxDtls",
                        WiseCamtNamespace + "RltdPties",
                        WiseCamtNamespace + "DbtrAcct",
                        WiseCamtNamespace + "Id",
                        WiseCamtNamespace + "Othr",
                        WiseCamtNamespace + "Id"
                    ],
                    [
                        WiseCamtNamespace + "NtryDtls",
                        WiseCamtNamespace + "TxDtls",
                        WiseCamtNamespace + "RltdPties",
                        WiseCamtNamespace + "DbtrAcct",
                        WiseCamtNamespace + "Id",
                        WiseCamtNamespace + "IBAN"
                    ]),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "NtryDtls",
                    WiseCamtNamespace + "TxDtls",
                    WiseCamtNamespace + "RltdPties",
                    WiseCamtNamespace + "Cdtr",
                    WiseCamtNamespace + "Pty",
                    WiseCamtNamespace + "Nm"),
                OptionalFirstText(
                    entry,
                    [
                        WiseCamtNamespace + "NtryDtls",
                        WiseCamtNamespace + "TxDtls",
                        WiseCamtNamespace + "RltdPties",
                        WiseCamtNamespace + "CdtrAcct",
                        WiseCamtNamespace + "Id",
                        WiseCamtNamespace + "Othr",
                        WiseCamtNamespace + "Id"
                    ],
                    [
                        WiseCamtNamespace + "NtryDtls",
                        WiseCamtNamespace + "TxDtls",
                        WiseCamtNamespace + "RltdPties",
                        WiseCamtNamespace + "CdtrAcct",
                        WiseCamtNamespace + "Id",
                        WiseCamtNamespace + "IBAN"
                    ]),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "NtryDtls",
                    WiseCamtNamespace + "TxDtls",
                    WiseCamtNamespace + "RmtInf",
                    WiseCamtNamespace + "Ustrd"),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "AmtDtls",
                    WiseCamtNamespace + "TxAmt",
                    WiseCamtNamespace + "CcyXchg",
                    WiseCamtNamespace + "SrcCcy"),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "AmtDtls",
                    WiseCamtNamespace + "TxAmt",
                    WiseCamtNamespace + "CcyXchg",
                    WiseCamtNamespace + "TrgtCcy"),
                OptionalText(
                    entry,
                    WiseCamtNamespace + "AmtDtls",
                    WiseCamtNamespace + "TxAmt",
                    WiseCamtNamespace + "CcyXchg",
                    WiseCamtNamespace + "XchgRate"));
            var classification = ClassifyWiseXmlEntry(raw, account);
            return CreateWiseRecord(
                account,
                raw.Time,
                raw.Amount,
                classification.CounterpartyText,
                classification.Reason,
                classification.IsInternal,
                BuildWiseRecordSource(raw));
        }

        private WiseXmlClassification ClassifyWiseXmlEntry(WiseXmlEntry entry, Account account)
        {
            if (entry.Code.StartsWith("FEE-", StringComparison.OrdinalIgnoreCase))
            {
                if (entry.Amount.v >= 0)
                    throw new MailParseException($"Wise fee should be negative: {entry.Code} {entry.Amount.v}/{entry.Amount.t}");

                return new WiseXmlClassification(ExtractWiseFeeTarget(entry.AdditionalInfo), "手续费", false);
            }

            if (entry.Code.StartsWith("BALANCE-", StringComparison.OrdinalIgnoreCase))
                return ClassifyWiseConversion(entry);

            if (entry.Code.StartsWith("CARD-", StringComparison.OrdinalIgnoreCase))
                return ClassifyWiseCardTransaction(entry);

            if (entry.Code.StartsWith("DIRECT_DEBIT-", StringComparison.OrdinalIgnoreCase))
                return ClassifyWiseDirectDebit(entry, account);

            if (entry.Code.StartsWith("TRANSFER-", StringComparison.OrdinalIgnoreCase))
                return ClassifyWiseTransfer(entry, account);

            throw new MailParseException($"Unsupported Wise transaction code: {entry.Code}; text={entry.AdditionalInfo}");
        }

        private static WiseXmlClassification ClassifyWiseConversion(WiseXmlEntry entry)
        {
            var match = WiseConvertedRegex.Match(entry.AdditionalInfo);
            if (!match.Success)
                throw new MailParseException($"Unsupported Wise conversion text: {entry.Code}; text={entry.AdditionalInfo}");
            if (String.IsNullOrWhiteSpace(entry.ExchangeFromCurrency) || String.IsNullOrWhiteSpace(entry.ExchangeToCurrency))
                throw new MailParseException($"Wise conversion missing exchange fields: {entry.Code}");
            if (!String.Equals(entry.ExchangeFromCurrency, match.Groups["fromCurrency"].Value, StringComparison.OrdinalIgnoreCase)
                || !String.Equals(entry.ExchangeToCurrency, match.Groups["toCurrency"].Value, StringComparison.OrdinalIgnoreCase))
            {
                throw new MailParseException($"Wise conversion currency mismatch: {entry.Code}; text={entry.AdditionalInfo}");
            }

            var currentCurrencyCode = ToWiseCurrencyCode(entry.Amount.t);
            if (entry.Amount.v > 0
                && !String.Equals(currentCurrencyCode, entry.ExchangeToCurrency, StringComparison.OrdinalIgnoreCase))
            {
                throw new MailParseException($"Wise conversion credit currency mismatch: {entry.Code}");
            }

            if (entry.Amount.v < 0
                && !String.Equals(currentCurrencyCode, entry.ExchangeFromCurrency, StringComparison.OrdinalIgnoreCase))
            {
                throw new MailParseException($"Wise conversion debit currency mismatch: {entry.Code}");
            }

            return new WiseXmlClassification($"{entry.ExchangeFromCurrency} -> {entry.ExchangeToCurrency}", "兑换", true);
        }

        private static WiseXmlClassification ClassifyWiseCardTransaction(WiseXmlEntry entry)
        {
            var match = WiseCardTransactionRegex.Match(entry.AdditionalInfo);
            if (!match.Success)
                throw new MailParseException($"Unsupported Wise card text: {entry.Code}; text={entry.AdditionalInfo}");
            if (entry.Amount.v >= 0)
                throw new MailParseException($"Wise card transaction should be negative: {entry.Code}");

            return new WiseXmlClassification(NormalizeCounterparty(match.Groups["merchant"].Value), "消费", false);
        }

        private WiseXmlClassification ClassifyWiseDirectDebit(WiseXmlEntry entry, Account account)
        {
            var match = WiseDirectDebitRegex.Match(entry.AdditionalInfo);
            if (!match.Success)
                throw new MailParseException($"Unsupported Wise direct debit text: {entry.Code}; text={entry.AdditionalInfo}");
            if (entry.Amount.v >= 0)
                throw new MailParseException($"Wise direct debit should be negative: {entry.Code}");

            var counterparty = NormalizeCounterparty(match.Groups["counterparty"].Value);
            var classification = ClassifyWiseCounterparty(
                account,
                counterparty,
                false,
                entry);
            return new WiseXmlClassification(classification.CounterpartyText, "直接付款", classification.IsInternal);
        }

        private WiseXmlClassification ClassifyWiseTransfer(WiseXmlEntry entry, Account account)
        {
            if (String.Equals(entry.AdditionalInfo, "Topped up account", StringComparison.OrdinalIgnoreCase))
            {
                if (entry.Amount.v <= 0)
                    throw new MailParseException($"Wise top-up should be positive: {entry.Code}");

                return new WiseXmlClassification("余额充值来源", "余额充值", false);
            }

            var sentMatch = WiseSentMoneyRegex.Match(entry.AdditionalInfo);
            if (sentMatch.Success)
            {
                if (entry.Amount.v >= 0)
                    throw new MailParseException($"Wise outgoing transfer should be negative: {entry.Code}");

                var counterparty = NormalizeCounterparty(String.IsNullOrWhiteSpace(entry.CreditorName)
                    ? sentMatch.Groups["counterparty"].Value
                    : entry.CreditorName);
                var classification = ClassifyWiseCounterparty(
                    account,
                    counterparty,
                    false,
                    entry);
                return new WiseXmlClassification(
                    classification.CounterpartyText,
                    "款项汇出",
                    classification.IsInternal);
            }

            var receivedMatch = WiseReceivedMoneyRegex.Match(entry.AdditionalInfo);
            if (receivedMatch.Success)
            {
                if (entry.Amount.v <= 0)
                    throw new MailParseException($"Wise incoming transfer should be positive: {entry.Code}");

                var counterparty = NormalizeCounterparty(receivedMatch.Groups["counterparty"].Value);
                var classification = ClassifyWiseCounterparty(
                    account,
                    counterparty,
                    true,
                    entry);
                return new WiseXmlClassification(
                    classification.CounterpartyText,
                    "汇款收款",
                    classification.IsInternal);
            }

            throw new MailParseException($"Unsupported Wise transfer text: {entry.Code}; text={entry.AdditionalInfo}");
        }

        private static string ExtractWiseFeeTarget(string text)
        {
            const string prefix = "Wise Charges for:";
            return text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? text[prefix.Length..].Trim()
                : "手续费";
        }

        private static Record CreateWiseRecord(
            Account account,
            DateTime time,
            Currency amount,
            string destAccount,
            string reason,
            bool isInternal,
            string source)
        {
            return new Record
            {
                Account = account,
                date = time,
                updateTime = DateTime.Now,
                DestAccount = destAccount,
                isInternal = isInternal,
                Source = source,
                Reason = reason,
                v = amount.v,
                t = amount.t
            };
        }

        private List<AccountInternalId> ParseWiseInternalCardNos(
            XElement statement,
            string fileName,
            string statementId,
            DateTime statementEndDate,
            Account account,
            CurrencyType currency)
        {
            var result = new List<AccountInternalId>();
            AddWiseInternalCardNo(
                result,
                account,
                OptionalText(
                    statement,
                    WiseCamtNamespace + "Acct",
                    WiseCamtNamespace + "Id",
                    WiseCamtNamespace + "Othr",
                    WiseCamtNamespace + "Id"),
                currency,
                "XML account id",
                fileName,
                statementEndDate);
            AddWiseInternalCardNo(
                result,
                account,
                OptionalText(
                    statement,
                    WiseCamtNamespace + "Acct",
                    WiseCamtNamespace + "Id",
                    WiseCamtNamespace + "IBAN"),
                currency,
                "XML IBAN",
                fileName,
                statementEndDate);

            return result;
        }

        private static void AddWiseInternalCardNo(
            List<AccountInternalId> result,
            Account account,
            string value,
            CurrencyType currency,
            string desc,
            string fileName,
            DateTime statementEndDate)
        {
            if (String.IsNullOrWhiteSpace(value))
                return;

            if (result.Any(item => String.Equals(item.cardNo, value, StringComparison.OrdinalIgnoreCase)))
                return;

            result.Add(new AccountInternalId
            {
                Account = account,
                cardNo = value,
                currencyType = currency,
                desc = desc,
                sourceText = $"Wise XML statement {statementEndDate:yyyy-MM-dd}; {fileName}; {desc}={value}"
            });
        }

        private bool SaveWiseParsedStatement(WiseParsedStatement parsed)
        {
            return database.SaveStatementRecordsOnce(
                WiseProvider,
                parsed.ImportTime,
                parsed.Records,
                parsed.EndingBalances,
                parsed.StatementKey,
                parsed.BeginningBalances,
                afterSaveInTransaction: _ => database.EnsureAccountInternalCardNos(parsed.InternalCardNos));
        }

        private static void PrintWiseParsedStatementSummary(string label, WiseParsedStatement parsed, bool saved)
        {
            Console.WriteLine(saved
                ? $"Import {label} {parsed.StatementKey}, records={parsed.Records.Count}"
                : $"Skip imported {label} {parsed.StatementKey}");
            Console.WriteLine("Wise cards:");
            foreach (var card in parsed.InternalCardNos.OrderBy(card => card.currencyType).ThenBy(card => card.cardNo))
                Console.WriteLine($"{card.cardNo}\t{card.currencyType}\t{card.desc}");
            Console.WriteLine("Wise balances:");
            foreach (var balance in parsed.EndingBalances.OrderBy(balance => balance.t))
                Console.WriteLine($"{balance.t}\t{balance.v}");
            Console.WriteLine("Wise records:");
            foreach (var record in parsed.Records.OrderBy(record => record.date).ThenBy(record => record.t).ThenBy(record => record.v))
            {
                Console.WriteLine(
                    $"{record.date:yyyy-MM-dd HH:mm:ss.ffffff}\t{record.t}\t{record.v}\tinternal={record.isInternal}\treason={record.Reason}\tdest={record.DestAccount}\tsource={record.Source}");
            }
        }

        private static void ValidateWiseTransactionSummary(XElement statement, Records records, string context)
        {
            var summary = statement.Element(WiseCamtNamespace + "TxsSummry");
            if (summary is null)
                return;

            var totalEntries = TryParseWiseInt(OptionalText(
                summary,
                WiseCamtNamespace + "TtlNtries",
                WiseCamtNamespace + "NbOfNtries"));
            var totalSum = TryParseWiseDecimal(OptionalText(
                summary,
                WiseCamtNamespace + "TtlNtries",
                WiseCamtNamespace + "Sum"));
            var creditEntries = TryParseWiseInt(OptionalText(
                summary,
                WiseCamtNamespace + "TtlCdtNtries",
                WiseCamtNamespace + "NbOfNtries"));
            var creditSum = TryParseWiseDecimal(OptionalText(
                summary,
                WiseCamtNamespace + "TtlCdtNtries",
                WiseCamtNamespace + "Sum"));
            var debitEntries = TryParseWiseInt(OptionalText(
                summary,
                WiseCamtNamespace + "TtlDbtNtries",
                WiseCamtNamespace + "NbOfNtries"));
            var debitSum = TryParseWiseDecimal(OptionalText(
                summary,
                WiseCamtNamespace + "TtlDbtNtries",
                WiseCamtNamespace + "Sum"));

            if (totalEntries.HasValue && totalEntries.Value != records.Count)
                throw new MailParseException($"Wise summary entry count mismatch: {context}");
            if (totalSum.HasValue && totalSum.Value != records.Sum(record => record.v))
                throw new MailParseException($"Wise summary total mismatch: {context}");
            if (creditEntries.HasValue && creditEntries.Value != records.Count(record => record.v > 0))
                throw new MailParseException($"Wise summary credit count mismatch: {context}");
            if (creditSum.HasValue && creditSum.Value != records.Where(record => record.v > 0).Sum(record => record.v))
                throw new MailParseException($"Wise summary credit sum mismatch: {context}");
            if (debitEntries.HasValue && debitEntries.Value != records.Count(record => record.v < 0))
                throw new MailParseException($"Wise summary debit count mismatch: {context}");
            if (debitSum.HasValue && Math.Abs(debitSum.Value) != Math.Abs(records.Where(record => record.v < 0).Sum(record => record.v)))
                throw new MailParseException($"Wise summary debit sum mismatch: {context}");
        }

        private static void ValidateWiseCurrencyStatementBalance(Currency beginning, Currency ending, Records records, string context)
        {
            var expected = beginning.v + records.Sum(record => record.v);
            if (expected != ending.v)
            {
                throw new MailParseException(
                    $"Wise balance mismatch: {context}, beginning={beginning.v}, records={records.Sum(record => record.v)}, ending={ending.v}");
            }
        }

        private static Currency ParseWiseSignedAmount(XElement element, CurrencyType expectedCurrency, string expectedCurrencyCode, string context)
        {
            var amountElement = RequireElement(element, WiseCamtNamespace + "Amt", context);
            var amountCurrencyCode = amountElement.Attribute("Ccy")?.Value ?? "";
            if (!String.Equals(amountCurrencyCode, expectedCurrencyCode, StringComparison.OrdinalIgnoreCase))
                throw new MailParseException($"Wise amount currency mismatch: {context}, expected={expectedCurrencyCode}, actual={amountCurrencyCode}");

            var amount = ParseWiseDecimal(amountElement.Value);
            var direction = RequireText(element, WiseCamtNamespace + "CdtDbtInd", context);
            return direction.ToUpperInvariant() switch
            {
                "CRDT" => new Currency(amount, expectedCurrency),
                "DBIT" => new Currency(-amount, expectedCurrency),
                _ => throw new MailParseException($"Unsupported Wise credit/debit indicator {direction}: {context}")
            };
        }

        private static DateTimeOffset ParseWiseXmlDateTime(string value)
        {
            return DateTimeOffset.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.None);
        }

        private static DateTime GetWiseStatementEndDate(DateTimeOffset toDateTime)
        {
            return toDateTime.TimeOfDay == TimeSpan.Zero
                ? toDateTime.Date.AddDays(-1)
                : toDateTime.Date;
        }

        private WiseCounterpartyClassification ClassifyWiseCounterparty(
            Account currentAccount,
            string counterparty,
            bool isIncoming,
            WiseXmlEntry entry)
        {
            var internalCounterparty = database.FindAccountByInternalCardNoText(
                null,
                BuildWiseMatchContext(entry, counterparty),
                counterparty,
                entry.CreditorAccount,
                entry.Reference,
                entry.AdditionalInfo);
            if (internalCounterparty is not null && !IsSameAccount(database.GetPostingAccount(internalCounterparty), currentAccount))
            {
                var counterpartyText = isIncoming
                    ? FormatTransferCounterparty(internalCounterparty.name, WiseAccountName)
                    : FormatTransferCounterparty(WiseAccountName, internalCounterparty.name);
                return new WiseCounterpartyClassification(counterpartyText, true);
            }

            return new WiseCounterpartyClassification(counterparty, false);
        }

        private static string BuildWiseMatchContext(WiseXmlEntry entry, string counterparty)
        {
            return $"Wise XML statement {entry.FileName}; code={entry.Code}; counterparty={counterparty}; creditorAccount={entry.CreditorAccount}; text={entry.AdditionalInfo}";
        }

        private static string BuildWiseRecordSource(WiseXmlEntry entry)
        {
            var parts = new List<string>
            {
                $"Wise XML statement {entry.FileName}",
                $"code={entry.Code}"
            };
            AddWiseSourcePart(parts, "ref", entry.Reference);
            AddWiseSourcePart(parts, "endToEndId", entry.EndToEndId);
            AddWiseSourcePart(parts, "txId", entry.TransactionId);
            AddWiseSourcePart(parts, "payer", entry.DebtorName);
            AddWiseSourcePart(parts, "payerAccount", entry.DebtorAccount);
            AddWiseSourcePart(parts, "payee", entry.CreditorName);
            AddWiseSourcePart(parts, "payeeAccount", entry.CreditorAccount);
            AddWiseSourcePart(parts, "remittance", entry.RemittanceInfo);
            AddWiseSourcePart(parts, "exchangeFrom", entry.ExchangeFromCurrency);
            AddWiseSourcePart(parts, "exchangeTo", entry.ExchangeToCurrency);
            AddWiseSourcePart(parts, "exchangeRate", entry.ExchangeRate);
            AddWiseSourcePart(parts, "text", entry.AdditionalInfo);
            return LimitWiseText(String.Join("; ", parts));
        }

        private static void AddWiseSourcePart(List<string> parts, string name, string value)
        {
            if (!String.IsNullOrWhiteSpace(value))
                parts.Add($"{name}={NormalizeCounterparty(value)}");
        }

        private static string FormatTransferCounterparty(string sourceAccount, string destinationAccount)
        {
            return $"{sourceAccount} -> {destinationAccount}";
        }

        private static CurrencyType ParseWiseCurrencyType(string currencyCode)
        {
            var normalized = currencyCode.Trim().ToUpperInvariant();
            if (normalized == "CNY")
                return CurrencyType.RMB;

            if (Enum.TryParse<CurrencyType>(normalized, out var currencyType))
                return currencyType;

            throw new MailParseException($"Unsupported Wise currency: {currencyCode}");
        }

        private static string ToWiseCurrencyCode(CurrencyType currencyType)
        {
            return currencyType == CurrencyType.RMB ? "CNY" : currencyType.ToString();
        }

        private static string NormalizeCounterparty(string counterparty)
        {
            return Regex.Replace(counterparty, @"\s+", " ")
                .Trim(' ', '。', '，', ',', '.');
        }

        private static decimal ParseWiseDecimal(string value)
        {
            return decimal.Parse(
                value.Replace(",", "", StringComparison.Ordinal),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture);
        }

        private static decimal? TryParseWiseDecimal(string value)
        {
            return String.IsNullOrWhiteSpace(value) ? null : ParseWiseDecimal(value);
        }

        private static int? TryParseWiseInt(string value)
        {
            return String.IsNullOrWhiteSpace(value)
                ? null
                : Int32.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
        }

        private static XElement RequireElement(XElement? parent, XName name, string context)
        {
            if (parent is null)
                throw new MailParseException($"Missing Wise XML parent for {name.LocalName}: {context}");

            var child = parent.Element(name);
            if (child is null)
                throw new MailParseException($"Missing Wise XML element {name.LocalName}: {context}");

            return child;
        }

        private static string RequireText(XElement parent, XName name, string context)
        {
            var value = OptionalText(parent, name);
            if (String.IsNullOrWhiteSpace(value))
                throw new MailParseException($"Missing Wise XML text {name.LocalName}: {context}");

            return value;
        }

        private static string RequireText(XElement parent, XName[] path, string context)
        {
            var value = OptionalText(parent, path);
            if (String.IsNullOrWhiteSpace(value))
                throw new MailParseException($"Missing Wise XML path {String.Join("/", path.Select(item => item.LocalName))}: {context}");

            return value;
        }

        private static string OptionalText(XElement parent, params XName[] path)
        {
            var current = parent;
            foreach (var name in path)
            {
                current = current.Element(name);
                if (current is null)
                    return "";
            }

            return current.Value.Trim();
        }

        private static string OptionalFirstText(XElement parent, params XName[][] paths)
        {
            foreach (var path in paths)
            {
                var value = OptionalText(parent, path);
                if (!String.IsNullOrWhiteSpace(value))
                    return value;
            }

            return "";
        }

        private static string LimitWiseText(string text)
        {
            return text.Length <= 1024 ? text : text[..1024];
        }

        private sealed record InMemoryWiseStatementAttachment(
            string FileName,
            DateTime? ImportTime,
            byte[] Content);

        private sealed record WiseCurrencyStatement(
            DateTime StatementStartDate,
            DateTime StatementEndDate,
            DateTime? ImportTime,
            string CustomerId,
            Account Account,
            CurrencyType Currency,
            Currency BeginningBalance,
            Currency EndingBalance,
            Records Records,
            List<AccountInternalId> InternalCardNos);

        private sealed record WiseParsedStatement(
            DateTime StatementStartDate,
            DateTime StatementEndDate,
            DateTime ImportTime,
            string StatementKey,
            Account Account,
            Records Records,
            List<AccountBalance> BeginningBalances,
            List<AccountBalance> EndingBalances,
            List<AccountInternalId> InternalCardNos);

        private sealed record WiseXmlEntry(
            string FileName,
            string Code,
            string Reference,
            DateTime Time,
            Currency Amount,
            string AdditionalInfo,
            string EndToEndId,
            string TransactionId,
            string DebtorName,
            string DebtorAccount,
            string CreditorName,
            string CreditorAccount,
            string RemittanceInfo,
            string ExchangeFromCurrency,
            string ExchangeToCurrency,
            string ExchangeRate);

        private sealed record WiseXmlClassification(string CounterpartyText, string Reason, bool IsInternal);

        private sealed record WiseCounterpartyClassification(string CounterpartyText, bool IsInternal);
    }

    static class WiseRecordEnumerableExtensions
    {
        public static Records ToRecords(this IEnumerable<Record> records)
        {
            var result = new Records();
            result.AddRange(records);
            return result;
        }
    }
}
