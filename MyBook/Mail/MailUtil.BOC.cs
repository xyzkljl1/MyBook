using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit;
using MailKit.Search;
using UglyToad.PdfPig;

namespace MyBook
{
    // Bank of China monthly credit-card PDF statement discovery and parsing.
    partial class MailUtil
    {
        private const StatementImportProvider BOCProvider = StatementImportProvider.BOCBillMail;
        private const string BOCAccountType = "BOC";
        private const string BOCMailSender = "boczhangdan@bankofchina.com";
        private const string BOCStatementSubject = "中国银行信用卡电子账单";
        private static readonly Regex BOCStatementAttachmentRegex = new(
            @"^中国银行信用卡电子合并账单(?<year>\d{4})年(?<month>\d{2})月账单\.PDF$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex BOCStatementTitleRegex = new(
            @"中国银行信用卡账单\((?<year>\d{4})年(?<month>\d{2})月\)",
            RegexOptions.CultureInvariant);
        private static readonly Regex BOCAccountHeadingRegex = new(
            @"\(卡号[：:](?<tail>\d{4})\)",
            RegexOptions.CultureInvariant);
        private static readonly Regex BOCDateRegex = new(
            @"^\d{4}-\d{2}-\d{2}$",
            RegexOptions.CultureInvariant);
        private static readonly Regex BOCAmountRegex = new(
            @"^\(?-?\d[\d,]*(?:\.\d+)?\)?$",
            RegexOptions.CultureInvariant);

        public async Task FetchBOCBills()
        {
            await RunWithMailSessionScope(FetchBOCBillsBatch).ConfigureAwait(false);
        }

        private async Task FetchBOCBillsBatch()
        {
            var latestStatementMonth = GetLatestBOCStatementMonth();
            var latestImportTime = database.GetLatestStatementImportTime(BOCProvider);
            var searchSince = latestImportTime.HasValue
                ? FirstDayOfMonth(latestImportTime.Value)
                : new DateTime(2000, 1, 1);
            var currentMonth = FirstDayOfMonth(DateTime.Today);
            var messages = await SearchBOCStatementMails(searchSince).ConfigureAwait(false);
            var messagesByMonth = messages
                .Select(message => new
                {
                    Message = message,
                    Month = ReadBOCStatementAttachmentMonth(message)
                })
                .Where(item => item.Month <= currentMonth)
                .GroupBy(item => item.Month)
                .OrderBy(group => group.Key)
                .ToList();

            var expectedMonth = latestStatementMonth?.AddMonths(1);
            foreach (var group in messagesByMonth)
            {
                var statementMonth = group.Key;
                if (latestStatementMonth.HasValue && statementMonth <= latestStatementMonth.Value)
                    continue;
                if (expectedMonth.HasValue && statementMonth > expectedMonth.Value && expectedMonth.Value < currentMonth)
                    throw new InvalidOperationException($"Missing BOC statement for {expectedMonth.Value:yyyy-MM}");

                var message = group
                    .Select(item => item.Message)
                    .OrderByDescending(GetMailDateTime)
                    .ThenByDescending(item => item.UniqueId)
                    .First();
                ImportBOCStatement(statementMonth, message);
                expectedMonth = statementMonth.AddMonths(1);
            }

            if (expectedMonth.HasValue && expectedMonth.Value < currentMonth)
                throw new InvalidOperationException($"Missing BOC statement for {expectedMonth.Value:yyyy-MM}");
        }

        private DateTime? GetLatestBOCStatementMonth()
        {
            var latestKey = database.GetLatestStatementImportKey(BOCProvider);
            if (String.IsNullOrWhiteSpace(latestKey))
                return null;
            if (!DateTime.TryParseExact(latestKey, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var statementDate))
                throw new InvalidOperationException($"Invalid BOC statement key: {latestKey}");

            return FirstDayOfMonth(statementDate);
        }

        private async Task<List<MailAttachmentMessage>> SearchBOCStatementMails(DateTime searchSince)
        {
            var query = SearchQuery.FromContains(BOCMailSender)
                .And(SearchQuery.SubjectContains(BOCStatementSubject))
                .And(SearchQuery.SentSince(searchSince.Date));
            return await SearchAttachmentMessages(
                $"BOC statement since {searchSince:yyyy-MM-dd}",
                query,
                IsBOCStatementSummary,
                IsBOCStatementAttachmentFileName,
                GetMailDateTime).ConfigureAwait(false);
        }

        private static bool IsBOCStatementSummary(IMessageSummary summary)
        {
            var subject = summary.Envelope?.Subject;
            return SummaryIsFrom(summary, BOCMailSender)
                && (String.IsNullOrWhiteSpace(subject)
                    || String.Equals(subject.Trim(), BOCStatementSubject, StringComparison.Ordinal))
                && SummaryHasMatchingAttachment(summary, IsBOCStatementAttachmentFileName);
        }

        private static bool IsBOCStatementAttachmentFileName(string fileName)
        {
            return BOCStatementAttachmentRegex.IsMatch(fileName.Trim());
        }

        private static DateTime ParseBOCStatementAttachmentMonth(string fileName)
        {
            var match = BOCStatementAttachmentRegex.Match(fileName.Trim());
            if (!match.Success)
                throw new MailParseException($"Invalid BOC statement attachment name: {fileName}");

            return new DateTime(
                Int32.Parse(match.Groups["year"].Value, CultureInfo.InvariantCulture),
                Int32.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture),
                1);
        }

        private static DateTime ReadBOCStatementAttachmentMonth(MailAttachmentMessage message)
        {
            var attachments = message.Attachments
                .Where(attachment => IsBOCStatementAttachmentFileName(attachment.FileName))
                .ToList();
            if (attachments.Count != 1)
                throw new MailParseException($"BOC statement mail should contain exactly one matching PDF attachment: {message.Subject}");

            return ParseBOCStatementAttachmentMonth(attachments[0].FileName);
        }

        private void ImportBOCStatement(DateTime statementMonth, MailAttachmentMessage message)
        {
            var attachment = ReadBOCStatementAttachment(message);
            var accounts = GetBOCCreditCardAccounts();
            var parsed = ParseBOCStatement(statementMonth, GetMailDate(message), attachment.Content, accounts);
            if (database.IsStatementKeyImported(BOCProvider, parsed.StatementKey))
            {
                Console.WriteLine($"Skip imported BOC statement {parsed.StatementKey}");
                return;
            }

            var saved = database.SaveStatementRecordsOnce(
                BOCProvider,
                parsed.ImportTime,
                parsed.Records,
                parsed.EndingBalances,
                parsed.StatementKey,
                parsed.BeginningBalances,
                forceValidateBeginningBalances: database.GetLatestStatementImportKey(BOCProvider) is not null);
            Console.WriteLine(saved
                ? $"Import BOC statement {parsed.StatementKey}, records={parsed.Records.Count}"
                : $"Skip imported BOC statement {parsed.StatementKey}");
        }

        private static BOCStatementAttachment ReadBOCStatementAttachment(MailAttachmentMessage message)
        {
            var attachments = message.Attachments
                .Where(attachment => IsBOCStatementAttachmentFileName(attachment.FileName))
                .Select(attachment => new BOCStatementAttachment(attachment.FileName, attachment.Content))
                .ToList();
            if (attachments.Count != 1)
                throw new MailParseException($"BOC statement mail should contain exactly one matching PDF attachment: {message.Subject}");

            return attachments[0];
        }

        private List<Account> GetBOCCreditCardAccounts()
        {
            var accounts = database.GetAllAccounts()
                .Where(account => account.isCredit
                    && account.name.StartsWith($"{BOCAccountType}_", StringComparison.OrdinalIgnoreCase))
                .OrderBy(account => account.name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (accounts.Count == 0)
                throw new InvalidOperationException("BOC statement import requires at least one configured BOC credit-card account");
            foreach (var account in accounts)
                _ = GetBOCAccountTail(account);

            return accounts;
        }

        private static string GetBOCAccountTail(Account account)
        {
            var prefix = $"{BOCAccountType}_";
            if (!account.name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Invalid BOC account name: {account.name}");

            var tail = account.name[prefix.Length..];
            if (!Regex.IsMatch(tail, @"^\d{4}$", RegexOptions.CultureInvariant))
                throw new InvalidOperationException($"BOC credit-card account name must end with four digits: {account.name}");

            return tail;
        }

        private BOCParsedStatement ParseBOCStatement(
            DateTime attachmentMonth,
            DateTime importTime,
            byte[] pdfBytes,
            List<Account> accounts)
        {
            var lines = ReadBOCPdfLines(pdfBytes);
            var titleMonth = ParseBOCStatementTitleMonth(lines);
            if (titleMonth != attachmentMonth)
                throw new MailParseException($"BOC statement attachment/title month mismatch: attachment={attachmentMonth:yyyy-MM}, title={titleMonth:yyyy-MM}");

            var header = ParseBOCStatementHeader(lines);
            if (FirstDayOfMonth(header.StatementClosingDate) != attachmentMonth)
                throw new MailParseException($"BOC statement closing month mismatch: attachment={attachmentMonth:yyyy-MM}, closing={header.StatementClosingDate:yyyy-MM-dd}");

            var accountsByTail = accounts.ToDictionary(GetBOCAccountTail, StringComparer.Ordinal);
            var summaries = ParseBOCAccountSummaries(lines, accountsByTail);
            var transactions = ParseBOCTransactions(lines, accountsByTail);
            ValidateBOCStatement(header, summaries, transactions);

            var statementKey = header.StatementClosingDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var now = DateTime.Now;
            var records = transactions
                .Select(transaction => CreateBOCRecord(transaction, statementKey, now))
                .ToList();
            var beginningBalances = summaries
                .Select(summary => new AccountBalance(summary.Account, new Currency(summary.BeginningBalance, summary.Currency)))
                .ToList();
            var endingBalances = summaries
                .Select(summary => new AccountBalance(summary.Account, new Currency(summary.EndingBalance, summary.Currency)))
                .ToList();

            return new BOCParsedStatement(importTime, statementKey, records, beginningBalances, endingBalances);
        }

        private static List<BOCPdfLine> ReadBOCPdfLines(byte[] pdfBytes)
        {
            try
            {
                using var document = PdfDocument.Open(pdfBytes);
                return document.GetPages()
                    .SelectMany(page => page.GetWords()
                        .GroupBy(word => Math.Round(word.BoundingBox.Bottom / 2.0) * 2.0)
                        .OrderByDescending(group => group.Key)
                        .Select(group => new BOCPdfLine(
                            page.Number,
                            group.Key,
                            group.OrderBy(word => word.BoundingBox.Left)
                                .Select(word => new BOCPdfWord(word.Text, word.BoundingBox.Left))
                                .ToList())))
                    .ToList();
            }
            catch (Exception exception)
            {
                throw new MailParseException($"Parse BOC statement PDF fail: {exception.Message}");
            }
        }

        private static DateTime ParseBOCStatementTitleMonth(List<BOCPdfLine> lines)
        {
            var matches = lines
                .Select(line => BOCStatementTitleRegex.Match(line.Text))
                .Where(match => match.Success)
                .ToList();
            if (matches.Count != 1)
                throw new MailParseException($"Parse BOC statement fail, expected one statement title, actual={matches.Count}");

            return new DateTime(
                Int32.Parse(matches[0].Groups["year"].Value, CultureInfo.InvariantCulture),
                Int32.Parse(matches[0].Groups["month"].Value, CultureInfo.InvariantCulture),
                1);
        }

        private static BOCStatementHeader ParseBOCStatementHeader(List<BOCPdfLine> lines)
        {
            foreach (var line in lines)
            {
                var dateWords = line.Words
                    .Where(word => BOCDateRegex.IsMatch(word.Text))
                    .OrderBy(word => word.Left)
                    .ToList();
                if (dateWords.Count != 2)
                    continue;

                var paymentDueDate = ParseBOCDate(dateWords[0].Text);
                var statementClosingDate = ParseBOCDate(dateWords[1].Text);
                var currentBalances = new Dictionary<CurrencyType, decimal>
                {
                    [CurrencyType.RMB] = ReadBOCSingleAmount(line, 300, 449, "current RMB total balance due")
                };
                var foreignWords = line.Words.Where(word => word.Left >= 449).ToList();
                for (var i = 0; i < foreignWords.Count; i++)
                {
                    if (!TryParseBOCCurrency(foreignWords[i].Text, out var currency))
                        continue;

                    var amountWord = foreignWords.Skip(i + 1).FirstOrDefault(word => IsBOCAmount(word.Text));
                    if (amountWord is null)
                        throw new MailParseException($"Parse BOC statement fail, missing {currency} total balance due");
                    if (!currentBalances.TryAdd(currency, ParseBOCAmount(amountWord.Text)))
                        throw new MailParseException($"Parse BOC statement fail, duplicate {currency} total balance due");
                }

                return new BOCStatementHeader(paymentDueDate, statementClosingDate, currentBalances);
            }

            throw new MailParseException("Parse BOC statement fail, missing account-summary date and balance row");
        }

        private static List<BOCAccountSummary> ParseBOCAccountSummaries(
            List<BOCPdfLine> lines,
            Dictionary<string, Account> accountsByTail)
        {
            var summaries = new List<BOCAccountSummary>();
            Account? currentAccount = null;
            foreach (var line in lines)
            {
                var accountMatch = BOCAccountHeadingRegex.Match(line.Text);
                if (accountMatch.Success)
                {
                    var tail = accountMatch.Groups["tail"].Value;
                    if (!accountsByTail.TryGetValue(tail, out currentAccount))
                        throw new MailParseException($"BOC statement contains an unconfigured credit-card account ending in {tail}");
                    continue;
                }

                var accountTypeWord = line.Words.FirstOrDefault(word =>
                    Regex.IsMatch(word.Text, @"^(?:人民币|外币)/[A-Z]{3}$", RegexOptions.CultureInvariant));
                if (accountTypeWord is null)
                    continue;
                if (currentAccount is null)
                    throw new MailParseException($"Parse BOC statement fail, balance row appears before an account heading: {line.Text}");

                var currencyText = accountTypeWord.Text[(accountTypeWord.Text.IndexOf('/') + 1)..];
                if (!TryParseBOCCurrency(currencyText, out var currency))
                    throw new MailParseException($"Parse BOC statement fail, unsupported balance currency: {currencyText}");

                var beginningAmount = ReadBOCSingleAmount(line, 165, 220, $"{currentAccount.name}/{currency} previous balance");
                var purchases = ReadBOCSingleAmount(line, 220, 310, $"{currentAccount.name}/{currency} purchases");
                var credits = ReadBOCSingleAmount(line, 310, 390, $"{currentAccount.name}/{currency} credits");
                var endingAmount = ReadBOCSingleAmount(line, 435, 495, $"{currentAccount.name}/{currency} new balance");
                var beginningStatus = String.Join(" ", line.Words
                    .Where(word => word.Left >= 110 && word.Left < 165)
                    .Select(word => word.Text));
                var endingStatus = String.Join(" ", line.Words
                    .Where(word => word.Left >= 390 && word.Left < 435)
                    .Select(word => word.Text));
                summaries.Add(new BOCAccountSummary(
                    currentAccount,
                    currency,
                    ApplyBOCBalanceStatus(beginningAmount, beginningStatus, "previous balance"),
                    purchases,
                    credits,
                    ApplyBOCBalanceStatus(endingAmount, endingStatus, "new balance")));
            }

            if (summaries.Count == 0)
                throw new MailParseException("Parse BOC statement fail, no account balance rows found");
            var duplicate = summaries
                .GroupBy(summary => (summary.Account.Id, summary.Currency))
                .FirstOrDefault(group => group.Count() > 1);
            if (duplicate is not null)
                throw new MailParseException($"Parse BOC statement fail, duplicate balance row for accountId={duplicate.Key.Id}/{duplicate.Key.Currency}");

            return summaries;
        }

        private static List<BOCParsedTransaction> ParseBOCTransactions(
            List<BOCPdfLine> lines,
            Dictionary<string, Account> accountsByTail)
        {
            var result = new List<BOCParsedTransaction>();
            CurrencyType? currentCurrency = null;
            foreach (var line in lines)
            {
                if (TryReadBOCTransactionSectionCurrency(line.Text, out var sectionCurrency))
                    currentCurrency = sectionCurrency;

                var transactionDateWord = line.Words.FirstOrDefault(word => word.Left < 120 && BOCDateRegex.IsMatch(word.Text));
                if (transactionDateWord is null)
                    continue;
                var postingDateWord = line.Words.FirstOrDefault(word => word.Left >= 120 && word.Left < 210 && BOCDateRegex.IsMatch(word.Text));
                if (postingDateWord is null)
                    continue;

                var tailWords = line.Words
                    .Where(word => word.Left >= 210 && word.Left < 280 && Regex.IsMatch(word.Text, @"^\d{4}$", RegexOptions.CultureInvariant))
                    .ToList();
                if (!currentCurrency.HasValue && tailWords.Count == 0)
                    continue;
                if (!currentCurrency.HasValue)
                    throw new MailParseException($"Parse BOC statement fail, transaction appears before a currency section: {line.Text}");
                if (tailWords.Count != 1)
                    throw new MailParseException($"Parse BOC statement fail, invalid transaction card tail: {line.Text}");
                var tail = tailWords[0].Text;
                if (!accountsByTail.TryGetValue(tail, out var account))
                    throw new MailParseException($"BOC transaction uses an unconfigured credit-card account ending in {tail}");

                var description = String.Join(" ", line.Words
                    .Where(word => word.Left >= 280 && word.Left < 400)
                    .Select(word => word.Text))
                    .Trim();
                if (String.IsNullOrWhiteSpace(description))
                    throw new MailParseException($"Parse BOC statement fail, missing transaction description: {line.Text}");

                var deposit = ReadBOCOptionalSingleAmount(line, 400, 490, "transaction deposit");
                var expenditure = ReadBOCOptionalSingleAmount(line, 490, Double.MaxValue, "transaction expenditure");
                if (deposit.HasValue == expenditure.HasValue)
                    throw new MailParseException($"Parse BOC statement fail, transaction must contain exactly one deposit or expenditure: {line.Text}");
                if (deposit < 0 || expenditure < 0)
                    throw new MailParseException($"Parse BOC statement fail, transaction columns must be unsigned: {line.Text}");

                result.Add(new BOCParsedTransaction(
                    account,
                    currentCurrency.Value,
                    ParseBOCDate(transactionDateWord.Text),
                    ParseBOCDate(postingDateWord.Text),
                    description,
                    deposit,
                    expenditure));
            }

            return result;
        }

        private static bool TryReadBOCTransactionSectionCurrency(string text, out CurrencyType currency)
        {
            currency = default;
            var match = Regex.Match(text, @"\((?<currency>[A-Z]{3})\).*(?:交易明细|Transaction\s+Detailed\s+List)", RegexOptions.CultureInvariant);
            if (match.Success)
                return TryParseBOCCurrency(match.Groups["currency"].Value, out currency);
            if (text.Contains("人民币交易明细", StringComparison.Ordinal))
            {
                currency = CurrencyType.RMB;
                return true;
            }

            return false;
        }

        private static void ValidateBOCStatement(
            BOCStatementHeader header,
            List<BOCAccountSummary> summaries,
            List<BOCParsedTransaction> transactions)
        {
            var summaryKeys = summaries
                .Select(summary => (summary.Account.Id, summary.Currency))
                .ToHashSet();
            foreach (var transaction in transactions)
            {
                if (!summaryKeys.Contains((transaction.Account.Id, transaction.Currency)))
                {
                    throw new MailParseException(
                        $"Parse BOC statement fail, transaction has no matching balance row: account={transaction.Account.name}, currency={transaction.Currency}");
                }
            }

            foreach (var summary in summaries)
            {
                var matchingTransactions = transactions
                    .Where(transaction => transaction.Account.Id == summary.Account.Id
                        && transaction.Currency == summary.Currency)
                    .ToList();
                var purchases = matchingTransactions.Sum(transaction => transaction.Expenditure ?? 0);
                var credits = matchingTransactions.Sum(transaction => transaction.Deposit ?? 0);
                if (purchases != summary.Purchases || credits != summary.Credits)
                {
                    throw new MailParseException(
                        $"BOC statement detail/summary mismatch for {summary.Account.name}/{summary.Currency}: detail purchases={purchases}, summary purchases={summary.Purchases}, detail credits={credits}, summary credits={summary.Credits}");
                }

                var expectedEnding = summary.BeginningBalance - summary.Purchases + summary.Credits;
                if (expectedEnding != summary.EndingBalance)
                {
                    throw new MailParseException(
                        $"BOC statement balance mismatch for {summary.Account.name}/{summary.Currency}: beginning={summary.BeginningBalance}, purchases={summary.Purchases}, credits={summary.Credits}, ending={summary.EndingBalance}");
                }
            }

            var summaryCurrencies = summaries.Select(summary => summary.Currency).Distinct().ToHashSet();
            foreach (var currency in summaryCurrencies.Union(header.CurrentBalanceDue.Keys))
            {
                var summaryDebt = summaries
                    .Where(summary => summary.Currency == currency)
                    .Sum(summary => Math.Max(-summary.EndingBalance, 0));
                var headerDebt = header.CurrentBalanceDue.TryGetValue(currency, out var balanceDue) ? balanceDue : 0;
                if (summaryDebt != headerDebt)
                {
                    throw new MailParseException(
                        $"BOC statement account/total mismatch for {currency}: account debt={summaryDebt}, total balance due={headerDebt}");
                }
            }
        }

        private Record CreateBOCRecord(BOCParsedTransaction transaction, string statementKey, DateTime updateTime)
        {
            var amount = transaction.Deposit ?? -transaction.Expenditure!.Value;
            var classification = ClassifyBOCTransaction(transaction.Description, amount);
            return new Record
            {
                Account = transaction.Account,
                date = transaction.TransactionDate,
                postingDate = transaction.PostingDate,
                updateTime = updateTime,
                DestAccount = transaction.Description,
                isInternal = classification.IsInternal,
                Source = $"BOC statement {statementKey}; transaction={transaction.TransactionDate:yyyy-MM-dd}; posting={transaction.PostingDate:yyyy-MM-dd}; description={transaction.Description}; deposit={transaction.Deposit}; expenditure={transaction.Expenditure}",
                Reason = classification.Reason,
                v = amount,
                t = transaction.Currency
            };
        }

        private static BOCTransactionClassification ClassifyBOCTransaction(string description, decimal amount)
        {
            if (amount < 0)
            {
                if (ContainsBOCDescription(description, "手续费", "FEE"))
                    return new BOCTransactionClassification("手续费", false);
                if (ContainsBOCDescription(description, "利息", "INTEREST"))
                    return new BOCTransactionClassification("利息", false);
                return new BOCTransactionClassification("消费", false);
            }

            if (ContainsBOCDescription(description, "还款", "PAYMENT"))
                return new BOCTransactionClassification("信用卡还款", true);
            if (ContainsBOCDescription(description, "退款", "退货", "冲正", "REFUND", "REVERSAL"))
                return new BOCTransactionClassification("退款", false);
            return new BOCTransactionClassification("存入", false);
        }

        private static bool ContainsBOCDescription(string description, params string[] values)
        {
            return values.Any(value => description.Contains(value, StringComparison.OrdinalIgnoreCase));
        }

        private static decimal ApplyBOCBalanceStatus(decimal amount, string status, string fieldName)
        {
            if (amount == 0)
                return 0;
            if (ContainsBOCDescription(status, "欠款", "DEBT"))
                return -Math.Abs(amount);
            if (ContainsBOCDescription(status, "存款", "CREDIT"))
                return Math.Abs(amount);

            throw new MailParseException($"Parse BOC statement fail, unrecognized {fieldName} status: {status}");
        }

        private static DateTime ParseBOCDate(string value)
        {
            return DateTime.ParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static bool TryParseBOCCurrency(string value, out CurrencyType currency)
        {
            return Enum.TryParse(value.Trim(), false, out currency)
                && Enum.IsDefined(currency);
        }

        private static bool IsBOCAmount(string value)
        {
            return BOCAmountRegex.IsMatch(value.Trim());
        }

        private static decimal ReadBOCSingleAmount(BOCPdfLine line, double left, double right, string fieldName)
        {
            return ReadBOCOptionalSingleAmount(line, left, right, fieldName)
                ?? throw new MailParseException($"Parse BOC statement fail, missing {fieldName}: {line.Text}");
        }

        private static decimal? ReadBOCOptionalSingleAmount(BOCPdfLine line, double left, double right, string fieldName)
        {
            var amountWords = line.Words
                .Where(word => word.Left >= left && word.Left < right && IsBOCAmount(word.Text))
                .ToList();
            if (amountWords.Count > 1)
                throw new MailParseException($"Parse BOC statement fail, multiple values for {fieldName}: {line.Text}");

            return amountWords.Count == 0 ? null : ParseBOCAmount(amountWords[0].Text);
        }

        private static decimal ParseBOCAmount(string value)
        {
            var normalized = value.Trim();
            var negative = normalized.StartsWith('(') && normalized.EndsWith(')');
            normalized = normalized.Trim('(', ')').Replace(",", "", StringComparison.Ordinal);
            var amount = Decimal.Parse(normalized, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            return negative ? -Math.Abs(amount) : amount;
        }

        private sealed record BOCStatementAttachment(string FileName, byte[] Content);
        private sealed record BOCPdfWord(string Text, double Left);
        private sealed record BOCPdfLine(int PageNumber, double Bottom, List<BOCPdfWord> Words)
        {
            public string Text => String.Join(" ", Words.Select(word => word.Text));
        }
        private sealed record BOCStatementHeader(
            DateTime PaymentDueDate,
            DateTime StatementClosingDate,
            Dictionary<CurrencyType, decimal> CurrentBalanceDue);
        private sealed record BOCAccountSummary(
            Account Account,
            CurrencyType Currency,
            decimal BeginningBalance,
            decimal Purchases,
            decimal Credits,
            decimal EndingBalance);
        private sealed record BOCParsedTransaction(
            Account Account,
            CurrencyType Currency,
            DateTime TransactionDate,
            DateTime PostingDate,
            string Description,
            decimal? Deposit,
            decimal? Expenditure);
        private sealed record BOCTransactionClassification(string Reason, bool IsInternal);
        private sealed record BOCParsedStatement(
            DateTime ImportTime,
            string StatementKey,
            List<Record> Records,
            List<AccountBalance> BeginningBalances,
            List<AccountBalance> EndingBalances);
    }
}
