using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit.Search;
using MimeKit;

namespace MyBook
{
    // OCBC mail discovery and parsing. OCBC mails currently provide balance deltas,
    // so imports save Records without updating AccountBalances.
    partial class MailUtil
    {
        private const StatementImportProvider OCBCProvider = StatementImportProvider.OCBCMail;
        private const string OCBCMailSender = "notifications@ocbc.com";
        private const string OCBCAccountName = "OCBC";
        private const string OCBCWiseAccountName = "WISE";
        private static readonly DateTime OCBCSearchStartDate = new(2024, 1, 1);
        private static readonly string[] OCBCTransactionSubjectKeywords =
        [
            "funds transfer",
            "sent money",
            "FX order"
        ];

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
            var account = database.GetAccountByName(OCBCAccountName);
            var messages = await SearchOCBCMails(since, before);
            Console.WriteLine($"Fetch OCBC mails: {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}, count={messages.Count}");

            var savedCount = 0;
            var skippedCount = 0;
            var recordCount = 0;
            foreach (var parsed in ParseOCBCMails(messages, account))
            {
                var saved = database.SaveStatementRecordsOnce(
                    OCBCProvider,
                    parsed.Time,
                    parsed.Records,
                    statementKey: parsed.StatementKey);
                if (saved)
                {
                    savedCount++;
                    recordCount += parsed.Records.Count;
                    Console.WriteLine($"Import OCBC mail {parsed.Time:yyyy-MM-dd HH:mm:ss} {parsed.StatementKey}, records={parsed.Records.Count}");
                }
                else
                {
                    skippedCount++;
                    Console.WriteLine($"Skip imported OCBC mail {parsed.Time:yyyy-MM-dd HH:mm:ss} {parsed.StatementKey}");
                }
            }

            Console.WriteLine($"Fetch OCBC mails done: saved={savedCount}, skipped={skippedCount}, records={recordCount}");
        }

        private async Task<List<MimeMessage>> SearchOCBCMails(DateTime since, DateTime before)
        {
            var query = SearchQuery.FromContains(OCBCMailSender)
                .And(SearchQuery.SentSince(since.Date))
                .And(SearchQuery.SentBefore(before.Date));
            return await SearchMessages(
                $"OCBC {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}",
                query,
                IsOCBCTransactionMail,
                GetMailDateTime);
        }

        private static bool IsOCBCTransactionMail(MimeMessage message)
        {
            return IsFrom(message, OCBCMailSender)
                && OCBCTransactionSubjectKeywords.Any(keyword =>
                    message.Subject?.Contains(keyword, StringComparison.OrdinalIgnoreCase) == true);
        }

        private static List<OCBCParsedMail> ParseOCBCMails(List<MimeMessage> messages, Account account)
        {
            return messages
                .Select(message => ParseOCBCMail(message, account))
                .OrderBy(parsed => parsed.Time)
                .ThenBy(parsed => parsed.StatementKey, StringComparer.Ordinal)
                .ToList();
        }

        private static OCBCParsedMail ParseOCBCMail(MimeMessage message, Account account)
        {
            if (!IsOCBCTransactionMail(message))
                throw new MailParseException($"Unsupported OCBC mail sender or subject: {message.Subject}");

            var subject = NormalizeMailText(message.Subject ?? "");
            var text = NormalizeMailText(GetMessageText(message));
            var source = $"OCBC mail: {subject}";
            var statementKey = BuildOCBCStatementKey(message, subject, text);

            if (subject.Contains("sent money", StringComparison.OrdinalIgnoreCase))
                return ParseOCBCPayNowTransfer(account, subject, text, source, statementKey);

            if (subject.Contains("FX order", StringComparison.OrdinalIgnoreCase))
                return ParseOCBCFxOrder(account, text, source, statementKey);

            if (subject.Contains("funds transfer", StringComparison.OrdinalIgnoreCase))
                return ParseOCBCFundsTransfer(account, text, source, statementKey);

            throw new MailParseException($"Unsupported OCBC mail format: subject={subject}; text={TrimForError(text)}");
        }

        private static OCBCParsedMail ParseOCBCPayNowTransfer(
            Account account,
            string subject,
            string text,
            string source,
            string statementKey)
        {
            var date = ParseOCBCTransferDateTime(text);
            var amount = ParseOCBCAmountAt(text, "Amount");
            var fromAccount = ExtractOCBCField(text, "From your account", ["Description", "Description/reference no.", "Reference number", "You can also", "For assistance"]);
            var description = ExtractOptionalOCBCField(text, "Description", ["OCBC Reference number", "You can also", "For assistance"])
                ?? ExtractOptionalOCBCField(text, "Description/reference no.", ["You can also", "For assistance"])
                ?? "";

            if (!subject.Contains("WISE ASIA-PACIFIC", StringComparison.OrdinalIgnoreCase))
                throw new MailParseException($"Unsupported OCBC PayNow recipient: subject={subject}; text={TrimForError(text)}");

            var destAccount = FormatTransferCounterparty(OCBCAccountName, OCBCWiseAccountName);
            var records = new Records
            {
                CreateOCBCRecord(
                    account,
                    date,
                    NegateOCBCAmount(amount),
                    String.IsNullOrWhiteSpace(description) ? destAccount : $"{destAccount} / {description}",
                    "转账",
                    true,
                    source)
            };
            Console.WriteLine($"Parse OCBC PayNow transfer from {fromAccount}: {amount.v:0.##} {amount.t}");
            return new OCBCParsedMail(date, statementKey, records);
        }

        private static OCBCParsedMail ParseOCBCFxOrder(
            Account account,
            string text,
            string source,
            string statementKey)
        {
            var date = ParseOCBCFxDateTime(text);
            var sourceAmount = ParseOCBCAmountAt(text, "Amount");
            var equivalentAmount = ParseOCBCAmountAt(text, "Equivalent amount");
            var fromAccount = ExtractOCBCField(text, "From your account", ["To account"]);
            var toAccount = ExtractOCBCField(text, "To account", ["Reference number", "For assistance"]);
            var destAccount = $"{fromAccount} -> {toAccount}";
            var records = new Records
            {
                CreateOCBCRecord(account, date, NegateOCBCAmount(sourceAmount), destAccount, "换汇", true, source),
                CreateOCBCRecord(account, date, equivalentAmount, destAccount, "换汇", true, source)
            };
            return new OCBCParsedMail(date, statementKey, records);
        }

        private static OCBCParsedMail ParseOCBCFundsTransfer(
            Account account,
            string text,
            string source,
            string statementKey)
        {
            var date = ParseOCBCTransferDateTime(text);
            var amount = ParseOCBCAmountAt(text, "Amount");
            var fromAccount = ExtractOCBCField(text, "From your account", ["To account"]);
            var toAccount = ExtractOCBCField(text, "To account", ["Reference number", "For assistance"]);
            if (IsOCBCOwnAccountText(fromAccount) && IsOCBCOwnAccountText(toAccount))
            {
                Console.WriteLine($"Skip OCBC own-account transfer: {amount.v:0.##} {amount.t}, {fromAccount} -> {toAccount}");
                return new OCBCParsedMail(date, statementKey, []);
            }

            throw new MailParseException($"Unsupported OCBC funds transfer: {TrimForError(text)}");
        }

        private static Record CreateOCBCRecord(
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
                DestAccount = LimitOCBCRecordText(destAccount),
                isInternal = isInternal,
                Reason = reason,
                Source = LimitOCBCRecordText(source),
                v = amount.v,
                t = amount.t
            };
        }

        private static DateTime ParseOCBCTransferDateTime(string text)
        {
            var match = Regex.Match(
                text,
                @"Date of Transfer\s*:\s*(?<date>\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s*Time of Transfer\s*:\s*(?<time>\d{1,2}[:.]\d{2}(?:\s*(?:AM|PM))?(?:\s*SGT)?)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                throw new MailParseException($"Parse OCBC Mail Fail, Missing Transfer Date: {TrimForError(text)}");

            var date = DateTime.ParseExact(match.Groups["date"].Value.Trim(), "d MMM yyyy", CultureInfo.InvariantCulture);
            return date.Date + ParseOCBCTime(match.Groups["time"].Value);
        }

        private static DateTime ParseOCBCFxDateTime(string text)
        {
            var match = Regex.Match(
                text,
                @"Date of exchange:\s*(?<date>\d{1,2}/\d{1,2}/\d{4})\s*Time of exchange:\s*(?<time>\d{1,2}[:.]\d{2}\s*(?:am|pm)?)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                throw new MailParseException($"Parse OCBC Mail Fail, Missing FX Date: {TrimForError(text)}");

            var date = DateTime.ParseExact(match.Groups["date"].Value.Trim(), "d/M/yyyy", CultureInfo.InvariantCulture);
            return date.Date + ParseOCBCTime(match.Groups["time"].Value);
        }

        private static TimeSpan ParseOCBCTime(string value)
        {
            var normalized = Regex.Replace(value, @"\bSGT\b", "", RegexOptions.IgnoreCase)
                .Replace('.', ':')
                .Trim()
                .ToUpperInvariant();
            var amPmMatch = Regex.Match(normalized, @"^(?<hour>\d{1,2}):(?<minute>\d{2})\s*(?<period>AM|PM)?$");
            if (!amPmMatch.Success)
                throw new MailParseException($"Parse OCBC Mail Fail, Invalid Time: {value}");

            var hour = Int32.Parse(amPmMatch.Groups["hour"].Value, CultureInfo.InvariantCulture);
            var minute = Int32.Parse(amPmMatch.Groups["minute"].Value, CultureInfo.InvariantCulture);
            var period = amPmMatch.Groups["period"].Value;
            if (!String.IsNullOrWhiteSpace(period) && hour <= 12)
            {
                if (period == "PM" && hour < 12)
                    hour += 12;
                if (period == "AM" && hour == 12)
                    hour = 0;
            }

            return new TimeSpan(hour, minute, 0);
        }

        private static Currency ParseOCBCAmountAt(string text, string label)
        {
            var match = Regex.Match(
                text,
                $@"{Regex.Escape(label)}\s*:\s*(?:(?<currency1>[A-Z]{{3}})\s*(?<amount1>[\d,]+(?:\.\d+)?)|(?<amount2>[\d,]+(?:\.\d+)?)\s*(?<currency2>[A-Z]{{3}}))",
                RegexOptions.Singleline);
            if (!match.Success)
                throw new MailParseException($"Parse OCBC Mail Fail, Missing Amount: {label}; {TrimForError(text)}");

            var amountText = match.Groups["amount1"].Success ? match.Groups["amount1"].Value : match.Groups["amount2"].Value;
            var currencyText = match.Groups["currency1"].Success ? match.Groups["currency1"].Value : match.Groups["currency2"].Value;
            var amount = Decimal.Parse(amountText.Replace(",", ""), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
            return new Currency(amount, ParseOCBCCurrencyType(currencyText));
        }

        private static CurrencyType ParseOCBCCurrencyType(string currencyText)
        {
            if (String.Equals(currencyText, "CNY", StringComparison.OrdinalIgnoreCase))
                return CurrencyType.RMB;
            if (Enum.TryParse<CurrencyType>(currencyText.Trim(), true, out var currencyType))
                return currencyType;
            throw new MailParseException($"Unsupported OCBC currency: {currencyText}");
        }

        private static string ExtractOCBCField(string text, string label, string[] nextLabels)
        {
            var value = ExtractOptionalOCBCField(text, label, nextLabels);
            if (String.IsNullOrWhiteSpace(value))
                throw new MailParseException($"Parse OCBC Mail Fail, Missing Field: {label}; {TrimForError(text)}");
            return value;
        }

        private static string? ExtractOptionalOCBCField(string text, string label, string[] nextLabels)
        {
            var nextPattern = String.Join("|", nextLabels.Select(Regex.Escape));
            var match = Regex.Match(
                text,
                $@"{Regex.Escape(label)}\s*:\s*(?<value>.*?)(?=\s*(?:{nextPattern})\s*:|\s*(?:{nextPattern})\b|$)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? NormalizeMailText(match.Groups["value"].Value).Trim() : null;
        }

        private static string BuildOCBCStatementKey(MimeMessage message, string subject, string text)
        {
            var reference = Regex.Match(text, @"Reference number\s*:\s*(?<id>\d+)", RegexOptions.IgnoreCase);
            if (reference.Success)
                return $"OCBC_{reference.Groups["id"].Value}";

            if (!String.IsNullOrWhiteSpace(message.MessageId))
                return $"OCBCMessage_{message.MessageId.Trim()}";

            var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(subject + "\n" + text)));
            return $"OCBCHash_{hash[..16]}";
        }

        private static bool IsOCBCOwnAccountText(string value)
        {
            return value.Contains("Account", StringComparison.OrdinalIgnoreCase)
                && Regex.IsMatch(value, @"-\d{3,}\)", RegexOptions.CultureInvariant);
        }

        private static Currency NegateOCBCAmount(Currency amount)
        {
            if (amount.v < 0)
                throw new MailParseException($"OCBC amount should be positive before applying direction: {amount.v}/{amount.t}");
            return new Currency(-amount.v, amount.t);
        }

        private static string LimitOCBCRecordText(string text)
        {
            return text.Length <= 1024 ? text : text[..1024];
        }

        private sealed record OCBCParsedMail(
            DateTime Time,
            string StatementKey,
            Records Records);
    }
}
