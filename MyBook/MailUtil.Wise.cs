using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;

namespace MyBook
{
    // Wise mail discovery and parsing. Wise mails only provide balance deltas,
    // so imports save Records without updating AccountBalances.
    partial class MailUtil
    {
        private const StatementImportProvider WiseProvider = StatementImportProvider.WiseMail;
        private const string WiseMailSender = "noreply@wise.com";
        private const string WiseAccountName = "WISE";
        private const int WiseMailClientTimeoutMilliseconds = 30000;

        private static readonly DateTime WiseSearchStartDate = new(2025, 1, 1);
        private static readonly TimeSpan WiseMailClientTimeout = TimeSpan.FromMilliseconds(WiseMailClientTimeoutMilliseconds);
        private static readonly string[] WiseTransactionSubjectKeywords =
        [
            "付款",
            "汇款",
            "款项",
            "消费",
            "充值"
        ];

        public async Task FetchWiseReports()
        {
            await FetchWiseReports(WiseSearchStartDate, DateTime.Today.AddDays(1));
        }

        public async Task FetchWiseReports(DateTime date)
        {
            var reportMonth = FirstDayOfMonth(date);
            await FetchWiseReports(reportMonth, reportMonth.AddMonths(1));
        }

        private async Task FetchWiseReports(DateTime since, DateTime before)
        {
            var account = database.GetAccountByName(WiseAccountName);
            var settings = ReadWiseSettings();
            var messages = await SearchWiseMails(since, before);
            Console.WriteLine($"Fetch Wise mails: {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}, count={messages.Count}");

            var savedCount = 0;
            var skippedCount = 0;
            var recordCount = 0;
            foreach (var parsed in ParseWiseMails(messages, account, settings))
            {
                var saved = database.SaveStatementRecordsWithoutBalanceOnce(
                    WiseProvider,
                    parsed.Time,
                    parsed.StatementKey,
                    parsed.Records);
                if (saved)
                {
                    savedCount++;
                    recordCount += parsed.Records.Count;
                    Console.WriteLine($"Import Wise mail {parsed.Time:yyyy-MM-dd} {parsed.StatementKey}, records={parsed.Records.Count}");
                }
                else
                {
                    skippedCount++;
                    Console.WriteLine($"Skip imported Wise mail {parsed.Time:yyyy-MM-dd} {parsed.StatementKey}");
                }
            }

            Console.WriteLine($"Fetch Wise mails done: saved={savedCount}, skipped={skippedCount}, records={recordCount}");
        }

        private async Task<List<MimeMessage>> SearchWiseMails(DateTime since, DateTime before)
        {
            try
            {
                using ImapClient client = new();
                client.Timeout = WiseMailClientTimeoutMilliseconds;
                client.CheckCertificateRevocation = false;
                client.ProxyClient = null;

                Console.WriteLine($"mail connect direct Wise {since:yyyy-MM-dd}..{before.AddDays(-1):yyyy-MM-dd}");
                await RunWiseMailOperation(token => client.ConnectAsync("imap.mail.yahoo.com", 993, true, token));
                Console.WriteLine("Wise mail connected");
                await RunWiseMailOperation(token => client.AuthenticateAsync(username, apppasswd, token));
                Console.WriteLine("Wise mail authenticated");
                await RunWiseMailOperation(token => client.Inbox.OpenAsync(FolderAccess.ReadOnly, token));
                Console.WriteLine("Wise mail inbox opened");

                var query = SearchQuery.FromContains(WiseMailSender)
                    .And(SearchQuery.SentSince(since.Date))
                    .And(SearchQuery.SentBefore(before.Date));
                var uids = await RunWiseMailOperation(token => client.Inbox.SearchAsync(query, token));
                Console.WriteLine($"Wise mail search found {uids.Count}");
                var messages = new List<MimeMessage>();
                foreach (var uid in uids)
                {
                    var message = await RunWiseMailOperation(token => client.Inbox.GetMessageAsync(uid, token));
                    if (IsWiseTransactionMail(message))
                        messages.Add(message);
                }

                return messages
                    .OrderBy(GetWiseMailDateTime)
                    .ThenBy(message => message.Subject, StringComparer.Ordinal)
                    .ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine($"fetch Wise mail fail :{e.Message}");
                throw new InvalidOperationException($"Fetch Wise mail failed: {e.Message}", e);
            }
        }

        private static bool IsWiseTransactionMail(MimeMessage message)
        {
            return IsWiseMail(message)
                && WiseTransactionSubjectKeywords.Any(keyword =>
                    message.Subject?.Contains(keyword, StringComparison.OrdinalIgnoreCase) == true);
        }

        private static bool IsWiseMail(MimeMessage message)
        {
            return message.From.Mailboxes.Any(mailbox =>
                String.Equals(mailbox.Address, WiseMailSender, StringComparison.OrdinalIgnoreCase));
        }

        private static List<WiseParsedMail> ParseWiseMails(List<MimeMessage> messages, Account account, WiseSettings settings)
        {
            return messages
                .Select(message => ParseWiseMail(message, account, settings))
                .Where(parsed => parsed is not null)
                .Cast<WiseParsedMail>()
                .GroupBy(parsed => parsed.StatementKey, StringComparer.Ordinal)
                .Select(group => group
                    .OrderByDescending(parsed => parsed.Priority)
                    .ThenByDescending(parsed => parsed.Time)
                    .First())
                .OrderBy(parsed => parsed.Time)
                .ThenBy(parsed => parsed.StatementKey, StringComparer.Ordinal)
                .ToList();
        }

        private static WiseParsedMail? ParseWiseMail(MimeMessage message, Account account, WiseSettings settings)
        {
            if (!IsWiseTransactionMail(message))
                return null;

            var subject = NormalizeWiseText(message.Subject ?? "");
            var text = NormalizeWiseText(GetMessageText(message));
            var time = GetWiseMailDateTime(message);
            var source = $"Wise mail: {subject}";
            var statementKey = BuildWiseStatementKey(message, subject, text);

            if (TryParseWiseBalanceTopUp(text, out var topUpAmount))
            {
                return CreateParsedMail(
                    time,
                    statementKey,
                    WiseImportPriority.Normal,
                    [CreateWiseRecord(account, time, topUpAmount, "Wise余额充值来源", "余额充值", false, source)]);
            }

            if (TryParseWiseDirectPayment(subject, text, out var directCounterparty, out var directAmount))
            {
                var classification = ClassifyWiseCounterparty(settings, directCounterparty, directAmount, false);
                return CreateParsedMail(
                    time,
                    statementKey,
                    WiseImportPriority.Normal,
                    [CreateWiseRecord(
                        account,
                        time,
                        Negate(directAmount),
                        classification.CounterpartyText,
                        "直接付款",
                        classification.IsInternal,
                        source)]);
            }

            if (TryParseWiseIncoming(text, out var incomingCounterparty, out var incomingAmount))
            {
                var classification = ClassifyWiseCounterparty(settings, incomingCounterparty, incomingAmount, true);
                return CreateParsedMail(
                    time,
                    statementKey,
                    WiseImportPriority.Normal,
                    [CreateWiseRecord(
                        account,
                        time,
                        incomingAmount,
                        classification.CounterpartyText,
                        GetWiseIncomingReason(subject, incomingCounterparty),
                        classification.IsInternal,
                        source)]);
            }

            if (TryParseWiseReceivedFunds(text, out var sourceAmount, out var receivedAmount, out var fee))
            {
                if (fee.t != sourceAmount.t)
                    throw new MailParseException($"Wise fee currency mismatch: {fee.v}/{fee.t}, source={sourceAmount.v}/{sourceAmount.t}");
                if (fee.v < 0 || fee.v > sourceAmount.v)
                    throw new MailParseException($"Invalid Wise fee: {fee.v}/{fee.t}, source={sourceAmount.v}/{sourceAmount.t}");

                var transferAmount = new Currency(sourceAmount.v - fee.v, sourceAmount.t);
                var counterparty = String.IsNullOrWhiteSpace(settings.ReceivedFundsDestinationAccount)
                    ? "Wise汇款收款方"
                    : FormatTransferCounterparty(WiseAccountName, settings.ReceivedFundsDestinationAccount);
                var records = new List<Record>();
                if (transferAmount.v > 0)
                {
                    records.Add(CreateWiseRecord(
                        account,
                        time,
                        Negate(transferAmount),
                        counterparty,
                        "款项汇出",
                        true,
                        source));
                }

                if (fee.v > 0)
                {
                    records.Add(CreateWiseRecord(
                        account,
                        time,
                        Negate(fee),
                        counterparty,
                        "手续费",
                        false,
                        source));
                }

                return CreateParsedMail(
                    time,
                    statementKey,
                    WiseImportPriority.Detailed,
                    records);
            }

            if (TryParseWiseCardPayment(subject, text, out var merchant, out var cardAmount))
            {
                return CreateParsedMail(
                    time,
                    statementKey,
                    WiseImportPriority.Normal,
                    [CreateWiseRecord(account, time, Negate(cardAmount), merchant, "消费", false, source)]);
            }

            if (TryParseWiseOutgoing(text, out var outgoingCounterparty, out var outgoingAmount, out var outgoingFee, out var outgoingInternal, out var isPendingConversion))
            {
                if (isPendingConversion)
                    return null;

                var classification = ClassifyWiseCounterparty(settings, outgoingCounterparty, outgoingAmount, false);
                var records = new List<Record>
                {
                    CreateWiseRecord(
                        account,
                        time,
                        Negate(outgoingAmount),
                        classification.CounterpartyText,
                        GetWiseOutgoingReason(subject, outgoingCounterparty),
                        outgoingInternal || classification.IsInternal,
                        source)
                };

                if (outgoingFee.v > 0)
                {
                    records.Add(CreateWiseRecord(
                        account,
                        time,
                        Negate(outgoingFee),
                        classification.CounterpartyText,
                        "手续费",
                        false,
                        source));
                }

                return CreateParsedMail(
                    time,
                    statementKey,
                    WiseImportPriority.Pending,
                    records);
            }

            throw new MailParseException($"Unsupported Wise mail format: subject={subject}; text={TrimForError(text)}");
        }

        private static WiseParsedMail CreateParsedMail(
            DateTime time,
            string statementKey,
            WiseImportPriority priority,
            List<Record> records)
        {
            if (records.Count == 0)
                throw new MailParseException($"Wise mail contains no records: {statementKey}");

            return new WiseParsedMail(time, statementKey, records, priority);
        }

        private static bool TryParseWiseBalanceTopUp(string text, out Currency amount)
        {
            var match = Regex.Match(
                text,
                @"您已充值了\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})\s*至您的余额",
                RegexOptions.Singleline);
            if (match.Success)
            {
                amount = ParseWiseCurrency(match.Groups["amount"].Value, match.Groups["currency"].Value);
                return true;
            }

            amount = new Currency();
            return false;
        }

        private static bool TryParseWiseDirectPayment(
            string subject,
            string text,
            out string counterparty,
            out Currency amount)
        {
            var counterpartyMatch = Regex.Match(subject, @"直接付款给\s*(?<counterparty>.+)$", RegexOptions.Singleline);
            var amountMatch = Regex.Match(
                text,
                @"这笔交易使用了您账户中的\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            if (counterpartyMatch.Success && amountMatch.Success)
            {
                counterparty = NormalizeCounterparty(counterpartyMatch.Groups["counterparty"].Value);
                amount = ParseWiseCurrency(amountMatch.Groups["amount"].Value, amountMatch.Groups["currency"].Value);
                return true;
            }

            counterparty = "";
            amount = new Currency();
            return false;
        }

        private static bool TryParseWiseIncoming(string text, out string counterparty, out Currency amount)
        {
            var match = Regex.Match(
                text,
                @"您已收到来自\s*(?<counterparty>.+?)\s*的\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            if (match.Success)
                return TryBuildWiseParseResult(match, out counterparty, out amount);

            var amountMatch = Regex.Match(
                text,
                @"已收到的金额：\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            var counterpartyMatch = Regex.Match(
                text,
                @"来自：\s*(?<counterparty>.+?)\s+已收到的金额：",
                RegexOptions.Singleline);
            if (amountMatch.Success && counterpartyMatch.Success)
            {
                counterparty = NormalizeCounterparty(counterpartyMatch.Groups["counterparty"].Value);
                amount = ParseWiseCurrency(amountMatch.Groups["amount"].Value, amountMatch.Groups["currency"].Value);
                return true;
            }

            counterparty = "";
            amount = new Currency();
            return false;
        }

        private static bool TryParseWiseReceivedFunds(
            string text,
            out Currency sourceAmount,
            out Currency receivedAmount,
            out Currency fee)
        {
            var receivedMatch = Regex.Match(
                text,
                @"(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})\s*现已汇入您的账户",
                RegexOptions.Singleline);
            var sourceMatch = Regex.Match(
                text,
                @"金额：\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            var feeMatch = Regex.Match(
                text,
                @"Wise\s*的手续费：\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            if (receivedMatch.Success && sourceMatch.Success && feeMatch.Success)
            {
                sourceAmount = ParseWiseCurrency(sourceMatch.Groups["amount"].Value, sourceMatch.Groups["currency"].Value);
                receivedAmount = ParseWiseCurrency(receivedMatch.Groups["amount"].Value, receivedMatch.Groups["currency"].Value);
                fee = ParseWiseCurrency(feeMatch.Groups["amount"].Value, feeMatch.Groups["currency"].Value);
                return true;
            }

            sourceAmount = new Currency();
            receivedAmount = new Currency();
            fee = new Currency();
            return false;
        }

        private static bool TryParseWiseCardPayment(
            string subject,
            string text,
            out string merchant,
            out Currency amount)
        {
            var match = Regex.Match(
                text,
                @"您在\s*(?<counterparty>.+?)\s*消费了\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            if (match.Success)
                return TryBuildWiseParseResult(match, out merchant, out amount);

            match = Regex.Match(
                subject,
                @"已在\s*(?<counterparty>.+?)\s*消费\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            if (match.Success)
                return TryBuildWiseParseResult(match, out merchant, out amount);

            merchant = "";
            amount = new Currency();
            return false;
        }

        private static bool TryParseWiseOutgoing(
            string text,
            out string counterparty,
            out Currency amount,
            out Currency fee,
            out bool isInternal,
            out bool isPendingConversion)
        {
            var match = Regex.Match(
                text,
                @"(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})\s*已存入\s*(?<counterparty>.+?)\s*的账户",
                RegexOptions.Singleline);
            if (match.Success)
            {
                var parsed = TryBuildWiseParseResult(match, out counterparty, out amount);
                fee = ParseWiseFeeOrZero(text, amount.t);
                isInternal = IsBrokerTransfer(counterparty);
                isPendingConversion = false;
                return parsed;
            }

            match = Regex.Match(
                text,
                @"您发送给\s*(?<counterparty>.+?)\s*的\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})\s*汇款",
                RegexOptions.Singleline);
            if (match.Success)
            {
                var parsed = TryBuildWiseParseResult(match, out counterparty, out amount);
                fee = ParseWiseFeeOrZero(text, amount.t);
                isInternal = IsBrokerTransfer(counterparty);
                isPendingConversion = false;
                return parsed;
            }

            match = Regex.Match(
                text,
                @"您将会收到\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})\s*的汇款",
                RegexOptions.Singleline);
            if (match.Success)
            {
                counterparty = "Wise汇款收款方";
                amount = ParseWiseCurrency(match.Groups["amount"].Value, match.Groups["currency"].Value);
                fee = ParseWiseFeeOrZero(text, amount.t);
                isInternal = false;
                isPendingConversion = text.Contains("汇率为", StringComparison.Ordinal)
                    && text.Contains("手续费", StringComparison.Ordinal);
                return true;
            }

            counterparty = "";
            amount = new Currency();
            fee = new Currency();
            isInternal = false;
            isPendingConversion = false;
            return false;
        }

        private static Currency ParseWiseFeeOrZero(string text, CurrencyType defaultCurrency)
        {
            var match = Regex.Match(
                text,
                @"Wise\s*的手续费：\s*(?<amount>[\d,]+(?:\.\d+)?)\s*(?<currency>[A-Z]{3})",
                RegexOptions.Singleline);
            if (!match.Success)
                return new Currency(0, defaultCurrency);

            var fee = ParseWiseCurrency(match.Groups["amount"].Value, match.Groups["currency"].Value);
            if (fee.v < 0)
                throw new MailParseException($"Invalid Wise fee: {fee.v}/{fee.t}");
            return fee;
        }

        private static bool TryBuildWiseParseResult(Match match, out string counterparty, out Currency amount)
        {
            counterparty = NormalizeCounterparty(match.Groups["counterparty"].Value);
            amount = ParseWiseCurrency(match.Groups["amount"].Value, match.Groups["currency"].Value);
            return true;
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

        private WiseSettings ReadWiseSettings()
        {
            return new WiseSettings(
                ReadConfigList("wise:own_names"),
                ReadConfigList("wise:paypal_counterparty_names"),
                config["wise:sgd_own_incoming_source_account"] ?? "",
                config["wise:paypal_source_account"] ?? "",
                config["wise:rmb_unknown_recipient_name"] ?? "",
                config["wise:rmb_unknown_recipient_account"] ?? "",
                config["wise:received_funds_destination_account"] ?? "");
        }

        private List<string> ReadConfigList(string key)
        {
            return config.GetSection(key)
                .GetChildren()
                .Select(child => child.Value)
                .Where(value => !String.IsNullOrWhiteSpace(value))
                .Select(value => value!.Trim())
                .ToList();
        }

        private static WiseCounterpartyClassification ClassifyWiseCounterparty(
            WiseSettings settings,
            string counterparty,
            Currency amount,
            bool isIncoming)
        {
            var normalizedCounterparty = NormalizeNameForComparison(counterparty);
            if (settings.OwnNames.Any(name => NormalizeNameForComparison(name) == normalizedCounterparty))
            {
                var counterpartyText = isIncoming && amount.t == CurrencyType.SGD && !String.IsNullOrWhiteSpace(settings.SgdOwnIncomingSourceAccount)
                    ? FormatTransferCounterparty(settings.SgdOwnIncomingSourceAccount, WiseAccountName)
                    : counterparty;
                return new WiseCounterpartyClassification(counterpartyText, true);
            }

            if (settings.PaypalCounterpartyNames.Any(name => NormalizeNameForComparison(name) == normalizedCounterparty))
            {
                if (amount.t == CurrencyType.USD && !String.IsNullOrWhiteSpace(settings.PaypalSourceAccount))
                {
                    var counterpartyText = isIncoming
                        ? FormatTransferCounterparty(settings.PaypalSourceAccount, WiseAccountName)
                        : FormatTransferCounterparty(WiseAccountName, settings.PaypalSourceAccount);
                    return new WiseCounterpartyClassification(counterpartyText, true);
                }

                return new WiseCounterpartyClassification(counterparty, false);
            }

            if (!isIncoming
                && amount.t == CurrencyType.RMB
                && !String.IsNullOrWhiteSpace(settings.RmbUnknownRecipientName)
                && NormalizeNameForComparison(settings.RmbUnknownRecipientName) == normalizedCounterparty)
            {
                var counterpartyText = String.IsNullOrWhiteSpace(settings.RmbUnknownRecipientAccount)
                    ? counterparty
                    : FormatTransferCounterparty(WiseAccountName, settings.RmbUnknownRecipientAccount);
                return new WiseCounterpartyClassification(counterpartyText, true);
            }

            return new WiseCounterpartyClassification(counterparty, false);
        }

        private static string FormatTransferCounterparty(string sourceAccount, string destinationAccount)
        {
            return $"{sourceAccount} -> {destinationAccount}";
        }

        private static bool IsIncomingTransferSubject(string subject)
        {
            return subject.Contains("汇款", StringComparison.Ordinal)
                && subject.Contains("已收到", StringComparison.Ordinal);
        }

        private static string GetWiseIncomingReason(string subject, string counterparty)
        {
            if (IsBrokerTransfer(counterparty))
                return "出金";

            return IsIncomingTransferSubject(subject)
                ? "汇款收款"
                : "付款收款";
        }

        private static string GetWiseOutgoingReason(string subject, string counterparty)
        {
            if (IsBrokerTransfer(counterparty))
                return "入金";

            return subject.Contains("款项已汇出", StringComparison.Ordinal)
                ? "款项汇出"
                : "汇款已发送";
        }

        private static Currency ParseWiseCurrency(string value, string currencyCode)
        {
            var amount = decimal.Parse(
                value.Replace(",", "", StringComparison.Ordinal),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture);
            return new Currency(amount, ParseWiseCurrencyType(currencyCode));
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

        private static Currency Negate(Currency amount)
        {
            if (amount.v < 0)
                throw new MailParseException($"Wise amount should be positive before applying direction: {amount.v}/{amount.t}");

            return new Currency(-amount.v, amount.t);
        }

        private static bool IsBrokerTransfer(string counterparty)
        {
            return counterparty.Contains("Interactive Brokers", StringComparison.OrdinalIgnoreCase)
                || counterparty.Contains("IBKR", StringComparison.OrdinalIgnoreCase);
        }

        private static DateTime GetWiseMailDateTime(MimeMessage message)
        {
            var localTime = message.Date.LocalDateTime;
            return localTime == default ? GetMailDate(message) : localTime;
        }

        private static string BuildWiseStatementKey(MimeMessage message, string subject, string text)
        {
            var match = Regex.Match(subject + " " + text, @"#(?<id>\d+)");
            if (match.Success)
                return $"Wise_{match.Groups["id"].Value}";

            if (!String.IsNullOrWhiteSpace(message.MessageId))
                return $"WiseMessage_{message.MessageId.Trim()}";

            var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(subject + "\n" + text)));
            return $"WiseHash_{hash[..16]}";
        }

        private static string NormalizeWiseText(string text)
        {
            var withoutInvisibleCharacters = Regex.Replace(
                text,
                @"[\u00AD\u034F\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]+",
                " ");
            return NormalizeMailText(withoutInvisibleCharacters);
        }

        private static string NormalizeCounterparty(string counterparty)
        {
            return NormalizeWiseText(counterparty)
                .Trim(' ', '。', '，', ',', '.');
        }

        private static string NormalizeNameForComparison(string name)
        {
            return Regex.Replace(name, @"\s+", "").ToUpperInvariant();
        }

        private static string TrimForError(string text)
        {
            const int maxLength = 240;
            return text.Length <= maxLength ? text : text[..maxLength] + "...";
        }

        private static async Task RunWiseMailOperation(Func<CancellationToken, Task> operation)
        {
            using var cancellation = new CancellationTokenSource(WiseMailClientTimeoutMilliseconds);
            await operation(cancellation.Token).WaitAsync(WiseMailClientTimeout);
        }

        private static async Task<T> RunWiseMailOperation<T>(Func<CancellationToken, Task<T>> operation)
        {
            using var cancellation = new CancellationTokenSource(WiseMailClientTimeoutMilliseconds);
            return await operation(cancellation.Token).WaitAsync(WiseMailClientTimeout);
        }

        private sealed record WiseParsedMail(
            DateTime Time,
            string StatementKey,
            List<Record> Records,
            WiseImportPriority Priority);

        private sealed record WiseSettings(
            List<string> OwnNames,
            List<string> PaypalCounterpartyNames,
            string SgdOwnIncomingSourceAccount,
            string PaypalSourceAccount,
            string RmbUnknownRecipientName,
            string RmbUnknownRecipientAccount,
            string ReceivedFundsDestinationAccount);

        private sealed record WiseCounterpartyClassification(string CounterpartyText, bool IsInternal);

        private enum WiseImportPriority
        {
            Pending = 0,
            Normal = 1,
            Detailed = 2
        }
    }
}
