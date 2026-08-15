using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace MyBook
{
    partial class SIMUtil
    {
        private const string ICBCSIMAccountType = "ICBC";
        private const StatementImportProvider ICBCSIMProvider = StatementImportProvider.ICBCSIMSMS;
        private const string ICBCSIMRowCodePrefix = "ICBCSIM-";
        private const string ICBCSIMCompensationCodePrefix = "ICBCSIMCompensation-";
        private static readonly Regex ICBCSIMTransactionRegex = new(
            @"尾号(?<cardTail>\d{4})卡(?<month>\d{1,2})月(?<day>\d{1,2})日(?<hour>\d{1,2}):(?<minute>\d{2})(?:手机银行|营业网点)?(?<direction>支出|收入)[(（](?<summary>[^)）]+)[)）](?<amount>[+-]?\d[\d,]*(?:\.\d+)?)元[，,]\s*余额(?<balance>[+-]?\d[\d,]*(?:\.\d+)?)元",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private static readonly Regex ICBCSIMIgnoredNotificationRegex = new(
            @"^您尾号\d{4}(?:卡人民币卡片临时额度将于.+到期|信用卡信用额度已恢复为人民币[\d,.]+元).+【工商银行】$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static bool IsICBCSender(string sender)
        {
            return String.Equals(sender, "95588", StringComparison.OrdinalIgnoreCase);
        }

        private static IReadOnlyList<SIMMessage> OrderSIMMessagesForProcessing(IReadOnlyList<SIMMessage> messages)
        {
            var items = messages
                .Select(message =>
                {
                    ICBCSIMTransaction? transaction = null;
                    if (IsICBCSender(message.Sender)
                        && TryParseICBCSIMTransaction(message, out var parsedTransaction))
                    {
                        transaction = parsedTransaction;
                    }

                    return new ICBCSIMMessageOrdering(
                        message,
                        transaction,
                        transaction?.TransactionTime ?? message.Time,
                        0);
                })
                .ToList();

            foreach (var group in items
                .Where(item => item.Transaction is not null)
                .GroupBy(item => (item.Transaction!.CardTail, item.Transaction.TransactionTime)))
            {
                if (group.Count() < 2)
                    continue;

                var chain = OrderICBCSIMBalanceChain(group.ToList());
                for (var index = 0; index < chain.Count; index++)
                    chain[index].ChainOrder = index;
            }

            return items
                .OrderBy(item => item.EffectiveTime)
                .ThenBy(item => item.Transaction?.CardTail ?? "", StringComparer.Ordinal)
                .ThenBy(item => item.ChainOrder)
                .ThenBy(item => item.Message.Index)
                .Select(item => item.Message)
                .ToList();
        }

        private static List<ICBCSIMMessageOrdering> OrderICBCSIMBalanceChain(List<ICBCSIMMessageOrdering> items)
        {
            var firstCandidates = items
                .Where(item => !items.Any(other =>
                    !ReferenceEquals(other, item)
                    && CurrencyEquals(other.Transaction!.Balance, GetICBCSIMRequiredBeginningBalance(item.Transaction!))))
                .ToList();
            if (firstCandidates.Count != 1)
                throw new MailParseException(
                    $"Cannot uniquely order {items.Count} ICBC SMS transactions at {items[0].EffectiveTime:yyyy-MM-dd HH:mm}.");

            var result = new List<ICBCSIMMessageOrdering> { firstCandidates[0] };
            var remaining = items.Where(item => !ReferenceEquals(item, firstCandidates[0])).ToList();
            while (remaining.Count > 0)
            {
                var endingBalance = result[^1].Transaction!.Balance;
                var nextCandidates = remaining
                    .Where(item => CurrencyEquals(
                        endingBalance,
                        GetICBCSIMRequiredBeginningBalance(item.Transaction!)))
                    .ToList();
                if (nextCandidates.Count != 1)
                    throw new MailParseException(
                        $"Broken or ambiguous ICBC SMS balance chain at {items[0].EffectiveTime:yyyy-MM-dd HH:mm}.");

                var next = nextCandidates[0];
                result.Add(next);
                remaining.Remove(next);
            }

            return result;
        }

        private static Currency GetICBCSIMRequiredBeginningBalance(ICBCSIMTransaction transaction)
        {
            return new Currency(
                transaction.Balance.v - transaction.Amount.v,
                transaction.Balance.t);
        }

        private static bool CurrencyEquals(Currency left, Currency right)
        {
            return left.t == right.t && left.v == right.v;
        }

        private SIMMessageProcessResult ParseICBCSIMMessage(SIMMessage message)
        {
            if (ICBCSIMIgnoredNotificationRegex.IsMatch(NormalizeICBCSIMText(message.Text)))
                return new SIMMessageProcessResult("ignored ICBC credit-limit notification", true, true);

            if (!TryParseICBCSIMTransaction(message, out var transaction))
                throw new MailParseException(
                    $"Unsupported ICBC SMS format: sender={message.Sender}, index={message.Index}");

            if (database is null)
                return new SIMMessageProcessResult("parsed ICBC transaction but no database is configured", false);

            var cardAccount = database.GetAccountByTypeAndId(ICBCSIMAccountType, transaction.CardTail);
            var postingAccount = database.GetPostingAccount(cardAccount);
            if (postingAccount.isCredit)
                throw new InvalidOperationException("ICBC SIM SMS for credit account is not allowed.");

            var statementKey = BuildICBCSIMStatementKey(postingAccount, transaction);
            if (database.IsStatementKeyImported(ICBCSIMProvider, statementKey))
                return new SIMMessageProcessResult("duplicate ICBC SMS transaction", true);

            if (IsMatchingICBCRecordAlreadyImported(postingAccount, transaction))
                return new SIMMessageProcessResult("matching ICBC transaction already exists", true);

            var hasCurrentBalance = database.TryGetAccountBalance(
                postingAccount,
                transaction.Amount.t,
                out var currentBalance);
            var requiredBeginningBalance = new Currency(transaction.Balance.v - transaction.Amount.v, transaction.Balance.t);
            var record = BuildICBCSIMRecord(postingAccount, transaction, statementKey, message);
            var records = new List<Record>();
            var compensation = hasCurrentBalance
                ? BuildICBCSIMCompensationRecordIfNeeded(
                    postingAccount,
                    transaction,
                    statementKey,
                    message,
                    currentBalance,
                    requiredBeginningBalance)
                : null;
            if (compensation is not null)
                records.Add(compensation);

            records.Add(record);
            var beginningBalance = new AccountBalance(postingAccount, hasCurrentBalance ? currentBalance : requiredBeginningBalance);
            var endingBalance = new AccountBalance(postingAccount, transaction.Balance);
            var saved = database.SaveStatementRecordsOnce(
                ICBCSIMProvider,
                message.Time,
                records,
                [endingBalance],
                statementKey,
                [beginningBalance],
                forceValidateBeginningBalances: hasCurrentBalance);

            return saved
                ? new SIMMessageProcessResult(compensation is null
                    ? "imported ICBC SMS transaction"
                    : "imported ICBC SMS transaction with balance compensation", true)
                : new SIMMessageProcessResult("duplicate ICBC SMS transaction", true);
        }

        private static bool TryParseICBCSIMTransaction(SIMMessage message, out ICBCSIMTransaction transaction)
        {
            var text = NormalizeICBCSIMText(message.Text);
            var match = ICBCSIMTransactionRegex.Match(text);
            if (!match.Success)
            {
                transaction = null!;
                return false;
            }

            var direction = match.Groups["direction"].Value;
            var sign = direction == "支出" ? -1 : 1;
            var amount = new Currency(
                Math.Abs(ParseICBCSIMDecimal(match.Groups["amount"].Value)) * sign,
                CurrencyType.RMB);
            var balance = new Currency(
                ParseICBCSIMDecimal(match.Groups["balance"].Value),
                CurrencyType.RMB);
            var summary = NormalizeICBCSIMText(match.Groups["summary"].Value);
            var reason = InferICBCSIMReason(direction, summary);
            var destAccount = BuildICBCSIMDestAccount(summary, reason);
            transaction = new ICBCSIMTransaction(
                match.Groups["cardTail"].Value,
                ResolveICBCSIMTransactionTime(
                    message.Time,
                    Int32.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture),
                    Int32.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture),
                    Int32.Parse(match.Groups["hour"].Value, CultureInfo.InvariantCulture),
                    Int32.Parse(match.Groups["minute"].Value, CultureInfo.InvariantCulture)),
                direction,
                summary,
                reason,
                destAccount,
                amount,
                balance);
            return true;
        }

        private bool IsMatchingICBCRecordAlreadyImported(Account postingAccount, ICBCSIMTransaction transaction)
        {
            return new[]
                {
                    ICBCSIMProvider,
                    StatementImportProvider.ICBCHistoryDetailMail,
                    StatementImportProvider.ICBCBillMail
                }
                .SelectMany(provider => database!.GetStatementRecords(
                    provider,
                    postingAccount,
                    transaction.TransactionTime.Date.AddDays(-1),
                    transaction.TransactionTime.Date.AddDays(2)))
                .Any(record => IsMatchingICBCSIMRecord(record, transaction));
        }

        private static bool IsMatchingICBCSIMRecord(Record record, ICBCSIMTransaction transaction)
        {
            if (record.v != transaction.Amount.v || record.t != transaction.Amount.t)
                return false;

            var recordTime = record.postingDate ?? record.date;
            if (recordTime.Year != transaction.TransactionTime.Year
                || recordTime.Month != transaction.TransactionTime.Month
                || recordTime.Day != transaction.TransactionTime.Day
                || recordTime.Hour != transaction.TransactionTime.Hour
                || recordTime.Minute != transaction.TransactionTime.Minute)
                return false;

            if (TryParseICBCSIMSourceBalance(record.Source, out var balance) && balance == transaction.Balance)
                return true;

            var recordDest = NormalizeICBCSIMComparableText(record.DestAccount);
            var transactionDest = NormalizeICBCSIMComparableText(transaction.DestAccount);
            return !String.IsNullOrWhiteSpace(recordDest)
                && !String.IsNullOrWhiteSpace(transactionDest)
                && (recordDest.Contains(transactionDest, StringComparison.Ordinal)
                    || transactionDest.Contains(recordDest, StringComparison.Ordinal));
        }

        private static Record BuildICBCSIMRecord(
            Account postingAccount,
            ICBCSIMTransaction transaction,
            string statementKey,
            SIMMessage message)
        {
            var record = new Record
            {
                Account = postingAccount,
                date = transaction.TransactionTime,
                postingDate = transaction.TransactionTime,
                updateTime = DateTime.Now,
                DestAccount = transaction.DestAccount,
                Source = BuildICBCSIMSource(transaction, statementKey, message),
                Reason = transaction.Reason,
                isInternal = IsICBCSIMCreditCardRepayment(transaction),
                DescCurrency = transaction.Amount
            };
            record.CopyFrom(transaction.Amount);
            return record;
        }

        private static bool IsICBCSIMCreditCardRepayment(ICBCSIMTransaction transaction)
        {
            return transaction.Direction == "支出"
                && transaction.Reason == "还款";
        }

        private Record? BuildICBCSIMCompensationRecordIfNeeded(
            Account postingAccount,
            ICBCSIMTransaction transaction,
            string statementKey,
            SIMMessage message,
            Currency currentBalance,
            Currency requiredBeginningBalance)
        {
            if (currentBalance.t != requiredBeginningBalance.t)
                throw new InvalidOperationException(
                    $"ICBC SIM balance currency mismatch: current={currentBalance.t}, required={requiredBeginningBalance.t}");

            var compensationAmount = requiredBeginningBalance.v - currentBalance.v;
            if (compensationAmount == 0)
                return null;

            var compensation = new Currency(compensationAmount, requiredBeginningBalance.t);
            var record = new Record
            {
                Account = postingAccount,
                date = transaction.TransactionTime.AddSeconds(-1),
                postingDate = transaction.TransactionTime.AddSeconds(-1),
                updateTime = DateTime.Now,
                DestAccount = "Inferred from SMS balance",
                Source = BuildICBCSIMCompensationSource(
                    transaction,
                    statementKey,
                    message,
                    currentBalance,
                    requiredBeginningBalance,
                    compensation),
                Reason = "Missing small transactions",
                DescCurrency = compensation
            };
            record.CopyFrom(compensation);
            return record;
        }

        private static string BuildICBCSIMStatementKey(Account postingAccount, ICBCSIMTransaction transaction)
        {
            var canonical = String.Join("|",
                postingAccount.name,
                transaction.TransactionTime.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
                transaction.Direction,
                transaction.Amount.v.ToString(CultureInfo.InvariantCulture),
                transaction.Amount.t,
                transaction.Balance.v.ToString(CultureInfo.InvariantCulture),
                transaction.Summary);
            return ICBCSIMRowCodePrefix + ComputeICBCSIMSha256(Encoding.UTF8.GetBytes(canonical))[..24];
        }

        private static string BuildICBCSIMSource(
            ICBCSIMTransaction transaction,
            string statementKey,
            SIMMessage message)
        {
            var source = String.Join("; ",
                $"code={statementKey}",
                "ICBCSIMSMS",
                $"smsTime={message.Time:yyyy-MM-dd HH:mm:ss}",
                $"smsBalance={transaction.Balance.t}:{transaction.Balance.v.ToString(CultureInfo.InvariantCulture)}",
                $"row={NormalizeICBCSIMText(message.Text)}");
            return LimitICBCSIMRecordText(source);
        }

        private static string BuildICBCSIMCompensationSource(
            ICBCSIMTransaction transaction,
            string statementKey,
            SIMMessage message,
            Currency currentBalance,
            Currency requiredBeginningBalance,
            Currency compensation)
        {
            var compensationCode = BuildICBCSIMCompensationCode(
                statementKey,
                currentBalance,
                requiredBeginningBalance,
                compensation);
            var source = String.Join("; ",
                $"code={compensationCode}",
                "ICBCSIMSMS",
                "ICBCSIMCompensation",
                $"generatedFrom={statementKey}",
                $"smsTime={message.Time:yyyy-MM-dd HH:mm:ss}",
                $"currentBalance={currentBalance.t}:{currentBalance.v.ToString(CultureInfo.InvariantCulture)}",
                $"smsRequiredBeginning={requiredBeginningBalance.t}:{requiredBeginningBalance.v.ToString(CultureInfo.InvariantCulture)}",
                $"smsBalance={requiredBeginningBalance.t}:{requiredBeginningBalance.v.ToString(CultureInfo.InvariantCulture)}",
                $"smsFinalBalance={transaction.Balance.t}:{transaction.Balance.v.ToString(CultureInfo.InvariantCulture)}",
                $"row={NormalizeICBCSIMText(message.Text)}");
            return LimitICBCSIMRecordText(source);
        }

        private static string BuildICBCSIMCompensationCode(
            string statementKey,
            Currency currentBalance,
            Currency requiredBeginningBalance,
            Currency compensation)
        {
            var canonical = String.Join("|",
                statementKey,
                currentBalance.v.ToString(CultureInfo.InvariantCulture),
                currentBalance.t,
                requiredBeginningBalance.v.ToString(CultureInfo.InvariantCulture),
                requiredBeginningBalance.t,
                compensation.v.ToString(CultureInfo.InvariantCulture),
                compensation.t);
            return ICBCSIMCompensationCodePrefix + ComputeICBCSIMSha256(Encoding.UTF8.GetBytes(canonical))[..24];
        }

        private static bool TryParseICBCSIMSourceBalance(string source, out Currency balance)
        {
            var match = Regex.Match(
                source ?? "",
                @"(?:^|;\s*)(?:smsBalance|historyBalance)=(?<currency>[A-Za-z0-9_]+):(?<amount>[+-]?\d[\d,]*(?:\.\d+)?)",
                RegexOptions.CultureInvariant);
            if (!match.Success
                || !Enum.TryParse<CurrencyType>(match.Groups["currency"].Value, out var currencyType))
            {
                balance = null!;
                return false;
            }

            balance = new Currency(ParseICBCSIMDecimal(match.Groups["amount"].Value), currencyType);
            return true;
        }

        private static string InferICBCSIMReason(string direction, string summary)
        {
            foreach (var prefix in new[] { "消费", "缴费", "退款", "还款", "转账", "转帐", "利息", "费用", "工资" })
            {
                if (summary.StartsWith(prefix, StringComparison.Ordinal))
                    return prefix == "转帐" ? "转账" : prefix;
            }

            return direction == "收入" ? "收入" : "支出";
        }

        private static string BuildICBCSIMDestAccount(string summary, string reason)
        {
            var value = summary;
            if (value.StartsWith(reason, StringComparison.Ordinal))
                value = value[reason.Length..];
            else if (reason == "转账" && value.StartsWith("转帐", StringComparison.Ordinal))
                value = value["转帐".Length..];

            value = value.Trim(' ', '-', '—', '_', ':', '：');
            return String.IsNullOrWhiteSpace(value) ? summary : value;
        }

        private static DateTime ResolveICBCSIMTransactionTime(
            DateTime smsTime,
            int month,
            int day,
            int hour,
            int minute)
        {
            var transactionTime = new DateTime(smsTime.Year, month, day, hour, minute, 0);
            if (transactionTime > smsTime.AddDays(1))
                transactionTime = transactionTime.AddYears(-1);
            return transactionTime;
        }

        private static decimal ParseICBCSIMDecimal(string value)
        {
            return Decimal.Parse(
                value.Replace(",", "", StringComparison.Ordinal),
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture);
        }

        private static string NormalizeICBCSIMText(string value)
        {
            return Regex.Replace(value ?? "", @"\s+", " ", RegexOptions.CultureInvariant).Trim();
        }

        private static string NormalizeICBCSIMComparableText(string value)
        {
            return Regex.Replace(value ?? "", @"[\s,，。.;；:：()（）\[\]【】\-—_]+", "", RegexOptions.CultureInvariant).Trim();
        }

        private static string LimitICBCSIMRecordText(string text)
        {
            const int maxLength = 1024;
            return text.Length <= maxLength ? text : text[..maxLength];
        }

        private static string ComputeICBCSIMSha256(byte[] bytes)
        {
            return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        }

        private sealed record ICBCSIMTransaction(
            string CardTail,
            DateTime TransactionTime,
            string Direction,
            string Summary,
            string Reason,
            string DestAccount,
            Currency Amount,
            Currency Balance);

        private sealed class ICBCSIMMessageOrdering(
            SIMMessage message,
            ICBCSIMTransaction? transaction,
            DateTime effectiveTime,
            int chainOrder)
        {
            public SIMMessage Message { get; } = message;
            public ICBCSIMTransaction? Transaction { get; } = transaction;
            public DateTime EffectiveTime { get; } = effectiveTime;
            public int ChainOrder { get; set; } = chainOrder;
        }
    }
}
