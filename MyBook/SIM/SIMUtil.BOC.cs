using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace MyBook
{
    partial class SIMUtil
    {
        private const string BOCSIMAccountType = "BOC";
        private const StatementImportProvider BOCSIMProvider = StatementImportProvider.BOCSIMSMS;
        private const string BOCSIMRowCodePrefix = "BOCSIM-";
        private static readonly Regex BOCSIMTransactionRegex = new(
            @"^您的借记卡账户(?<cardTail>\d{4})[，,]于(?<month>\d{1,2})月(?<day>\d{1,2})日(?<summary>.+?)(?<direction>支取|存入)人民币(?<amount>[+-]?\d[\d,]*(?:\.\d+)?)元[，,]交易后余额(?<balance>[+-]?\d[\d,]*(?:\.\d+)?)(?:元)?【中国银行】$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static bool IsBOCSender(string sender)
        {
            return String.Equals(sender, "95566", StringComparison.OrdinalIgnoreCase);
        }

        private SIMMessageProcessResult ParseBOCSIMMessage(SIMMessage message)
        {
            if (!TryParseBOCSIMTransaction(message, out var transaction))
                return new SIMMessageProcessResult("unsupported BOC SMS format", false);

            if (database is null)
                return new SIMMessageProcessResult("parsed BOC transaction but no database is configured", false);

            var cardAccount = database.GetAccountByTypeAndId(BOCSIMAccountType, transaction.CardTail);
            var postingAccount = database.GetPostingAccount(cardAccount);
            if (cardAccount.isCredit || postingAccount.isCredit)
                throw new InvalidOperationException("BOC SIM SMS for credit account is not allowed.");

            var statementKey = BuildBOCSIMStatementKey(postingAccount, transaction);
            if (database.IsStatementKeyImported(BOCSIMProvider, statementKey))
                return new SIMMessageProcessResult("duplicate BOC SMS transaction", true);

            if (IsMatchingBOCRecordAlreadyImported(postingAccount, transaction))
                return new SIMMessageProcessResult("matching BOC transaction already exists", true);

            var hasCurrentBalance = database.TryGetAccountBalance(
                postingAccount,
                transaction.Amount.t,
                out var currentBalance);
            var requiredBeginningBalance = new Currency(
                transaction.Balance.v - transaction.Amount.v,
                transaction.Balance.t);
            var record = BuildBOCSIMRecord(postingAccount, transaction, statementKey, message);
            var beginningBalance = new AccountBalance(
                postingAccount,
                hasCurrentBalance ? currentBalance : requiredBeginningBalance);
            var endingBalance = new AccountBalance(postingAccount, transaction.Balance);
            var saved = database.SaveStatementRecordsOnce(
                BOCSIMProvider,
                message.Time,
                [record],
                [endingBalance],
                statementKey,
                [beginningBalance],
                forceValidateBeginningBalances: hasCurrentBalance);

            return saved
                ? new SIMMessageProcessResult("imported BOC SMS transaction", true)
                : new SIMMessageProcessResult("duplicate BOC SMS transaction", true);
        }

        private static bool TryParseBOCSIMTransaction(SIMMessage message, out BOCSIMTransaction transaction)
        {
            var text = NormalizeBOCSIMText(message.Text);
            var match = BOCSIMTransactionRegex.Match(text);
            if (!match.Success)
            {
                transaction = null!;
                return false;
            }

            var direction = match.Groups["direction"].Value;
            var amount = new Currency(
                Math.Abs(ParseBOCSIMDecimal(match.Groups["amount"].Value)) * (direction == "支取" ? -1 : 1),
                CurrencyType.RMB);
            var balance = new Currency(
                ParseBOCSIMDecimal(match.Groups["balance"].Value),
                CurrencyType.RMB);
            var summary = NormalizeBOCSIMText(match.Groups["summary"].Value);
            transaction = new BOCSIMTransaction(
                match.Groups["cardTail"].Value,
                ResolveBOCSIMTransactionDate(
                    message.Time,
                    Int32.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture),
                    Int32.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture)),
                direction,
                summary,
                direction == "支取" ? "支出" : "收入",
                amount,
                balance);
            return true;
        }

        private bool IsMatchingBOCRecordAlreadyImported(Account postingAccount, BOCSIMTransaction transaction)
        {
            return new[] { BOCSIMProvider, StatementImportProvider.BOCBillMail }
                .SelectMany(provider => database!.GetStatementRecords(
                    provider,
                    postingAccount,
                    transaction.TransactionDate.AddDays(-1),
                    transaction.TransactionDate.AddDays(2)))
                .Any(record => IsMatchingBOCSIMRecord(record, transaction));
        }

        private static bool IsMatchingBOCSIMRecord(Record record, BOCSIMTransaction transaction)
        {
            if (record.v != transaction.Amount.v
                || record.t != transaction.Amount.t
                || (record.postingDate ?? record.date).Date != transaction.TransactionDate)
            {
                return false;
            }

            if (TryParseBOCSIMSourceBalance(record.Source, out var balance) && balance == transaction.Balance)
                return true;

            return NormalizeBOCSIMComparableText(record.DestAccount)
                == NormalizeBOCSIMComparableText(transaction.Summary);
        }

        private static Record BuildBOCSIMRecord(
            Account postingAccount,
            BOCSIMTransaction transaction,
            string statementKey,
            SIMMessage message)
        {
            var record = new Record
            {
                Account = postingAccount,
                date = transaction.TransactionDate,
                postingDate = transaction.TransactionDate,
                updateTime = DateTime.Now,
                DestAccount = transaction.Summary,
                Source = BuildBOCSIMSource(transaction, statementKey, message),
                Reason = transaction.Reason,
                DescCurrency = transaction.Amount
            };
            record.CopyFrom(transaction.Amount);
            return record;
        }

        private static string BuildBOCSIMStatementKey(Account postingAccount, BOCSIMTransaction transaction)
        {
            var canonical = String.Join("|",
                postingAccount.name,
                transaction.TransactionDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                transaction.Direction,
                transaction.Amount.v.ToString(CultureInfo.InvariantCulture),
                transaction.Amount.t,
                transaction.Balance.v.ToString(CultureInfo.InvariantCulture),
                transaction.Summary);
            return BOCSIMRowCodePrefix + ComputeBOCSIMSha256(Encoding.UTF8.GetBytes(canonical))[..24];
        }

        private static string BuildBOCSIMSource(
            BOCSIMTransaction transaction,
            string statementKey,
            SIMMessage message)
        {
            var source = String.Join("; ",
                $"code={statementKey}",
                "BOCSIMSMS",
                $"smsTime={message.Time:yyyy-MM-dd HH:mm:ss}",
                $"smsBalance={transaction.Balance.t}:{transaction.Balance.v.ToString(CultureInfo.InvariantCulture)}",
                $"row={NormalizeBOCSIMText(message.Text)}");
            return LimitBOCSIMRecordText(source);
        }

        private static bool TryParseBOCSIMSourceBalance(string source, out Currency balance)
        {
            var match = Regex.Match(
                source ?? "",
                @"(?:^|;\s*)smsBalance=(?<currency>[A-Za-z0-9_]+):(?<amount>[+-]?\d[\d,]*(?:\.\d+)?)",
                RegexOptions.CultureInvariant);
            if (!match.Success
                || !Enum.TryParse<CurrencyType>(match.Groups["currency"].Value, out var currencyType))
            {
                balance = null!;
                return false;
            }

            balance = new Currency(ParseBOCSIMDecimal(match.Groups["amount"].Value), currencyType);
            return true;
        }

        private static DateTime ResolveBOCSIMTransactionDate(DateTime smsTime, int month, int day)
        {
            var transactionDate = new DateTime(smsTime.Year, month, day);
            if (transactionDate > smsTime.Date.AddDays(1))
                transactionDate = transactionDate.AddYears(-1);
            return transactionDate;
        }

        private static decimal ParseBOCSIMDecimal(string value)
        {
            return Decimal.Parse(
                value.Replace(",", "", StringComparison.Ordinal),
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture);
        }

        private static string NormalizeBOCSIMText(string value)
        {
            return Regex.Replace(value ?? "", @"\s+", " ", RegexOptions.CultureInvariant).Trim();
        }

        private static string NormalizeBOCSIMComparableText(string value)
        {
            return Regex.Replace(value ?? "", @"[\s,，。.;；:：()（）\[\]【】\-—_]+", "", RegexOptions.CultureInvariant).Trim();
        }

        private static string LimitBOCSIMRecordText(string text)
        {
            const int maxLength = 1024;
            return text.Length <= maxLength ? text : text[..maxLength];
        }

        private static string ComputeBOCSIMSha256(byte[] bytes)
        {
            return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        }

        private sealed record BOCSIMTransaction(
            string CardTail,
            DateTime TransactionDate,
            string Direction,
            string Summary,
            string Reason,
            Currency Amount,
            Currency Balance);
    }
}
