using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MailKit.Search;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    private const string IFastSender = "noreply@ifastgb.com";
    private const StatementImportProvider IFastProvider = StatementImportProvider.IFastMail;
    private const string IFastReceiptSubject = "Your account has received a new payment!";
    private const string IFastConversionSubject = "Your Currency Conversion order has been accepted";
    private const string IFastPaymentSubject = "QR Payment is successful";
    private const string IFastMoneyPattern = @"(?<currency>GBP|USD|EUR|HKD|SGD|CNY|RMB|HK\$|S\$|US\$|£|€)\s*(?<amount>(?:\d+|\d{1,3}(?:,\d{3})+)\.\d{2})(?![\d.])";

    public async Task FetchIFastMessages(DateTime since)
    {
        ImportIFastInitialStatements();
        // Import times are stored as dates. Include that day and deduplicate its successful mails.
        // IMAP header dates may use another timezone, so keep a small date-boundary margin.
        var searchDate = since.ToUniversalTime().AddHours(-14).Date;
        Console.WriteLine($"IFast mail scan from: {since:yyyy-MM-dd HH:mm:ss}");
        await RunWithMailSessionScope(async () =>
        {
            var messages = await SearchMessagesFromMailbox(
                CreateYahooMailbox() with { Proxy = null }, "IFast transactions",
                SearchQuery.FromContains(IFastSender).And(SearchQuery.SentSince(searchDate)),
                summary => SummaryIsFrom(summary, IFastSender)
                    && GetSummaryDateTime(summary) >= since
                    && IsIFastTransactionSubject(summary.Envelope?.Subject ?? "")
                    && !database.IsStatementKeyImported(IFastProvider, IFastMessageKey(summary.Envelope?.MessageId)),
                message => IsFrom(message, IFastSender), GetMailDateTime).ConfigureAwait(false);
            foreach (var message in messages)
                ImportIFastMessage(message);
            await UpdateIFastInterestRates().ConfigureAwait(false);
        }).ConfigureAwait(false);
        await GenerateIFastInterest().ConfigureAwait(false);
        ValidateIFastStatements();
    }

    private static bool IsIFastTransactionSubject(string subject)
    {
        if (subject is IFastReceiptSubject or IFastConversionSubject or IFastPaymentSubject)
            return true;
        if (Regex.IsMatch(subject, @"^(?:Your (?:Currency Conversion|payment|transfer|refund|interest)|QR Payment)", RegexOptions.IgnoreCase))
            throw new MailParseException("Unsupported IFast transaction notification subject.");
        return false;
    }

    private void ImportIFastMessage(MimeMessage message)
    {
        if (!IsFrom(message, IFastSender))
            throw new MailParseException("Invalid IFast sender.");
        if (database.IsStatementKeyImported(IFastProvider, IFastMessageKey(message.MessageId)))
            return;
        var (key, records) = ParseIFastMessage(message);
        if (database.IsStatementKeyImported(IFastProvider, key))
            return;
        var baselineEnd = GetIFastInitializationRange().End;
        if (records.All(record => IFastBankDate(record) < baselineEnd))
        {
            var actual = GetIFastAccountRecords();
            foreach (var record in records)
                FindIFastMatch(record, actual, true);
            database.MarkStatementProcessedOnce(IFastProvider, GetMailDateTime(message), key);
            return;
        }
        if (records.Any(record => IFastBankDate(record) < baselineEnd))
            throw new MailParseException("IFast mail crosses the initialization boundary.");
        database.SaveStatementRecordsOnce(IFastProvider, GetMailDateTime(message), records, statementKey: key);
        Console.WriteLine($"Imported IFast transaction: records={records.Count}");
    }

    private (string Key, List<Record> Records) ParseIFastMessage(MimeMessage message)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(message.HtmlBody ?? "");
        var text = NormalizeMailText(WebUtility.HtmlDecode(message.TextBody ?? doc.DocumentNode.InnerText));
        var time = GetMailDateTime(message);
        var key = IFastMessageKey(message.MessageId);
        if (message.Subject == IFastReceiptSubject)
        {
            var match = MatchIFast(text, @"Payment Ref No\.: (?<ref>\d+) Payment Received Date: (?<date>\d{2} [A-Za-z]{3} \d{4}) Payment Amount: "
                + IFastMoneyPattern + @" From Payer: (?<payer>.+?)(?: Important Information:| iFAST Global Bank Limited|$)");
            var amount = ParseIFastMoney(match);
            var date = DateTime.ParseExact(match.Groups["date"].Value, "dd MMM yyyy", CultureInfo.InvariantCulture);
            key = $"IFast-receipt-{match.Groups["ref"].Value}";
            return (key, [BuildIFastRecord(amount, date, "转入", match.Groups["payer"].Value, key)]);
        }
        if (message.Subject == IFastConversionSubject)
        {
            var reference = MatchIFast(text, @"Order Ref No: (?<ref>\d+)").Groups["ref"].Value;
            if (MatchIFast(text, @"Status: (?<status>\w+)").Groups["status"].Value != "Completed")
                throw new MailParseException("IFast currency conversion is not completed.");
            var dateMatch = MatchIFast(text, @"Order Acceptance Date: (?<date>\d{2} [A-Za-z]{3} \d{4} \d{2}:\d{2} [AP]M) (?<zone>BST|GMT)");
            var localDate = DateTime.ParseExact(dateMatch.Groups["date"].Value, "dd MMM yyyy hh:mm tt", CultureInfo.InvariantCulture);
            time = new DateTimeOffset(DateTime.SpecifyKind(localDate, DateTimeKind.Unspecified),
                TimeSpan.FromHours(dateMatch.Groups["zone"].Value == "BST" ? 1 : 0)).LocalDateTime;
            var sell = ParseIFastMoney(MatchIFast(text, "Selling Amount: " + IFastMoneyPattern));
            var buy = ParseIFastMoney(MatchIFast(text, "Buying Amount: " + IFastMoneyPattern));
            if (sell.t == buy.t)
                throw new MailParseException("IFast conversion currencies must differ.");
            sell.v = -sell.v;
            var code = $"BALANCE-IFAST-{reference}";
            var accountName = GetIFastAccount().name;
            return ($"IFast-conversion-{reference}",
                [BuildIFastRecord(sell, time, "换汇", accountName, code, true),
                 BuildIFastRecord(buy, time, "换汇", accountName, code, true)]);
        }
        if (message.Subject == IFastPaymentSubject)
        {
            var match = MatchIFast(text, @"Your QR payment of " + IFastMoneyPattern
                + @" \((?<originalCurrency>[A-Z]{3}) (?<originalAmount>\d[\d,]*\.\d{2})\) to (?<merchant>.+?) is successful\.");
            var amount = ParseIFastMoney(match);
            amount.v = -amount.v;
            var record = BuildIFastRecord(amount, time, "消费", match.Groups["merchant"].Value, key);
            record.DescCurrency = new Currency(-Decimal.Parse(match.Groups["originalAmount"].Value, NumberStyles.Number, CultureInfo.InvariantCulture),
                match.Groups["originalCurrency"].Value);
            return (key, [record]);
        }
        throw new MailParseException("Unsupported IFast mail format.");
    }

    private Record BuildIFastRecord(Currency amount, DateTime time, string reason, string counterparty, string code, bool isInternal = false)
    {
        var account = GetIFastAccount();
        var record = new Record
        {
            Account = account, _account_Id = account.Id,
            date = time, postingDate = time, updateTime = DateTime.Now,
            Reason = reason, DestAccount = counterparty,
            DescCurrency = new Currency(amount.v, amount.t),
            Source = $"code={code}; IFast mail", isInternal = isInternal
        };
        record.CopyFrom(amount);
        ResolveIFastTransferAccount(record, counterparty, counterparty);
        return record;
    }

    private void ResolveIFastTransferAccount(Record record, string counterparty, string? counterpartyName = null)
    {
        if (record.Reason is not ("转入" or "转出")) return;
        // 保留原有 IBAN 查找；只有明确的对方姓名字段才能参与本人别名识别。
        var ibans = Regex.Matches(counterparty, @"\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
            .Select(match => match.Value).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (ibans.Count > 1)
            throw new MailParseException("Ambiguous IFast counterparty IBAN.");
        var account = ibans.Count == 0 ? null : database.FindAccountByInternalCardNo(ibans[0]);
        var match = database.ResolveTransferCounterparty(account,
            ibans.Concat(counterpartyName is null ? [] : new[] { counterpartyName }).ToArray(), counterpartyName);
        if (match.Account?.Id == record._account_Id)
            throw new MailParseException("IFast transfer counterparty resolves to the source account.");
        DatabaseUtil.ApplyTransferCounterparty(record, match);
    }

    private Account GetIFastAccount()
    {
        var accounts = database.GetAccountsByNamePrefix("IFAST_");
        if (accounts.Count != 1)
            throw new InvalidOperationException("IFast requires exactly one configured account.");
        return accounts[0];
    }

    private static string IFastMessageKey(string? messageId)
    {
        if (String.IsNullOrWhiteSpace(messageId))
            throw new MailParseException("IFast transaction mail has no Message-Id.");
        return "IFast-mail-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(messageId)));
    }

    private static Match MatchIFast(string text, string pattern)
    {
        var matches = Regex.Matches(text, pattern, RegexOptions.CultureInvariant);
        if (matches.Count != 1)
            throw new MailParseException("IFast transaction field is missing or ambiguous.");
        return matches[0];
    }

    private static Currency ParseIFastMoney(Match match)
    {
        var currency = match.Groups["currency"].Value switch
        {
            "£" => "GBP", "€" => "EUR", "HK$" => "HKD", "S$" => "SGD", "US$" => "USD",
            var value => value
        };
        var amount = Decimal.Parse(match.Groups["amount"].Value, NumberStyles.Number, CultureInfo.InvariantCulture);
        if (amount <= 0)
            throw new MailParseException("IFast transaction amount must be positive.");
        return new Currency(amount, currency);
    }
}
