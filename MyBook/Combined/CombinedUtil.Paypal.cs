using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook;

partial class CombinedUtil
{
    internal const string PayPalPrefix = "PayPalCombined/";
    private static readonly SemaphoreSlim payPalLock = new(1, 1);
    private const string MoneyPattern = @"\$?\s*(?<value>\d[\d,]*\.\d{2})\s*(?<currency>USD|SGD|HKD|GBP|EUR|JPY|CNY|RMB)\b";
    private const string CardPattern = @"(?:VISA|Master\s*Card|American Express|AMEX|Discover|UnionPay|银联)\s*(?:信用卡|借记卡|Credit|Debit)?\s*(?:x\s*[-*]|[-*•·])+\s*(?<tail>\d{4})";
    internal static string PayPalHash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    private static MailParseException PayPalError(string message) => new("PayPal combined: " + message);
    private static string Clean(string value) => Regex.Replace(value, @"[\s\p{Cf}]+", " ").Trim();
    private static string PartyKey(string value) => Regex.Replace(value, @"[^\p{L}\p{N}]", "").ToUpperInvariant();
    private static bool PartyMatches(string party, string text)
    {
        var key = PartyKey(party);
        if (key.Length < 3) return false;
        var normalized = PartyKey(text);
        if (normalized.Contains(key, StringComparison.Ordinal)) return true;
        var parts = Regex.Matches(party, @"[\p{L}\p{N}]+").Select(m => PartyKey(m.Value)).ToList();
        return parts.Count > 1 && parts.All(part => normalized.Contains(part, StringComparison.Ordinal));
    }

    internal sealed record PayPalMail(int AccountId, string MessageId, DateTimeOffset Date, string Subject, string Text);
    internal sealed record PayPalAmount(decimal Value, CurrencyType Currency);
    internal enum PayPalKind { Receive, Send, CardPurchase, Withdraw, Refund }
    internal sealed record PayPalNotice(PayPalMail Mail, string Key, string NativeId, PayPalKind Kind,
        PayPalAmount Gross, decimal Fee, decimal Net, string Party, string Card, PayPalAmount? CardAmount,
        string BankIdentifier, bool Nexus, bool ForeignExchange);
    internal sealed record PayPalItemState(int ItemRowId, int AccountId, string RemoteAccountId,
        CurrencyType Currency, decimal Balance, string Cursor, List<JObject> Transactions, bool IsInitial = false);
    internal sealed class PayPalState
    {
        public List<PayPalItemState> Items { get; set; } = [];
        public List<PayPalMail> Mails { get; set; } = [];
    }
    internal sealed class PayPalPlan
    {
        public List<Record> Records { get; } = [];
        public List<RecordSourceSupplement> Supplements { get; } = [];
        public List<string> Problems { get; } = [];
        public Dictionary<int, string> ExpectedBankRecords { get; } = [];
        public Dictionary<(int AccountId, CurrencyType Currency), decimal> ExpectedBalances { get; } = [];
        public Dictionary<int, int> ExpectedItemAccounts { get; } = [];
        public Dictionary<int, DateTime> ReportDates { get; } = [];
        public List<PayPalPair> Pairs { get; } = [];
        public int IgnoredCards { get; set; }
        public int MatchedTransactions { get; set; }
    }
    internal sealed record PayPalPair(string LeftSource, string? RightSource = null, int? BankRecordId = null);
    internal sealed record PayPalImport(StatementImportProvider Provider, int AccountId, string? PreviousKey, string Key);
    internal static readonly StatementImportProvider[] PayPalProviders = [StatementImportProvider.PayPalCN, StatementImportProvider.PayPalUS];

    internal static StatementImportProvider GetPayPalProvider(Account account) => account.name switch
    {
        "PAYPAL_CN" => StatementImportProvider.PayPalCN,
        "PAYPAL_US" => StatementImportProvider.PayPalUS,
        _ => throw PayPalError("unsupported PayPal account name")
    };

    internal static string ReadPayPalCursor(string key) => Uri.UnescapeDataString(key.Split('/')[2]);

    internal static string BuildPayPalKey(DateTime reportDate, string cursor, string fingerprint)
    {
        var key = $"PayPal/{reportDate:yyyy-MM-dd}/{Uri.EscapeDataString(cursor)}/{fingerprint}";
        if (key.Length > 255) throw PayPalError("statement key exceeds 255 characters");
        return key;
    }

    internal static PayPalNotice? ParsePayPalNotice(PayPalMail source)
    {
        var text = Clean(source.Text);
        var subject = Clean(source.Subject);
        // Restrict parsing to the receipt, excluding generic marketing/footer text.
        text = Regex.Split(text, @"帮助及联系我们|Help\s*(?:&|and)\s*Contact", RegexOptions.IgnoreCase)[0];
        var nativeIds = Regex.Matches(text, @"(?:交易号|交易编号|Transaction ID)\s*[:：]?\s*([A-Z0-9]{17})(?![A-Z0-9])", RegexOptions.IgnoreCase)
            .Select(m => m.Groups[1].Value.ToUpperInvariant()).Distinct().ToList();
        if (nativeIds.Count > 1) throw PayPalError("receipt contains conflicting transaction identifiers");
        var nativeId = nativeIds.SingleOrDefault() ?? "";
        var cards = Regex.Matches(text, CardPattern, RegexOptions.IgnoreCase).Select(m => m.Value).ToList();
        var tails = cards.Select(c => Regex.Match(c, @"\d{4}$").Value).Distinct().ToList();
        if (tails.Count > 1) throw PayPalError("multiple funding cards require an explicit split");
        var card = cards.FirstOrDefault() ?? "";
        PayPalAmount? cardAmount = null;
        if (card != "")
        {
            var matches = Regex.Matches(text, CardPattern + @"\s*" + MoneyPattern, RegexOptions.IgnoreCase);
            var amounts = matches.Select(ReadAmount).Distinct().ToList();
            if (amounts.Count > 1) throw PayPalError("conflicting card amounts");
            cardAmount = amounts.SingleOrDefault();
            cardAmount ??= LabelMoney(text, "您将支付的金额|Amount you paid");
        }
        var fee = LabelMoney(text, "手续费|费用|Fee(?:s)?");
        var received = LabelMoney(text, "收到的款项|Amount received");
        PayPalKind kind;
        PayPalAmount? amount;
        string party = "", bank = "";
        if (Regex.IsMatch(subject, "退款|refund", RegexOptions.IgnoreCase))
        {
            kind = PayPalKind.Refund;
            amount = LabelMoney(text, "退款总额|已退款金额|Refund total|Amount refunded");
            party = Capture(text, @"退款(?:支付方|方)\s*(?<party>.*?)\s*(?:联系|[\w.+-]+@[\w.-]+)");
        }
        else if (Regex.IsMatch(subject, "提现|withdraw|transfer.*bank", RegexOptions.IgnoreCase))
        {
            kind = PayPalKind.Withdraw;
            amount = LabelMoney(text, "总额|Total");
            party = Capture(text, @"银行账户\s+(?<party>[A-Z][A-Z .,]+?)(?:,?\s*x\s*[-*]|\s+帮助|$)");
            bank = Capture(text, @"银行账户\s+[A-Z][A-Z .,]+?,?\s*(?<party>x\s*[-*]\s*\d{3,})");
        }
        else if (received is not null)
        {
            kind = PayPalKind.Receive;
            amount = received;
            party = Capture(text, @"(?:您好[:：]?\s*)(?<party>.*?)\s*(?:向您发送了|给您发送了)");
        }
        else if (Regex.IsMatch(subject, "发送了一笔付款|sent a payment", RegexOptions.IgnoreCase))
        {
            kind = PayPalKind.Send;
            amount = LabelMoney(text, "所付款项|您已支付|Amount sent");
            party = Capture(text, @"您向(?<party>.*?)发送了");
        }
        else if (card != "" && Regex.IsMatch(subject + " " + text, "付款的收据|付款收据|您已授权|商家|Merchant|Paid .* with", RegexOptions.IgnoreCase))
        {
            kind = PayPalKind.CardPurchase;
            amount = LabelMoney(text, "总额|总计|共计|Total|净额");
            party = Capture(text, @"(?:您向|You (?:sent|paid).*? to )(?<party>.*?)\s*(?:支付了|付款|交易号|Transaction)");
            if (party == "") party = Capture(text, @"商家\s*(?<party>.*?)(?:[\w.+-]+@[\w.-]+|[+\d]|查看)");
            if (party == "") party = Capture(text, @"收款方\s*(?<party>.*?)(?:[\w.+-]+@[\w.-]+|给收款方的留言)");
        }
        else
        {
            // Pending notices are retained as source evidence; they do not prove a completed payment.
            if (Regex.IsMatch(subject, "Money is waiting|款项待领取|电子支票付款正在处理中", RegexOptions.IgnoreCase)) return null;
            if (nativeId != "" || Regex.IsMatch(subject, "退款|提现|发送.*付款|收到了|付款.*收据|donation|sent.*payment|sent you money|电子支票付款结清", RegexOptions.IgnoreCase))
                throw PayPalError("unsupported financial receipt format");
            return null;
        }
        if (amount is null) throw PayPalError("recognized receipt has no unambiguous amount");
        if (fee is not null && fee.Currency != amount.Currency) throw PayPalError("receipt fee currency differs from principal");
        var feeValue = fee?.Value ?? 0;
        if (kind == PayPalKind.CardPurchase && (LabelMoney(text, @"PayPal余额(?:（[A-Z]+）|\([A-Z]+\))?|PayPal balance")?.Value ?? 0) != 0)
            throw PayPalError("mixed card and PayPal balance funding is unsupported");
        var net = kind == PayPalKind.Receive ? LabelMoney(text, "共计|Net amount") : null;
        if (net is not null && (net.Currency != amount.Currency || amount.Value - feeValue != net.Value))
            throw PayPalError("receipt gross minus explicit fee differs from net");
        var key = $"{source.AccountId}/{kind}/" + PayPalHash(nativeId == "" ? source.MessageId : nativeId);
        return new(source, key, nativeId, kind, amount, feeValue, net?.Value ?? amount.Value - feeValue,
            Clean(party), card, cardAmount, bank,
            Regex.IsMatch(subject + text, "Donation Points.*Nexus Mods|Nexus Mods.*Donation Points", RegexOptions.IgnoreCase),
            Regex.IsMatch(text, "PayPal汇率|PayPal.s conversion rate", RegexOptions.IgnoreCase));
    }

    private static string Capture(string text, string pattern) => Regex.Match(text, pattern, RegexOptions.IgnoreCase).Groups["party"].Value.Trim();
    private static PayPalAmount? LabelMoney(string text, string labels)
    {
        var values = Regex.Matches(text, @"(?:" + labels + @")\s*[:：]?\s*" + MoneyPattern, RegexOptions.IgnoreCase)
            .Select(ReadAmount).Distinct().ToList();
        if (values.Count > 1) throw PayPalError("receipt contains conflicting labeled amounts");
        return values.SingleOrDefault();
    }
    private static PayPalAmount ReadAmount(Match match)
    {
        var value = Decimal.Parse(match.Groups["value"].Value, NumberStyles.Number, CultureInfo.InvariantCulture);
        var currency = match.Groups["currency"].Value.ToUpperInvariant();
        return new(value, Enum.Parse<CurrencyType>(currency == "CNY" ? "RMB" : currency));
    }

    public async Task FetchPayPalAsync(IReadOnlyDictionary<StatementImportProvider, DateTime> since,
        CancellationToken cancellationToken = default)
    {
        if (!await payPalLock.WaitAsync(0, cancellationToken).ConfigureAwait(false)) return;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromMinutes(5));
        try
        {
            var items = database.GetPlaidItems(PlaidUtil.UseProductionEnvironment ? PlaidEnvironment.Production : PlaidEnvironment.Sandbox)
                .Where(i => i.institutionId == PlaidUtil.PayPalInstitutionId).ToList();
            if (items.Count == 0) return;
            var accounts = items.Select(i => PlaidUtil.GetLinkedAccount(i, "PAYPAL")).ToList();
            if (accounts.Select(a => a.Id).Distinct().Count() != accounts.Count || accounts.Any(a => !a.relativeBalance || a.isCredit || a._primaryAccount_Id.HasValue))
                throw PayPalError("each Item requires a distinct primary relative-balance PayPal account");
            if (accounts.Any(a => String.IsNullOrWhiteSpace(a.email))
                || accounts.Select(a => a.email!.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).Count() != accounts.Count)
                throw PayPalError("each bound PayPal account requires its own configured mailbox");
            var checkpoints = accounts.ToDictionary(a => a.Id, a => database.GetStatementImportCheckpointTime(GetPayPalProvider(a))
                ?? throw PayPalError("fixed import checkpoint is missing"));
            var imports = new List<PayPalImport>();
            var state = new PayPalState();
            var previousImports = new Dictionary<int, StatementImport?>();
            var recordsByAccount = new Dictionary<int, List<Record>>();
            List<Record> ReadRecords(Account account) => recordsByAccount.TryGetValue(account.Id, out var records) ? records
                : recordsByAccount[account.Id] = database.GetAccountRecords(account);
            foreach (var item in items)
            {
                var account = accounts.Single(a => a.Id == item._account_Id);
                var provider = GetPayPalProvider(account);
                var previousImport = database.GetLatestStatementImport(provider);
                previousImports.Add(account.Id, previousImport);
                var cursor = previousImport is null ? null : ReadPayPalCursor(previousImport.statementKey);
                var data = await plaid.ReadPayPalAsync(item, cursor, timeout.Token).ConfigureAwait(false);
                var current = ReadPayPalSync(item.Id, account.Id, ReadRecords(account), data,
                    isInitial: previousImport is null);
                var messages = await mail.ReadPayPalMailsAsync(account, since[provider], timeout.Token).ConfigureAwait(false);
                state.Items.Add(current);
                state.Mails.AddRange(messages.DistinctBy(m => m.MessageId));
            }
            var plan = BuildPayPalPlan(state, database, checkpoints, ReadRecords);
            if (plan.Problems.Count != 0)
                throw PayPalError($"source reconciliation failed ({plan.Problems.Count}); nothing saved; " + String.Join("; ", plan.Problems.Take(8)));
            foreach (var current in state.Items)
            {
                var previous = previousImports[current.AccountId];
                var previousDate = previous is null ? checkpoints[current.AccountId]
                    : DateTime.ParseExact(previous.statementKey.Split('/')[1], "yyyy-MM-dd", CultureInfo.InvariantCulture);
                var messages = state.Mails.Where(m => m.AccountId == current.AccountId).ToList();
                var reportDate = plan.ReportDates.GetValueOrDefault(current.AccountId, previousDate);
                if (reportDate < previousDate) reportDate = previousDate;
                var unchanged = previous is not null && ReadPayPalCursor(previous.statementKey) == current.Cursor
                    && reportDate == previousDate && plan.Records.Count == 0 && plan.Pairs.Count == 0 && plan.Supplements.Count == 0;
                var key = unchanged ? previous!.statementKey : BuildPayPalKey(reportDate, current.Cursor,
                    PayPalHash((previous?.statementKey ?? "") + String.Join("\n", messages.Select(m => m.MessageId).Order(StringComparer.Ordinal))));
                imports.Add(new(GetPayPalProvider(accounts.Single(a => a.Id == current.AccountId)), current.AccountId,
                    previous?.statementKey, key));
            }
            timeout.Token.ThrowIfCancellationRequested();
            database.SavePayPalCombined(imports, plan);
        }
        catch (Exception e) when (e is not MailParseException and not PlaidUtil.PlaidRequestException)
        { throw PayPalError($"import failed: {e.GetType().Name}"); }
        finally { payPalLock.Release(); }
    }

    internal static PayPalItemState ReadPayPalSync(int itemRowId, int accountId, IReadOnlyList<Record> records,
        PlaidUtil.TransactionSyncData data, bool isInitial = false)
    {
        var accounts = (JArray?)data.Accounts["accounts"] ?? throw PayPalError("missing account snapshot");
        if (accounts.Count != 1) throw PayPalError("expected one remote PayPal account");
        var remote = accounts[0];
        var remoteId = Required(remote, "account_id");
        var currency = ParseCurrency(Required(remote["balances"]!, "iso_currency_code"));
        var rows = new Dictionary<string, JObject>(StringComparer.Ordinal);
        foreach (var page in data.Pages)
        {
            foreach (var field in new[] { "added", "modified" })
                foreach (var tx in ((JArray?)page[field] ?? throw PayPalError("missing sync collection")).Cast<JObject>())
                {
                    var id = Required(tx, "transaction_id");
                    if (Required(tx, "account_id") != remoteId || ParseCurrency(Required(tx, "iso_currency_code")) != currency
                        || tx["pending"]?.Type != JTokenType.Boolean) throw PayPalError("transaction account/currency/status mismatch");
                    Cash(tx, "amount");
                    var saved = records.Where(r => r.Source.Contains(PayPalTransactionMarker(id), StringComparison.Ordinal)).ToList();
                    if (saved.Count != 0)
                    {
                        if (saved.Any(r => !r.Source.Contains(";plaid-hash=" + TransactionFingerprint(tx) + ";", StringComparison.Ordinal)))
                            throw PayPalError("posted transaction changed; transaction=" + PayPalHash(id)[..12]);
                        continue;
                    }
                    rows[id] = tx;
                }
            foreach (var removed in ((JArray?)page["removed"] ?? throw PayPalError("missing removals")))
            {
                var id = Required(removed, "transaction_id");
                if (records.Any(r => r.Source.Contains(PayPalTransactionMarker(id), StringComparison.Ordinal)))
                    throw PayPalError("posted transaction removed; transaction=" + PayPalHash(id)[..12]);
                rows.Remove(id);
            }
        }
        if (data.Pages.Count == 0) throw PayPalError("empty sync pages");
        return new(itemRowId, accountId, remoteId, currency, Cash(remote["balances"]!, "current"),
            Required(data.Pages[^1], "next_cursor"), rows.OrderBy(p => p.Key, StringComparer.Ordinal)
                .Select(p => p.Value).Where(t => !t["pending"]!.Value<bool>()).ToList(), isInitial);
    }
    private static string PayPalTransactionMarker(string id) => ";plaid=" + Uri.EscapeDataString(id) + ";";
    private static string PayPalTransactionEvidence(PayPalTransaction tx) => PayPalTransactionMarker(tx.Id) + "plaid-hash=" + tx.Fingerprint + ";";
    private static string Required(JToken row, string name) => row[name]?.Type == JTokenType.String && !String.IsNullOrWhiteSpace((string?)row[name])
        ? (string)row[name]! : throw PayPalError("missing source field " + name);
    private static decimal Cash(JToken row, string name)
    {
        if (row[name]?.Type is not (JTokenType.Float or JTokenType.Integer)) throw PayPalError("missing monetary field " + name);
        var amount = row[name]!.Value<decimal>();
        if (Decimal.Round(amount, 2) != amount) throw PayPalError("unsupported sub-cent cash amount");
        return amount;
    }
    private static CurrencyType ParseCurrency(string currency) => Enum.TryParse<CurrencyType>(currency == "CNY" ? "RMB" : currency, out var result)
        && Enum.IsDefined(result) ? result : throw PayPalError("unsupported currency");
    private static string TransactionFingerprint(JObject row) => PayPalHash(new JObject
    {
        ["account"] = row["account_id"], ["amount"] = row["amount"], ["currency"] = row["iso_currency_code"],
        ["date"] = row["date"], ["datetime"] = row["datetime"], ["pending"] = row["pending"], ["description"] = row["original_description"] ?? row["name"]
    }.ToString(Formatting.None));
}
