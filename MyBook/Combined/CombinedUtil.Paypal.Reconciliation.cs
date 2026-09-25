using System.Globalization;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook;

partial class CombinedUtil
{
    // Readers provide evidence only. This method builds the entire batch before any financial write.
    internal static PayPalPlan BuildPayPalPlan(PayPalState state, DatabaseUtil database, DateTime checkpoint)
    {
        var accounts = database.GetAllAccounts();
        return ReconcilePayPal(state, checkpoint, accounts,
            text => database.FindAccountByInternalCardNoText(null, "PayPal counterparty", false, text) is { } account
                ? database.GetPostingAccount(account) : null,
            database.GetAccountRecords,
            (account, currency) => database.GetAccountBalance(account, currency).v);
    }

    private sealed record PayPalTransaction(int AccountId, string Id, decimal Delta, CurrencyType Currency,
        DateTime Date, DateTimeOffset? Time, string Description);

    internal static PayPalPlan ReconcilePayPal(PayPalState state, DateTime checkpoint, IReadOnlyList<Account> accounts,
        Func<string, Account?> resolveAccount, Func<Account, List<Record>> readRecords,
        Func<Account, CurrencyType, decimal> readBalance)
    {
        var plan = new PayPalPlan();
        foreach (var item in state.Items) plan.ExpectedItemAccounts.Add(item.ItemRowId, item.AccountId);
        var linked = state.Items.Select(i => i.AccountId).ToHashSet();
        var notices = new Dictionary<string, PayPalNotice>(StringComparer.Ordinal);
        foreach (var mail in state.Mails.Where(m => m.Date.LocalDateTime.Date > checkpoint.Date))
        {
            try
            {
                if (!linked.Contains(mail.AccountId)) throw PayPalError("mail has no bound Item");
                var notice = ParsePayPalNotice(mail);
                if (notice is null) continue;
                if (notices.TryGetValue(notice.Key, out var old) && NoticeFingerprint(old) != NoticeFingerprint(notice))
                    throw PayPalError("conflicting copies of a receipt");
                // Prefer the earliest delivery; duplicate copies do not change posting dates.
                if (old is null || old.Mail.Date > notice.Mail.Date) notices[notice.Key] = notice;
            }
            catch (MailParseException e) { plan.Problems.Add(e.Message + "; message=" + PayPalHash(mail.MessageId)[..12]); }
        }
        var transactions = state.Items.SelectMany(item => item.Transactions.Where(t => !t["pending"]!.Value<bool>())
            .Select(t => new PayPalTransaction(item.AccountId, Required(t, "transaction_id"), -Cash(t, "amount"),
                ParseCurrency(Required(t, "iso_currency_code")), t["date"]!.Value<DateTime>().Date,
                t["datetime"]?.Type is JTokenType.Null or null ? null : (DateTimeOffset)t["datetime"]!,
                (string?)t["original_description"] ?? (string?)t["name"] ?? "")))
            .Where(t => t.Date > checkpoint.Date).ToList();
        if (transactions.Select(t => t.Id).Distinct(StringComparer.Ordinal).Count() != transactions.Count)
            throw PayPalError("duplicate transaction identifiers across Items");
        var candidates = notices.Values.ToDictionary(n => n.Key,
            n => transactions.Where(t => Matches(n, t)).ToList(), StringComparer.Ordinal);
        foreach (var n in notices.Values)
        {
            if (candidates[n.Key].Count <= 1) continue;
            // Repeated purchases of the same amount require a unique timestamp, never list order.
            var timed = candidates[n.Key].Where(t => t.Time.HasValue
                && n.Mail.Date - t.Time.Value >= TimeSpan.Zero && n.Mail.Date - t.Time.Value <= TimeSpan.FromMinutes(5)).ToList();
            if (timed.Count == 1) candidates[n.Key] = timed;
        }
        var used = new HashSet<string>(StringComparer.Ordinal);
        var recordsByAccount = new Dictionary<int, List<Record>>();
        List<Record> Existing(Account a) => recordsByAccount.TryGetValue(a.Id, out var rows) ? rows
            : recordsByAccount[a.Id] = readRecords(a);
        var reservedBankRecords = new HashSet<int>();

        foreach (var notice in notices.Values.OrderBy(n => n.Mail.Date).ThenBy(n => n.Key, StringComparer.Ordinal))
        {
            try
            {
                var matches = candidates[notice.Key];
                if (matches.Count > 1 || (matches.Count == 1 && candidates.Count(p => p.Value.Any(t => t.Id == matches[0].Id)) != 1))
                    throw PayPalError("ambiguous mail/Plaid match");
                var tx = matches.SingleOrDefault();
                var account = accounts.Single(a => a.Id == notice.Mail.AccountId);
                var currency = notice.Gross.Currency;
                var amount = notice.Gross.Value;
                if (amount <= 0 || notice.Fee < 0 || notice.Net < 0) throw PayPalError("invalid receipt amounts");
                if (notice.ForeignExchange && notice.Kind != PayPalKind.CardPurchase)
                    throw PayPalError("transfer currency conversion requires explicit reconciliation");
                if (notice.Kind != PayPalKind.Receive && notice.Fee != 0)
                    throw PayPalError("unsupported transfer fee layout");
                var earliest = transactions.Where(t => t.AccountId == account.Id).Select(t => (DateTime?)t.Date).Min();
                if (tx is null && notice.Kind != PayPalKind.CardPurchase
                    && (notice.Kind != PayPalKind.Receive || !earliest.HasValue || notice.Mail.Date.LocalDateTime.Date >= earliest.Value.AddDays(-2)))
                    throw PayPalError("receipt awaits a posted Plaid transaction");

                var eventRecords = new List<Record>();
                Record Add(Account owner, decimal value, string role, string reason, Account? other = null, string? party = null)
                {
                    if (!owner.relativeBalance) throw PayPalError("counterpart requires its own statement import");
                    var record = new Record { Account = owner, _account_Id = owner.Id, v = value, t = currency,
                        date = tx?.Time?.LocalDateTime ?? notice.Mail.Date.LocalDateTime,
                        postingDate = tx?.Date ?? notice.Mail.Date.LocalDateTime.Date,
                        Source = PayPalPrefix + notice.Key + "/" + role, Reason = reason,
                        DestAccount = other?.name ?? party ?? notice.Party, isInternal = other is not null };
                    eventRecords.Add(record);
                    return record;
                }
                Account CardAccount()
                {
                    if (notice.Card == "" || notice.CardAmount is null || notice.CardAmount.Currency != currency || notice.CardAmount.Value != amount)
                        throw PayPalError("card transfer requires an explicit full funding/refund amount");
                    var card = resolveAccount(notice.Card) ?? throw PayPalError("funding/refund card is not registered");
                    if (linked.Contains(card.Id)) throw PayPalError("funding card resolves to a PayPal account");
                    return card;
                }
                void PairBank(Record transfer, Record bankRecord)
                {
                    if (!reservedBankRecords.Add(bankRecord.Id)) throw PayPalError("bank record claimed by multiple transfers");
                    if (bankRecord.backup is not null) throw PayPalError("bank counterpart was manually edited");
                    if (bankRecord.matchedRecordId.HasValue && !Existing(account).Any(r => r.Id == bankRecord.matchedRecordId
                        && r.Source == transfer.Source)) throw PayPalError("bank record already paired to another transfer");
                    var code = PayPalHash(transfer.Source)[..24];
                    plan.ExpectedBankRecords[bankRecord.Id] = PayPalBankRecordFingerprint(bankRecord);
                    plan.Supplements.Add(new(bankRecord.Id, code, "PayPal combined transfer; code=" + code,
                        new(DestAccount: account.name, Reason: "转账", IsInternal: true)));
                    plan.Pairs.Add(new(transfer.Source, BankRecordId: bankRecord.Id));
                }
                void CardLeg(decimal value, string role)
                {
                    var card = CardAccount();
                    var transfer = Add(account, value, role, "转账", card);
                    var bankMatches = Existing(card).Where(r => r.v == -value && r.t == currency
                        && Math.Abs((r.date - notice.Mail.Date.LocalDateTime).TotalDays) <= 7
                        && (r.DestAccount.Contains("PAYPAL", StringComparison.OrdinalIgnoreCase)
                            || r.Source.Contains("PAYPAL", StringComparison.OrdinalIgnoreCase))).ToList();
                    if (bankMatches.Count > 1) throw PayPalError("ambiguous card transfer counterpart");
                    if (bankMatches.Count == 1) PairBank(transfer, bankMatches[0]);
                }
                switch (notice.Kind)
                {
                    case PayPalKind.CardPurchase:
                        if (notice.Card == "" || notice.CardAmount is null
                            || (!notice.ForeignExchange && (notice.CardAmount.Currency != currency || notice.CardAmount.Value != amount)))
                            throw PayPalError("purchase funding is not fully explained by the card");
                        plan.IgnoredCards++;
                        break;
                    case PayPalKind.Receive:
                    {
                        var sends = notices.Values.Where(n => n.Kind == PayPalKind.Send && n.Mail.AccountId != account.Id
                            && n.Gross == notice.Gross && Math.Abs((n.Mail.Date - notice.Mail.Date).TotalDays) <= 2
                            && PartyMatches(n.Party, Greeting(notice.Mail.Text)) && PartyMatches(notice.Party, Greeting(n.Mail.Text))).ToList();
                        if (sends.Count > 1) throw PayPalError("ambiguous transfer between PayPal accounts");
                        var origin = sends.Count == 1 ? accounts.Single(a => a.Id == sends[0].Mail.AccountId) : resolveAccount(notice.Party);
                        if (notice.Nexus && (origin is null || !origin.name.StartsWith("NEXUS", StringComparison.OrdinalIgnoreCase)))
                            throw PayPalError("Donation Points payout requires a registered Nexus source account");
                        if (origin is not null && origin.Id == account.Id) throw PayPalError("receipt source resolves to itself");
                        Add(account, amount, "principal", origin is null ? "收款" : "转账", origin);
                        if (notice.Fee != 0) Add(account, -notice.Fee, "fee", "收款手续费", party: "PayPal");
                        if (origin is not null && !linked.Contains(origin.Id))
                        {
                            // Only the explicit Donation Points payout receipt supports a Nexus debit.
                            if (!notice.Nexus || !origin.name.StartsWith("NEXUS", StringComparison.OrdinalIgnoreCase))
                                throw PayPalError("own-account receipt needs source-side evidence");
                            if (Existing(origin).Any(r => r.v == -amount && r.t == currency
                                && Math.Abs((r.date - notice.Mail.Date.LocalDateTime).TotalDays) <= 2
                                && r.DestAccount == account.name && r.Source != PayPalPrefix + notice.Key + "/payout"))
                                throw PayPalError("payout already has another source record");
                            var payout = Add(origin, -amount, "payout", "转账", account);
                            plan.Pairs.Add(new(PayPalPrefix + notice.Key + "/principal", payout.Source));
                        }
                        break;
                    }
                    case PayPalKind.Send:
                    {
                        var receives = notices.Values.Where(n => n.Kind == PayPalKind.Receive && n.Mail.AccountId != account.Id
                            && n.Gross == notice.Gross && Math.Abs((n.Mail.Date - notice.Mail.Date).TotalDays) <= 2
                            && PartyMatches(notice.Party, Greeting(n.Mail.Text)) && PartyMatches(n.Party, Greeting(notice.Mail.Text))).ToList();
                        if (receives.Count != 1) throw PayPalError("sent payment requires a unique recipient receipt");
                        var target = accounts.Single(a => a.Id == receives[0].Mail.AccountId);
                        if (notice.Card != "") CardLeg(amount, "card-funding");
                        var outgoing = Add(account, -amount, "principal", "转账", target);
                        plan.Pairs.Add(new(outgoing.Source, PayPalPrefix + receives[0].Key + "/principal"));
                        break;
                    }
                    case PayPalKind.Refund:
                    {
                        var reverse = transactions.Where(t => t.AccountId != account.Id && t.Currency == currency && t.Delta == -amount
                            && Math.Abs((t.Date - notice.Mail.Date.LocalDateTime.Date).TotalDays) <= 2
                            && Regex.IsMatch(t.Description, "^Refund to\\b", RegexOptions.IgnoreCase)
                            && (PartyMatches(Greeting(notice.Mail.Text), t.Description) || PartyMatches(t.Description[10..], Greeting(notice.Mail.Text))
                                || notices.Values.Any(n => n.Kind == PayPalKind.Send && n.Mail.AccountId == account.Id
                                    && PartyMatches(Greeting(n.Mail.Text), t.Description[10..])))
                            && notices.Values.Any(n => n.Mail.AccountId == t.AccountId && n.Kind == PayPalKind.Receive
                                && PartyMatches(notice.Party, Greeting(n.Mail.Text)))).ToList();
                        if (reverse.Count > 1) throw PayPalError("ambiguous refund sender");
                        var origin = reverse.Count == 1 ? accounts.Single(a => a.Id == reverse[0].AccountId) : null;
                        if (origin is not null)
                        {
                            if (!used.Add(reverse[0].Id)) throw PayPalError("refund transaction already claimed");
                            Add(origin, -amount, "refund-out", "转账", account);
                            plan.Pairs.Add(new(PayPalPrefix + notice.Key + "/refund-out", PayPalPrefix + notice.Key + "/refund-in"));
                        }
                        else if (resolveAccount(notice.Party) is not null)
                            throw PayPalError("own-account refund requires both transaction sides");
                        if (origin is null && notice.Card != "")
                            throw PayPalError("card refund requires original-payment evidence; no inferred income is created");
                        if (notice.Card != "")
                        {
                            // A card refund never increases PayPal cash. Retain both documented transit legs.
                            Add(account, amount, "refund-in", origin is null ? "退款" : "转账", origin);
                            CardLeg(-amount, "refund-card");
                        }
                        else Add(account, amount, "refund-in", origin is null ? "退款" : "转账", origin);
                        break;
                    }
                    case PayPalKind.Withdraw:
                    {
                        var target = notice.BankIdentifier == "" ? null : resolveAccount(notice.BankIdentifier);
                        var bankMatches = accounts.Where(a => !linked.Contains(a.Id) && a._primaryAccount_Id is null
                            && (target is null || a.Id == target.Id)).SelectMany(a => Existing(a).Where(r => r.v == amount && r.t == currency
                                && Math.Abs((r.date - notice.Mail.Date.LocalDateTime).TotalDays) <= 7
                                && (r.DestAccount.Contains("PAYPAL", StringComparison.OrdinalIgnoreCase)
                                    || r.Source.Contains("PAYPAL", StringComparison.OrdinalIgnoreCase)))
                            .Select(r => (Account: a, Record: r))).ToList();
                        if (bankMatches.Count > 1) throw PayPalError("ambiguous bank withdrawal counterpart");
                        if (target is null && bankMatches.Count == 1) target = bankMatches[0].Account;
                        if (target is null || linked.Contains(target.Id)) throw PayPalError("withdrawal bank account is not identified");
                        var withdrawal = Add(account, -amount, "withdrawal", "转账", target);
                        if (bankMatches.Count == 1)
                        {
                            PairBank(withdrawal, bankMatches[0].Record);
                        }
                        break;
                    }
                }
                if (tx is not null && !used.Add(tx.Id)) throw PayPalError("transaction already claimed");
                var fingerprint = PayPalHash(NoticeFingerprint(notice) + JsonConvert.SerializeObject(eventRecords.Select(r =>
                    new { r._account_Id, r.v, r.t, r.date, r.postingDate, r.DestAccount, r.Reason, r.isInternal, r.Source })));
                if (state.Applied.TryGetValue(notice.Key, out var prior))
                {
                    if (prior != fingerprint) throw PayPalError("previously applied receipt changed; explicit reconciliation required");
                }
                else plan.Records.AddRange(eventRecords);
                plan.Applied[notice.Key] = fingerprint;
            }
            catch (Exception e) when (e is MailParseException or InvalidOperationException)
            {
                plan.Problems.Add($"event={PayPalHash(notice.Key)[..12]}, kind={notice.Kind}: "
                    + (e is MailParseException ? e.Message : "account lookup/reconciliation failed"));
            }
        }
        foreach (var tx in transactions.Where(t => !used.Contains(t.Id)))
            plan.Problems.Add("unmatched posted transaction=" + PayPalHash(tx.Id)[..12]);
        foreach (var key in state.Applied.Keys.Where(key => !plan.Applied.ContainsKey(key)))
            plan.Problems.Add("previously applied receipt missing=" + PayPalHash(key)[..12]);
        foreach (var prior in state.BankMatches)
            if (!plan.Pairs.Any(p => p.LeftSource == prior.Key && p.BankRecordId == prior.Value))
                plan.Problems.Add("previous bank counterpart missing or changed=" + PayPalHash(prior.Key)[..12]);
        // API balance is a validation only, never a balancing Record or an opening balance.
        foreach (var item in state.Items)
        {
            var account = accounts.Single(a => a.Id == item.AccountId);
            var current = readBalance(account, item.Currency);
            plan.ExpectedBalances[(account.Id, item.Currency)] = current;
            var projected = current + plan.Records.Where(r => r._account_Id == item.AccountId && r.t == item.Currency).Sum(r => r.v);
            if (projected != item.Balance) plan.Problems.Add($"account={item.AccountId}: detailed ledger differs from API balance");
            if (plan.Records.Any(r => r._account_Id == item.AccountId && r.t != item.Currency))
                plan.Problems.Add($"account={item.AccountId}: receipt currency has no API balance snapshot");
        }
        plan.MatchedTransactions = used.Count;
        return plan;
    }

    private static string Greeting(string text) => Capture(Clean(text), @"(?<party>[\p{L} ]+)[，,]\s*您好");
    private static string NoticeFingerprint(PayPalNotice n) => PayPalHash(JsonConvert.SerializeObject(new
        { n.Kind, n.Gross, n.Fee, n.Net, n.Party, n.Card, n.CardAmount, n.BankIdentifier, n.Nexus, n.ForeignExchange }));
    internal static string PayPalBankRecordFingerprint(Record r) => PayPalHash(JsonConvert.SerializeObject(new
        { r.Id, r._account_Id, r.v, r.t, r.date, r.postingDate, r.Source, r.DestAccount, r.Reason, r.isInternal, r.matchedRecordId, r.backup }));
    private static bool Matches(PayPalNotice notice, PayPalTransaction tx)
    {
        var delta = notice.Kind is PayPalKind.Receive or PayPalKind.Refund ? notice.Net : -notice.Gross.Value;
        if (tx.AccountId != notice.Mail.AccountId || tx.Currency != notice.Gross.Currency || tx.Delta != delta
            || Math.Abs((tx.Date - notice.Mail.Date.LocalDateTime.Date).TotalDays) > 2) return false;
        if (notice.NativeId.Length > 0 && tx.Description.Contains(notice.NativeId, StringComparison.OrdinalIgnoreCase)) return true;
        var party = Regex.Replace(tx.Description, @"^(?:Payment (?:from|to)|Refund (?:from|to)|Transfer to)\s+", "", RegexOptions.IgnoreCase);
        return PartyMatches(notice.Party, party) || PartyMatches(party, notice.Party);
    }
}
