// Source: ChatGPT web financial plugin -> Plaid-linked Schwab account -> ChatGPT
// scheduled task -> Resend email containing the plugin's raw financial responses.
// This is an untrusted relay export, not a bank-issued statement or a direct Plaid API response.
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using MailKit.Search;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    private const string SchwabRawSubject = "FINANCE_RAW_V1_SCHWAB";
    private const string SchwabRawPrefix = "SchwabRaw/";
    private const StatementImportProvider SchwabRawProvider = StatementImportProvider.SchwabReportMail;
    private static readonly SemaphoreSlim schwabRawLock = new(1, 1);

    public async Task FetchSchwabReports()
    {
        if (!await schwabRawLock.WaitAsync(0).ConfigureAwait(false)) return;
        var stage = "configuration";
        try
        {
            var sender = config["schwab_mail_sender"];
            if (String.IsNullOrWhiteSpace(sender) || !MailboxAddress.TryParse(sender, out var address)
                || address.Address != sender) throw SchwabRawError("missing or invalid schwab_mail_sender");
            var checkpoint = database.GetStatementImportCheckpointTime(SchwabRawProvider)
                ?? throw SchwabRawError("missing fixed mail checkpoint");
            var imports = database.GetStatementImports(SchwabRawProvider).Where(i => i.statementKey != "").ToList();
            var since = imports.Select(i => i.time.Date.AddDays(-7)).Append(checkpoint.Date).Max();
            stage = "Gmail search and download";
            await RunWithMailSessionScope(async () =>
            {
                var messages = await SearchMessagesFromMailbox(
                    CreateGmailMailbox("Schwab raw export", config["gmail_user"] ?? "") with { Proxy = null },
                    "Schwab raw export", SearchQuery.DeliveredAfter(since).And(SearchQuery.SubjectContains(SchwabRawSubject))
                        .And(SearchQuery.FromContains(sender)),
                    summary => summary.Envelope.Subject == SchwabRawSubject && SummaryIsFrom(summary, sender),
                    message => message.Subject == SchwabRawSubject, GetMailDateTime).ConfigureAwait(false);
                stage = "mail authentication and JSON validation";
                var batches = messages.Select(m => (Mail: m, Reports: ParseSchwabRawMail(m, sender)))
                    .OrderBy(b => b.Reports.Min(r => r.GeneratedAt)).ToList();
                foreach (var batch in batches)
                {
                    stage = "financial reconciliation and atomic import";
                    ImportSchwabRawReports(batch.Reports, GetMailDateTime(batch.Mail));
                }
            }).ConfigureAwait(false);
        }
        catch (MailParseException) { throw; }
        catch (Exception e) { throw SchwabRawError($"{stage} failed ({e.GetType().Name}); private details suppressed"); }
        finally { schwabRawLock.Release(); }
    }

    internal static List<SchwabRawReport> ParseSchwabRawMail(MimeMessage message, string sender)
    {
        if (message.Subject != SchwabRawSubject || message.From.Mailboxes.Count() != 1
            || !String.Equals(message.From.Mailboxes.Single().Address, sender, StringComparison.OrdinalIgnoreCase))
            throw SchwabRawError("unexpected subject or sender");
        // Only trust Gmail's top Authentication-Results, never authentication claims inside the body.
        var auth = message.Headers.FirstOrDefault(h => h.Id == HeaderId.AuthenticationResults)?.Value ?? "";
        var domain = sender.Split('@').Last();
        if (!auth.TrimStart().StartsWith("mx.google.com;", StringComparison.OrdinalIgnoreCase)
            || !Regex.IsMatch(auth, @"\bdkim=pass\b", RegexOptions.IgnoreCase)
            || !Regex.IsMatch(auth, @"\bdmarc=pass\b[^;]*\bheader\.from=" + Regex.Escape(domain) + @"(?=\s|;|$)", RegexOptions.IgnoreCase))
            throw SchwabRawError("Gmail DKIM/DMARC authentication failed or missing");
        var parts = message.Attachments.ToList();
        if (parts.Count != 1 || parts[0] is not MimePart part || part.FileName != "schwab_raw.json"
            || part.Content.Stream.Length > 4 * 1024 * 1024)
            throw SchwabRawError("expected one bounded schwab_raw.json attachment");
        using var output = new MemoryStream();
        part.Content.DecodeTo(output);
        if (output.Length > 2 * 1024 * 1024) throw SchwabRawError("JSON attachment exceeds size limit");
        try { return ParseSchwabRawJson(new UTF8Encoding(false, true).GetString(output.ToArray())); }
        catch (MailParseException) { throw; }
        catch (Exception) { throw SchwabRawError("invalid JSON attachment"); }
    }

    internal static List<SchwabRawReport> ParseSchwabRawJson(string json)
    {
        if (Encoding.UTF8.GetByteCount(json) > 2 * 1024 * 1024) throw SchwabRawError("JSON attachment exceeds size limit");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 40 });
        var root = document.RootElement;
        CheckUniqueProperties(root);
        if (RawText(root, "schema_version") != "finance_raw_export_v1" || RawText(root, "institution") != "Charles Schwab")
            throw SchwabRawError("unsupported export schema or institution");
        var generated = RawTimestamp(root, "generated_at");
        var window = root.GetProperty("window");
        var start = RawDate(window, "start_date");
        var end = RawDate(window, "end_date");
        if (RawText(window, "basis") != "inclusive_calendar_dates" || start > end || end > generated.Date)
            throw SchwabRawError("invalid reporting window");
        var responses = root.GetProperty("responses");
        var institutions = responses.GetProperty("get_linked_accounts").GetProperty("institutions").EnumerateArray()
            .Where(i => RawText(i.GetProperty("institution"), "name") == "Charles Schwab").ToList();
        if (institutions.Count != 1) throw SchwabRawError("Schwab institution is missing or ambiguous");
        var institution = institutions[0];
        var itemId = RawText(institution, "item_id");
        if (RawText(institution, "status") != "linked" || RawText(institution, "sync_status") != "synced"
            || institution.GetProperty("connection_error_code").ValueKind != JsonValueKind.Null
            || RawText(root.GetProperty("resolved_identifiers"), "item_id") != itemId)
            throw SchwabRawError("institution is not synchronized or identity differs");
        var accounts = institution.GetProperty("accounts").EnumerateArray().ToList();
        var accountIds = accounts.Select(a => RawText(a, "account_id")).ToHashSet(StringComparer.Ordinal);
        var resolved = root.GetProperty("resolved_identifiers").GetProperty("account_ids").EnumerateArray().Select(a => a.GetString()!).ToList();
        if (accounts.Count == 0 || accounts.Count != accountIds.Count || !accountIds.SetEquals(resolved))
            throw SchwabRawError("account selection is incomplete or ambiguous");
        var requests = root.GetProperty("requests");
        foreach (var name in new[] { "investment_holdings", "investment_transactions", "posted_transactions", "pending_transactions" })
        {
            var request = requests.GetProperty(name);
            if (RawText(request, "item_id") != itemId) throw SchwabRawError("request institution mismatch");
            if (name == "investment_holdings") continue;
            var range = request.GetProperty("date");
            var dates = range.GetProperty("val").EnumerateArray().Select(d => ParseRawDate(d.GetString()!)).ToArray();
            if (RawText(range, "op") != "between" || dates.Length != 2 || dates[0] != start || dates[1] != end)
                throw SchwabRawError("request window differs from export window");
        }
        var positions = ReadRawPages(responses, "investment_holdings", "investment_holdings", itemId, accountIds);
        var transactions = ReadRawPages(responses, "investment_transactions", "investment_transactions", itemId, accountIds);
        var posted = ReadRawPages(responses, "posted_transactions", "transactions", itemId, accountIds);
        _ = ReadRawPages(responses, "pending_transactions", "transactions", itemId, accountIds);
        // Investments and bank transactions can describe the same event. Never double-book an unproven mapping.
        if (posted.Count != 0) throw SchwabRawError("nonempty posted bank transactions require an explicit investment-event mapping");
        var result = new List<SchwabRawReport>();
        foreach (var account in accounts)
        {
            var id = RawText(account, "account_id");
            if (RawText(account, "type") != "investment" || RawText(account, "subtype") != "brokerage"
                || RawText(account, "item_id") != itemId) throw SchwabRawError("unsupported account type or identity");
            var balances = account.GetProperty("balances");
            RequireRawUsd(balances);
            var report = new SchwabRawReport
            {
                AccountId = id, AccountTail = RawText(account, "mask"), ItemId = itemId,
                GeneratedAt = generated, Start = start, End = end,
                AsOf = RawTimestamp(account, "last_successful_update"),
                Total = RawNumber(balances, "current"),
                AccountCoverageComplete = institution.GetProperty("account_coverage").GetProperty("is_complete_for_query").GetBoolean()
            };
            if (!Regex.IsMatch(report.AccountTail, @"^\d{4}$") || report.AsOf > generated)
                throw SchwabRawError("invalid explicit account mask or source update time");
            RawEqual(report.Total, RawNumber(balances, "net_balance"), "account net balance");
            RawEqual(report.Total, RawNumber(balances, "display_balance"), "account display balance");
            foreach (var row in positions.Where(r => RawText(r, "account_id") == id))
            {
                RequireRawUsd(row);
                report.Positions.Add(new(RawText(row, "security_id"), RawText(row, "name"), RawText(row, "type"),
                    RawOptionalText(row, "ticker_symbol"), RawNumber(row, "quantity"), RawNumber(row, "institution_price"),
                    RawNumber(row, "institution_value"), RawDate(row, "institution_price_as_of")));
            }
            foreach (var row in transactions.Where(r => RawText(r, "account_id") == id))
            {
                RequireRawUsd(row);
                var date = RawDate(row, "date");
                var tradeDate = row.TryGetProperty("transaction_datetime", out var stamp) && stamp.ValueKind != JsonValueKind.Null
                    ? RawTimestamp(row, "transaction_datetime").Date : date;
                if (date < start || date > end || tradeDate > date) throw SchwabRawError("transaction date outside query window");
                report.Transactions.Add(new(RawText(row, "investment_transaction_id"), RawOptionalText(row, "security_id"),
                    date, tradeDate, RawText(row, "type"), RawText(row, "subtype"), RawNumber(row, "quantity"),
                    RawNumber(row, "amount"), RawNumber(row, "price"), RawNumber(row, "fees"), RawText(row, "name")));
            }
            if (report.Positions.Select(p => p.Id).Distinct().Count() != report.Positions.Count
                || report.Transactions.Select(t => t.Id).Distinct().Count() != report.Transactions.Count)
                throw SchwabRawError("duplicate security or transaction ID");
            report.Positions = report.Positions.OrderBy(p => p.Id, StringComparer.Ordinal).ToList();
            report.Transactions = report.Transactions.OrderBy(t => t.Date).ThenBy(t => t.Id, StringComparer.Ordinal).ToList();
            result.Add(report);
        }
        return result;
    }

    private static List<JsonElement> ReadRawPages(JsonElement responses, string name, string queryType, string itemId, HashSet<string> accounts)
    {
        var pages = responses.GetProperty(name + "_pages").EnumerateArray().ToList();
        if (pages.Count == 0 || pages.Count > 100) throw SchwabRawError("missing or excessive result pages");
        var rows = new List<JsonElement>();
        var cursors = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < pages.Count; index++)
        {
            var results = pages[index].GetProperty("results").EnumerateArray().ToList();
            if (results.Count != 1 || RawText(results[0], "query_type") != queryType || results[0].GetProperty("is_error").GetBoolean())
                throw SchwabRawError("failed or unexpected query result: " + name);
            var result = results[0].GetProperty("result");
            var more = result.GetProperty("has_more").GetBoolean();
            if (more != (index < pages.Count - 1) || more == (result.GetProperty("next_cursor").ValueKind == JsonValueKind.Null))
                throw SchwabRawError("incomplete pagination: " + name);
            if (more && !cursors.Add(RawText(result, "next_cursor"))) throw SchwabRawError("repeated pagination cursor: " + name);
            if (result.TryGetProperty("coverage", out var coverage) && !coverage.GetProperty("is_complete_for_query").GetBoolean())
                throw SchwabRawError("incomplete query coverage: " + name);
            foreach (var row in result.GetProperty("items").EnumerateArray())
            {
                if (RawText(row, "item_id") != itemId || !accounts.Contains(RawText(row, "account_id")))
                    throw SchwabRawError("query contains another institution or unselected account");
                rows.Add(row);
            }
        }
        return rows;
    }

    internal void ImportSchwabRawReports(List<SchwabRawReport> reports, DateTime mailTime)
    {
        if (reports.Count == 0 || reports.Select(r => r.AccountTail).Distinct().Count() != reports.Count)
            throw SchwabRawError("empty or duplicate account batch");
        var imports = database.GetStatementImports(SchwabRawProvider).Where(i => i.statementKey != "").ToList();
        if (imports.Any(i => !i.statementKey.StartsWith(SchwabRawPrefix, StringComparison.Ordinal) || i.sourceDataJson is null))
            throw SchwabRawError("legacy or missing import metadata; explicit migration required");
        var pending = new List<StatementRecordHoldingImport>();
        var resolvedAccounts = new HashSet<int>();
        foreach (var report in reports)
        {
            var account = database.FindAccountByInternalCardNo(report.AccountTail, "SCHWAB")
                ?? throw SchwabRawError("explicit account identifier is not registered");
            if (!resolvedAccounts.Add(account.Id)) throw SchwabRawError("multiple source accounts resolve to the same database account");
            var accountKeyPrefix = SchwabRawPrefix + account.Id + "/";
            var prior = imports.Where(i => i.statementKey.StartsWith(accountKeyPrefix, StringComparison.Ordinal))
                .Select(i => JsonSerializer.Deserialize<SchwabRawReport>(i.sourceDataJson!)
                    ?? throw SchwabRawError("invalid stored metadata"))
                .OrderBy(r => r.GeneratedAt).ToList();
            var json = JsonSerializer.Serialize(report);
            if (prior.Any(r => JsonSerializer.Serialize(r) == json)) continue;
            var beginning = database.GetCurrentAccountHoldings(account);
            if (prior.Count == 0 && (database.HasAccountHistory(account) || beginning.Any(h => h.totalPrice.v != 0 || h.holdingType != HoldingType.Cash && h.quantity != 0)))
                throw SchwabRawError("first raw import requires a clean account or an explicit migration");
            pending.Add(BuildSchwabRawImport(report, prior, account, beginning, mailTime, database.GetKnownEquityHoldingType));
        }
        if (pending.Count == 0) return;
        database.SaveStatementRecordsAndHoldingsOnce(pending);
        Console.WriteLine($"Schwab raw export imported accounts={pending.Count}; records={pending.Sum(p => p.Records.Count)}");
    }

    internal static StatementRecordHoldingImport BuildSchwabRawImport(SchwabRawReport report, List<SchwabRawReport> history,
        Account account, List<Holding> beginning, DateTime mailTime, Func<string, HoldingType> resolveEquity)
    {
        if (account.relativeBalance || account.isCredit || account.usage != AccountUsage.Investment || account._primaryAccount_Id.HasValue)
            throw SchwabRawError("a primary absolute-balance investment account is required");
        var previous = history.MaxBy(r => r.GeneratedAt);
        if (previous is not null && (report.GeneratedAt <= previous.GeneratedAt || report.AsOf < previous.AsOf
            || report.Start > previous.AsOf.Date || report.End < previous.End
            || report.AccountId != previous.AccountId || report.ItemId != previous.ItemId))
            throw SchwabRawError("report order, overlapping coverage or persistent account identity changed");
        var known = new Dictionary<string, SchwabRawTransaction>(StringComparer.Ordinal);
        foreach (var tx in history.SelectMany(r => r.Transactions))
        {
            if (known.TryGetValue(tx.Id, out var old) && old != tx) throw SchwabRawError("stored transaction history is inconsistent");
            known[tx.Id] = tx;
        }
        var incoming = report.Transactions.ToDictionary(t => t.Id, StringComparer.Ordinal);
        foreach (var tx in report.Transactions)
            if (known.TryGetValue(tx.Id, out var old) && old != tx) throw SchwabRawError("a previously imported transaction was revised");
        foreach (var tx in known.Values.Where(t => t.Date >= report.Start && t.Date <= report.End))
            if (!incoming.TryGetValue(tx.Id, out var current) || current != tx)
                throw SchwabRawError("a previously imported transaction was removed or revised in the overlap window");
        var securities = history.SelectMany(r => r.Positions).Concat(report.Positions).GroupBy(p => p.Id)
            .ToDictionary(g => g.Key, g => g.Last(), StringComparer.Ordinal);
        var ending = report.Positions.Select(p => BuildRawHolding(p, account, resolveEquity)).ToList();
        if (ending.Select(h => (h.code, h.holdingType)).Distinct().Count() != ending.Count
            || ending.Count(h => h.holdingType == HoldingType.Cash) != 1)
            throw SchwabRawError("ambiguous holding identity or missing USD cash holding");
        RawEqual(ending.Sum(h => h.totalPrice.v), report.Total, "holding details versus account total");
        if (report.Positions.Any(p => p.PriceDate > report.AsOf.Date)) throw SchwabRawError("holding price is newer than source update time");
        if (previous is not null)
        {
            var expected = previous.Positions.Select(p => BuildRawHolding(p, account, resolveEquity)).ToList();
            if (!HoldingState(beginning).SequenceEqual(HoldingState(expected)))
                throw SchwabRawError("database holdings differ from the last imported source");
        }
        var cash = beginning.Where(h => h.holdingType == HoldingType.Cash).Sum(h => h.totalPrice.v);
        var quantities = beginning.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => (h.code, h.holdingType), h => h.quantity);
        var values = beginning.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => (h.code, h.holdingType), h => h.totalPrice.v);
        var records = new List<Record>();
        foreach (var tx in report.Transactions)
        {
            if (known.ContainsKey(tx.Id)) continue;
            if (tx.Date > report.AsOf.Date || previous is not null && tx.Date <= previous.AsOf.Date)
                throw SchwabRawError("transaction is later than holdings or backdated before the prior valuation");
            var source = SchwabRawPrefix + "transaction/" + RawHash(tx.Id);
            var detail = "; " + tx.Name;
            if (tx.Type is "buy" or "sell")
            {
                if (!securities.TryGetValue(tx.SecurityId, out var security)) throw SchwabRawError("trade security metadata is missing");
                var holding = BuildRawHolding(security, account, resolveEquity);
                if (holding.holdingType == HoldingType.Cash || tx.Price <= 0 || tx.Fees < 0
                    || (tx.Type == "buy" ? tx.Quantity <= 0 || tx.Amount <= 0 : tx.Quantity >= 0 || tx.Amount >= 0))
                    throw SchwabRawError("invalid signed security trade");
                var principal = Decimal.Round(tx.Quantity * tx.Price / (holding.holdingType == HoldingType.UST ? 100m : 1m), 2, MidpointRounding.AwayFromZero);
                // User-authorized exception ONLY for Schwab raw-mail UST trades: the signed
                // settlement minus clean principal and explicit fees is settled accrued interest.
                // Do not generalize this to other products, providers, or balance reconciliation.
                var settledAccruedInterest = 0m;
                if (holding.holdingType == HoldingType.UST)
                {
                    settledAccruedInterest = tx.Amount - principal - tx.Fees;
                    if (Decimal.Round(settledAccruedInterest, 2) != settledAccruedInterest
                        || settledAccruedInterest != 0 && Math.Sign(settledAccruedInterest) != Math.Sign(tx.Quantity))
                        throw SchwabRawError("invalid UST settled accrued-interest sign or precision");
                }
                else
                    RawEqual(tx.Amount, principal + tx.Fees, $"trade settlement on {tx.Date:yyyy-MM-dd} ({tx.Type}); explicit interest/fee detail required");
                var key = (holding.code, holding.holdingType);
                quantities[key] = quantities.GetValueOrDefault(key) + tx.Quantity;
                values[key] = values.GetValueOrDefault(key) + principal;
                Add(principal, tx.Type == "buy" ? "买入" : "卖出", source + "/asset" + detail, tx.TradeDate, tx.Date, true, holding, tx.Quantity);
                Add(-principal, tx.Type == "buy" ? "买入" : "卖出", source + "/cash" + detail, tx.TradeDate, tx.Date, true, null, 0);
                if (tx.Fees != 0) Add(-tx.Fees, "手续费", source + "/fee" + detail, tx.TradeDate, tx.Date, false, null, 0);
                // This interest is already paid/received in cash, not an outstanding accrued holding.
                if (settledAccruedInterest != 0)
                    Add(-settledAccruedInterest, "债券利息", source + "/accrued-interest-settlement" + detail,
                        tx.TradeDate, tx.Date, false, null, 0);
            }
            else
            {
                if (tx.Quantity != 0 || tx.Price != 0 || tx.Fees != 0) throw SchwabRawError("non-trade has security movement or unsplit fees");
                var reason = (tx.Type, tx.Subtype) switch
                {
                    ("transfer", "transfer" or "contribution" or "deposit" or "withdrawal" or "distribution") => tx.Amount < 0 ? "转入" : "转出",
                    ("cash", "dividend" or "qualified dividend" or "non-qualified dividend") when tx.Amount < 0 => "股息",
                    ("cash", "interest") when tx.Amount < 0 => securities.TryGetValue(tx.SecurityId, out var p) && p.Type == "fixed income" ? "债券利息" : "现金利息",
                    ("fee", "margin expense") when tx.Amount > 0 => "现金利息",
                    ("fee", "tax" or "tax withheld" or "non-resident tax") when tx.Amount > 0 => "税费",
                    ("fee", "account fee" or "management fee" or "transfer fee" or "miscellaneous fee") when tx.Amount > 0 => "手续费",
                    _ => throw SchwabRawError("unsupported investment transaction type/subtype")
                };
                Add(-tx.Amount, reason, source + "/cash" + detail, tx.TradeDate, tx.Date, false, null, 0);
            }
            cash -= tx.Amount;
        }
        RawEqual(cash, ending.Single(h => h.holdingType == HoldingType.Cash).totalPrice.v, "cash from individual transactions");
        var endingAssets = ending.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => (h.code, h.holdingType));
        foreach (var key in quantities.Keys.Union(endingAssets.Keys))
        {
            var holding = endingAssets.GetValueOrDefault(key) ?? beginning.FirstOrDefault(h => (h.code, h.holdingType) == key)
                ?? securities.Values.Select(p => BuildRawHolding(p, account, resolveEquity)).First(h => (h.code, h.holdingType) == key);
            RawEqual(quantities.GetValueOrDefault(key), endingAssets.GetValueOrDefault(key)?.quantity ?? 0, "security quantity from individual transactions");
            var change = (endingAssets.GetValueOrDefault(key)?.totalPrice.v ?? 0) - values.GetValueOrDefault(key);
            if (change != 0)
            {
                var priceDate = report.Positions.FirstOrDefault(p => BuildRawHolding(p, account, resolveEquity).code == key.code)?.PriceDate ?? report.AsOf.Date;
                if (previous is not null && priceDate < previous.AsOf.Date) throw SchwabRawError("changed valuation has stale price date");
                Add(change, "持仓价格变动", SchwabRawPrefix + "valuation/" + RawHash(JsonSerializer.Serialize(report)) + "/" + holding.code,
                    priceDate, report.AsOf.Date, false, holding, 0);
            }
        }
        var beginningValue = beginning.Sum(h => h.totalPrice.v);
        RawEqual(beginningValue + records.Sum(r => r.v), report.Total, "opening value plus records");
        return new(SchwabRawProvider, mailTime, SchwabRawPrefix + account.Id + "/" + RawHash(JsonSerializer.Serialize(report)),
            account, records, ending, [new(account, new(report.Total, CurrencyType.USD))],
            [new(account, new(beginningValue, CurrencyType.USD))], beginning, recordDate: report.AsOf.Date,
            sourceDataJson: JsonSerializer.Serialize(report));

        void Add(decimal value, string reason, string source, DateTime date, DateTime posted, bool internalTrade, Holding? holding, decimal quantity) =>
            records.Add(new Record { Account = account, v = value, t = CurrencyType.USD, date = date, postingDate = posted,
                updateTime = DateTime.Now, Reason = reason, Source = source, isInternal = internalTrade, Holding = holding, HoldingQuantity = quantity });
    }

    private static IEnumerable<(string, HoldingType, decimal, decimal, CurrencyType)> HoldingState(IEnumerable<Holding> holdings) =>
        holdings.Where(h => h.totalPrice.v != 0 || h.holdingType != HoldingType.Cash && h.quantity != 0)
            .OrderBy(h => h.code, StringComparer.Ordinal).ThenBy(h => h.holdingType)
            .Select(h => (h.code, h.holdingType, h.quantity, h.totalPrice.v, h.currentPrice.t));

    private static Holding BuildRawHolding(SchwabRawPosition position, Account account, Func<string, HoldingType> resolveEquity)
    {
        Holding holding;
        if (position.Type == "cash" && position.Symbol == "CUR:USD")
        {
            RawEqual(position.Price, 1, "cash unit price");
            RawEqual(position.Quantity, position.Value, "cash quantity");
            holding = new("USD", HoldingType.Cash) { currentPrice = new(position.Value, CurrencyType.USD) };
        }
        else
        {
            var price = position.Price;
            var code = position.Symbol;
            HoldingType type;
            if (position.Type == "fixed income")
            {
                var bond = Regex.Match(position.Name, @"^US Treasury Bond - (?<coupon>\d+(?:\.\d+)?)% (?<date>\d{2}/\d{2}/\d{4}) USD 100$");
                if (!bond.Success) throw SchwabRawError("unsupported fixed-income product metadata");
                var maturity = DateTime.ParseExact(bond.Groups["date"].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                code = $"T {Decimal.Parse(bond.Groups["coupon"].Value, CultureInfo.InvariantCulture).ToString("G29", CultureInfo.InvariantCulture)} {maturity:MM/dd/yy}";
                type = HoldingType.UST;
                price /= 100;
            }
            else if (position.Type is "equity" or "etf" && code.Length > 0)
            {
                type = resolveEquity(code);
                if (type is not (HoldingType.NASDAQ or HoldingType.ARCA)) throw SchwabRawError("invalid equity exchange metadata");
            }
            else throw SchwabRawError("unsupported financial product");
            if (position.Quantity < 0 || price < 0) throw SchwabRawError("short or negative-priced position");
            holding = new(code, type) { quantity = position.Quantity, currentPrice = new(price, CurrencyType.USD) };
        }
        holding.Account = account;
        holding.desc = position.Name;
        holding.displayText = position.Name;
        RawEqual(holding.totalPrice.v, position.Value, "stored quantity times unit price");
        return holding;
    }

    private static void CheckUniqueProperties(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in element.EnumerateObject())
            {
                if (!names.Add(property.Name)) throw SchwabRawError("duplicate JSON property");
                CheckUniqueProperties(property.Value);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array) foreach (var item in element.EnumerateArray()) CheckUniqueProperties(item);
    }
    private static string RawText(JsonElement row, string field) => row.TryGetProperty(field, out var value)
        && value.ValueKind == JsonValueKind.String && !String.IsNullOrWhiteSpace(value.GetString()) && value.GetString()!.Length <= 512
        ? value.GetString()! : throw SchwabRawError("missing or invalid text field: " + field);
    private static string RawOptionalText(JsonElement row, string field) => !row.TryGetProperty(field, out var value) || value.ValueKind == JsonValueKind.Null ? "" : RawText(row, field);
    private static decimal RawNumber(JsonElement row, string field)
    {
        if (!row.TryGetProperty(field, out var value) || value.ValueKind != JsonValueKind.Number || !value.TryGetDecimal(out var number))
            throw SchwabRawError("missing or invalid decimal field: " + field);
        MySqlDecimalColumnTypes.ValidateCurrencyValue(number, "Schwab raw " + field);
        return number;
    }
    private static DateTime RawDate(JsonElement row, string field) => ParseRawDate(RawText(row, field));
    private static DateTime ParseRawDate(string value) => DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
        ? date : throw SchwabRawError("invalid calendar date");
    private static DateTimeOffset RawTimestamp(JsonElement row, string field) => DateTimeOffset.TryParse(RawText(row, field), CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
        ? date : throw SchwabRawError("invalid source timestamp");
    private static void RequireRawUsd(JsonElement row)
    {
        if (RawText(row, "iso_currency_code") != "USD") throw SchwabRawError("unsupported currency; explicit FX details required");
    }
    private static void RawEqual(decimal left, decimal right, string field)
    {
        if (left != right) throw SchwabRawError(field + " mismatch");
    }
    private static string RawHash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    private static MailParseException SchwabRawError(string message) => new("Schwab raw export: " + message);

    internal sealed class SchwabRawReport
    {
        public string AccountId { get; set; } = "";
        public string AccountTail { get; set; } = "";
        public string ItemId { get; set; } = "";
        public DateTimeOffset GeneratedAt { get; set; }
        public DateTimeOffset AsOf { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public decimal Total { get; set; }
        public bool AccountCoverageComplete { get; set; }
        public List<SchwabRawPosition> Positions { get; set; } = [];
        public List<SchwabRawTransaction> Transactions { get; set; } = [];
    }
    internal sealed record SchwabRawPosition(string Id, string Name, string Type, string Symbol, decimal Quantity, decimal Price, decimal Value, DateTime PriceDate);
    internal sealed record SchwabRawTransaction(string Id, string SecurityId, DateTime Date, DateTime TradeDate, string Type, string Subtype,
        decimal Quantity, decimal Amount, decimal Price, decimal Fees, string Name = "");
}
