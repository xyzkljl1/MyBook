using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace MyBook;

partial class PlaidUtil
{
    // Direct read-only Plaid Investments API, without email or Drive relay.
    private const string SchwabInstitutionId = "ins_11";
    private const string SchwabRawPrefix = "PlaidSchwab/";
    private const StatementImportProvider SchwabRawProvider = StatementImportProvider.PlaidSchwab;
    private static readonly SemaphoreSlim schwabLock = new(1, 1);
    public Task<PlaidItem?> FindSchwabItemAsync(CancellationToken cancellationToken = default) =>
        FindItemByInstitutionAsync(SchwabInstitutionId, cancellationToken);

    public async Task FetchSchwabAsync(CancellationToken cancellationToken = default)
    {
        if (!await schwabLock.WaitAsync(0, cancellationToken).ConfigureAwait(false)) return;
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(TimeSpan.FromMinutes(3));
        var stage = "Item selection";
        try
        {
            var item = await FindSchwabItemAsync(deadline.Token).ConfigureAwait(false)
                ?? throw SchwabRawError("no Schwab Item linked in the active environment");
            var db = database ?? throw SchwabRawError("database unavailable");
            var checkpoint = db.GetStatementImportCheckpointTime(SchwabRawProvider)
                ?? throw SchwabRawError("missing fixed import checkpoint");
            var history = db.GetStatementImports(SchwabRawProvider).Where(i => i.statementKey != "").OrderBy(i => i.Id)
                .Select(i => JsonSerializer.Deserialize<SchwabRawReport>(i.sourceDataJson
                    ?? throw SchwabRawError("stored source metadata missing")) ?? throw SchwabRawError("invalid stored metadata")).ToList();
            var since = history.Select(r => r.End.AddDays(-7)).Append(checkpoint.Date).Max();
            var queryEnd = DateTime.UtcNow.Date;
            stage = "investment retrieval";
            var data = await GetInvestmentsAsync(item, since, queryEnd, deadline.Token).ConfigureAwait(false);
            stage = "financial reconciliation";
            var reports = ParseSchwabInvestments(data, item.itemId, since, queryEnd);
            var imports = new List<StatementRecordHoldingImport>();
            var accounts = new HashSet<int>();
            foreach (var report in reports)
            {
                var account = db.FindAccountByInternalCardNo(report.AccountTail, "SCHWAB")
                    ?? throw SchwabRawError("explicit account identifier is not registered");
                if (!accounts.Add(account.Id)) throw SchwabRawError("multiple source accounts resolve to one local account");
                var prior = history.Where(r => r.AccountId == report.AccountId).ToList();
                var beginning = db.GetCurrentAccountHoldings(account);
                if (prior.Count == 0 && (db.HasAccountHistory(account) || beginning.Any(h => h.totalPrice.v != 0 || h.quantity != 0)))
                    throw SchwabRawError("existing account data requires explicit migration");
                HoldingType Resolve(string symbol)
                {
                    var codes = prior.Append(report).SelectMany(r => r.Markets).Where(m => m.Key == symbol).Select(m => m.Value).Distinct().ToList();
                    return codes.Count == 1 ? codes[0] switch
                    {
                        "XNAS" => HoldingType.NASDAQ, "ARCX" => HoldingType.ARCA,
                        _ => throw SchwabRawError("unsupported equity market")
                    } : throw SchwabRawError("missing or inconsistent equity market");
                }
                var import = BuildSchwabRawImport(report, prior, account, beginning, report.AsOf.Date, Resolve);
                // A successful empty day must also advance the query boundary.
                if (prior.Count == 0 || import.Records.Count != 0 || report.End > prior[^1].End) imports.Add(import);
            }
            if (db.GetAccountsByNamePrefix("SCHWAB_").Any(account => !accounts.Contains(account.Id)))
                throw SchwabRawError("response is missing a configured Schwab account; no accounts were saved");
            deadline.Token.ThrowIfCancellationRequested();
            stage = "atomic database write";
            db.SaveStatementRecordsAndHoldingsOnce(imports);
            Console.WriteLine($"Plaid Schwab: accounts={reports.Count}, saved={imports.Count}, records={imports.Sum(i => i.Records.Count)}");
        }
        catch (MailParseException) { throw; }
        catch (PlaidRequestException e) { throw SchwabRawError(stage + ": " + e.Message); }
        catch (TimeoutException e) { throw SchwabRawError(stage + ": " + e.Message); }
        catch (OperationCanceledException) { throw SchwabRawError(stage + ": cancelled or overall timeout"); }
        catch (Exception e) { throw SchwabRawError(stage + " failed (" + e.GetType().Name + "); private details suppressed"); }
        finally { schwabLock.Release(); }
    }

    internal static List<SchwabRawReport> ParseSchwabInvestments(InvestmentData data, string itemId, DateTime start, DateTime end)
    {
        CheckFields(data.Holdings, "accounts holdings item request_id securities");
        var securities = new Dictionary<string, JObject>(StringComparer.Ordinal);
        foreach (var group in data.Securities.GroupBy(s => Text(s, "security_id"), StringComparer.Ordinal))
        {
            var selected = group.First();
            foreach (var row in group)
            {
                CheckFields(row, "cfi_code close_price close_price_as_of cusip figi fixed_income industry institution_id institution_security_id is_cash_equivalent isin iso_currency_code market_identifier_code name option_contract proxy_security_id sector security_id sedol subtype ticker_symbol type unofficial_currency_code update_datetime");
                RequireUsd(row);
                if (row["option_contract"]?.Type is not (null or JTokenType.Null)) throw SchwabRawError("unsupported option contract");
                foreach (var field in new[] { "type", "ticker_symbol", "market_identifier_code", "fixed_income" })
                    if (!JToken.DeepEquals(selected[field], row[field])) throw SchwabRawError("inconsistent security metadata");
            }
            securities.Add(group.Key, selected);
        }
        var accountRows = Rows(data.Holdings, "accounts").ToList();
        var accountIds = accountRows.Select(a => Text(a, "account_id")).ToHashSet(StringComparer.Ordinal);
        if (accountIds.Count != accountRows.Count || accountIds.Count == 0) throw SchwabRawError("empty or duplicate accounts");
        var positions = Rows(data.Holdings, "holdings").ToList();
        foreach (var row in positions.Concat(data.Transactions))
            if (!accountIds.Contains(Text(row, "account_id"))) throw SchwabRawError("unselected account in investment data");
        var result = new List<SchwabRawReport>();
        foreach (var account in accountRows)
        {
            CheckFields(account, "account_id apy balances mask name official_name subtype type persistent_account_id");
            if (Text(account, "type") != "investment" || Text(account, "subtype") != "brokerage") throw SchwabRawError("unsupported account type");
            var balances = account["balances"] as JObject ?? throw SchwabRawError("missing balances");
            CheckFields(balances, "available current iso_currency_code limit margin_loan_amount unofficial_currency_code last_updated_datetime");
            RequireUsd(balances);
            var report = new SchwabRawReport
            {
                AccountId = Text(account, "account_id"), AccountTail = Text(account, "mask"), ItemId = itemId,
                Start = start, End = end, Total = Number(balances, "current")
            };
            foreach (var row in positions.Where(p => Text(p, "account_id") == report.AccountId))
            {
                CheckFields(row, "account_id cost_basis institution_price institution_price_as_of institution_price_datetime institution_value iso_currency_code quantity security_id tax_lots unofficial_currency_code vested_quantity vested_value");
                RequireUsd(row);
                report.Positions.Add(Position(Text(row, "security_id"), Number(row, "quantity"), Number(row, "institution_price"),
                    Number(row, "institution_value"), Date(row, "institution_price_as_of")));
            }
            if (report.Positions.Count == 0) throw SchwabRawError("missing holdings or cash detail");
            // Observation/query date, not a bank update timestamp. Quotes retain their own dates.
            report.AsOf = new DateTimeOffset(DateTime.SpecifyKind(end.Date, DateTimeKind.Utc));
            report.GeneratedAt = report.AsOf;
            if (report.Positions.Any(p => p.PriceDate > end)) throw SchwabRawError("future holdings date");
            foreach (var row in data.Transactions.Where(t => Text(t, "account_id") == report.AccountId))
            {
                CheckFields(row, "account_id amount cancel_transaction_id date fees investment_transaction_id iso_currency_code name price quantity security_id subtype transaction_datetime type unofficial_currency_code");
                RequireUsd(row);
                if (OptionalText(row, "cancel_transaction_id") != "") throw SchwabRawError("cancelled transaction requires explicit reconciliation");
                var date = Date(row, "date");
                var stamp = OptionalText(row, "transaction_datetime");
                var tradeDate = stamp == "" ? date : DateTimeOffset.TryParse(stamp, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
                    ? parsed.Date : throw SchwabRawError("invalid transaction timestamp");
                if (date < start || date > end || tradeDate > date) throw SchwabRawError("invalid transaction date or coverage");
                report.Transactions.Add(new(Text(row, "investment_transaction_id"), OptionalText(row, "security_id"), date, tradeDate,
                    Text(row, "type"), Text(row, "subtype"), Number(row, "quantity"), Number(row, "amount"), Number(row, "price"), Number(row, "fees"), Text(row, "name")));
            }
            foreach (var id in report.Positions.Select(p => p.Id).Concat(report.Transactions.Select(t => t.SecurityId).Where(s => s != "")).Distinct())
            {
                if (!securities.TryGetValue(id, out var security)) throw SchwabRawError("missing security metadata");
                report.Securities.Add(Position(id, 0, Text(security, "type") == "cash" ? 1 : 0, 0, report.AsOf.Date));
                if (Text(security, "type") is "equity" or "etf")
                {
                    var symbol = Text(security, "ticker_symbol");
                    var market = Text(security, "market_identifier_code");
                    if (report.Markets.TryGetValue(symbol, out var old) && old != market) throw SchwabRawError("ambiguous equity market");
                    report.Markets[symbol] = market;
                }
            }
            if (report.Positions.Select(p => p.Id).Distinct().Count() != report.Positions.Count
                || report.Transactions.Select(t => t.Id).Distinct().Count() != report.Transactions.Count) throw SchwabRawError("duplicate holdings or transactions");
            report.Positions = report.Positions.OrderBy(p => p.Id, StringComparer.Ordinal).ToList();
            report.Transactions = report.Transactions.OrderBy(t => t.Date).ThenBy(t => t.Id, StringComparer.Ordinal).ToList();
            result.Add(report);
        }
        return result;

        SchwabRawPosition Position(string id, decimal quantity, decimal price, decimal value, DateTime priceDate)
        {
            if (!securities.TryGetValue(id, out var row)) throw SchwabRawError("missing security metadata");
            var name = Text(row, "name");
            var type = Text(row, "type");
            if (type == "fixed income")
            {
                var fixedIncome = row["fixed_income"] as JObject ?? throw SchwabRawError("missing fixed-income metadata");
                CheckFields(fixedIncome, "face_value issue_date maturity_date yield_rate");
                var yield = fixedIncome["yield_rate"] as JObject ?? throw SchwabRawError("missing bond coupon");
                CheckFields(yield, "percentage type");
                if (!name.StartsWith("US Treasury Bond - ", StringComparison.Ordinal)
                    || Number(fixedIncome, "face_value") != 100 || Text(yield, "type") != "coupon") throw SchwabRawError("unsupported fixed-income instrument");
                name = $"US Treasury Bond - {Number(yield, "percentage").ToString("G29", CultureInfo.InvariantCulture)}% {Date(fixedIncome, "maturity_date"):dd/MM/yyyy} USD 100";
            }
            return new(id, name, type, OptionalText(row, "ticker_symbol"), quantity, price, value, priceDate);
        }
    }

    private static void CheckFields(JObject row, string fields)
    {
        var allowed = fields.Split(' ').ToHashSet(StringComparer.Ordinal);
        if (row.Properties().Any(p => !allowed.Contains(p.Name))) throw SchwabRawError("unexpected object field; schema review required");
    }

    internal static StatementRecordHoldingImport BuildSchwabRawImport(SchwabRawReport report, List<SchwabRawReport> history,
        Account account, List<Holding> beginning, DateTime sourceTime, Func<string, HoldingType> resolveEquity)
    {
        if (account.relativeBalance || account.isCredit || account.usage != AccountUsage.Investment || account._primaryAccount_Id.HasValue)
            throw SchwabRawError("a primary absolute-balance investment account is required");
        // Import revision order matters when quotes change more than once on the same date.
        var previous = history.LastOrDefault();
        if (previous is not null && (report.Start > previous.End || report.End < previous.End
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
        var securities = history.SelectMany(r => r.Securities).Concat(report.Securities).GroupBy(p => p.Id)
            .ToDictionary(g => g.Key, g => g.Last(), StringComparer.Ordinal);
        var ending = report.Positions.Select(p => BuildRawHolding(p, account, resolveEquity)).ToList();
        if (ending.Select(h => (h.code, h.holdingType)).Distinct().Count() != ending.Count
            || ending.Count(h => h.holdingType == HoldingType.Cash) != 1)
            throw SchwabRawError("ambiguous holding identity or missing USD cash holding");
        RawEqual(ending.Sum(h => h.totalPrice.v), report.Total, "holding details versus account total");
        if (report.Positions.Any(p => p.PriceDate > report.End)) throw SchwabRawError("holding price is newer than query coverage");
        if (previous is not null)
        {
            var expected = previous.Positions.Select(p => BuildRawHolding(p, account, resolveEquity)).ToList();
            if (!HoldingState(beginning).SequenceEqual(HoldingState(expected)))
                throw SchwabRawError("database holdings differ from the last imported source");
        }
        var cash = beginning.Where(h => h.holdingType == HoldingType.Cash).Sum(h => h.totalPrice.v);
        var quantities = beginning.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => (h.code, h.holdingType), h => h.quantity);
        var values = beginning.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => (h.code, h.holdingType), h => h.totalPrice.v);
        var latestTradeDates = new Dictionary<(string, HoldingType), DateTime>();
        var records = new List<Record>();
        foreach (var tx in report.Transactions)
        {
            if (known.ContainsKey(tx.Id)) continue;
            RawEqual(tx.Amount, Decimal.Round(tx.Amount, 2), "transaction amount precision");
            RawEqual(tx.Fees, Decimal.Round(tx.Fees, 2), "transaction fee precision");
            if (tx.Date > report.End) throw SchwabRawError("transaction is later than query coverage");
            var source = SchwabRawPrefix + "transaction/" + RawHash(tx.Id);
            if (tx.Type is "buy" or "sell")
            {
                var priorPosition = previous?.Positions.SingleOrDefault(p => p.Id == tx.SecurityId);
                if (priorPosition is not null && tx.TradeDate < priorPosition.PriceDate)
                    throw SchwabRawError("trade is backdated before this security's prior valuation");
                if (!securities.TryGetValue(tx.SecurityId, out var security)) throw SchwabRawError("trade security metadata is missing");
                var holding = BuildRawHolding(security, account, resolveEquity);
                if (holding.holdingType == HoldingType.Cash || tx.Price <= 0 || tx.Fees < 0
                    || (tx.Type == "buy" ? tx.Quantity <= 0 || tx.Amount <= 0 : tx.Quantity >= 0 || tx.Amount >= 0))
                    throw SchwabRawError("invalid signed security trade");
                var principal = Decimal.Round(tx.Quantity * tx.Price / (holding.holdingType == HoldingType.UST ? 100m : 1m), 2, MidpointRounding.AwayFromZero);
                // User-authorized exception ONLY for Schwab UST trades: the signed
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
                if (tx.TradeDate > latestTradeDates.GetValueOrDefault(key)) latestTradeDates[key] = tx.TradeDate;
                quantities[key] = quantities.GetValueOrDefault(key) + tx.Quantity;
                values[key] = values.GetValueOrDefault(key) + principal;
                Add(principal, tx.Type == "buy" ? "买入" : "卖出", source + "/asset", tx.TradeDate, tx.Date, true, holding, tx.Quantity);
                Add(-principal, tx.Type == "buy" ? "买入" : "卖出", source + "/cash", tx.TradeDate, tx.Date, true, null, 0);
                if (tx.Fees != 0) Add(-tx.Fees, "手续费", source + "/fee", tx.TradeDate, tx.Date, false, null, 0);
                // This interest is already paid/received in cash, not an outstanding accrued holding.
                if (settledAccruedInterest != 0)
                    Add(-settledAccruedInterest, "债券利息", source + "/accrued-interest-settlement",
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
                Add(-tx.Amount, reason, source + "/cash", tx.TradeDate, tx.Date, false, null, 0);
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
                var position = report.Positions.FirstOrDefault(p =>
                {
                    var current = BuildRawHolding(p, account, resolveEquity);
                    return (current.code, current.holdingType) == key;
                });
                var priorPosition = previous?.Positions.FirstOrDefault(p =>
                {
                    var prior = BuildRawHolding(p, account, resolveEquity);
                    return (prior.code, prior.holdingType) == key;
                });
                if (position is not null && priorPosition is not null && position.PriceDate < priorPosition.PriceDate)
                    throw SchwabRawError("changed valuation has stale price date");
                // A trade can change value using an older quote; the resulting value cannot predate that trade.
                var priceDate = position?.PriceDate ?? priorPosition?.PriceDate ?? report.End;
                if (latestTradeDates.GetValueOrDefault(key) > priceDate) priceDate = latestTradeDates[key];
                Add(change, "持仓价格变动", SchwabRawPrefix + "valuation/" + RawHash(JsonSerializer.Serialize(report)) + "/" + holding.code,
                    priceDate, report.AsOf.Date, false, holding, 0);
            }
        }
        var beginningValue = beginning.Sum(h => h.totalPrice.v);
        RawEqual(beginningValue + records.Sum(r => r.v), report.Total, "opening value plus records");
        return new(SchwabRawProvider, sourceTime, SchwabRawPrefix + account.Id + "/" + RawHash(JsonSerializer.Serialize(report)),
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

    private static void RawEqual(decimal left, decimal right, string field)
    {
        if (left != right) throw SchwabRawError(field + $" mismatch: expected={left}, actual={right}");
    }
    private static string RawHash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    private static MailParseException SchwabRawError(string message) => new("Plaid Schwab: " + message);

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
        public List<SchwabRawPosition> Securities { get; set; } = [];
        public Dictionary<string, string> Markets { get; set; } = new(StringComparer.Ordinal);
        public List<SchwabRawPosition> Positions { get; set; } = [];
        public List<SchwabRawTransaction> Transactions { get; set; } = [];
    }
    internal sealed record SchwabRawPosition(string Id, string Name, string Type, string Symbol, decimal Quantity, decimal Price, decimal Value, DateTime PriceDate);
    internal sealed record SchwabRawTransaction(string Id, string SecurityId, DateTime Date, DateTime TradeDate, string Type, string Subtype,
        decimal Quantity, decimal Amount, decimal Price, decimal Fees, string Name = "");
}
