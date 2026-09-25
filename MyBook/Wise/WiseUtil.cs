using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace MyBook;

// Personal-token, read-only API. Tested inaccessible on 2026-09-24; do not call:
// GET /v1/profiles/{id}/balance-statements/{balanceId}/statement.json: 403, SCA required.
// GET /v1/identity/one-time-token/status: 404, ott_not_found.
// GET /v4/spend/profiles/{id}/cards/transactions/{id}: 403 (also documented 90-day limit).
// GET /2026Q3/profiles/{id}/account-details: 403, access.denied.
// GET /v2/profiles/{id}/balance-movements/{id}: 404; POST is a money movement, NOT a read API.
// The separate incoming-transfer API requires partner access and is not used.
// Activity is a summary, not a statement: retain unsplittable gross amounts explicitly.
internal sealed partial class WiseUtil(IConfiguration config, DatabaseUtil database)
{
    private const StatementImportProvider Provider = StatementImportProvider.WiseApi;
    private static readonly SemaphoreSlim importLock = new(1, 1);
    private static readonly JsonSerializerSettings JsonSettings = new() { DateParseHandling = DateParseHandling.None, FloatParseHandling = FloatParseHandling.Decimal };
    private static readonly Regex AmountPattern = new(@"^(?<sign>[+-]?)\s*(?<value>\d+(?:,\d{3})*(?:\.\d+)?) (?<currency>[A-Z]{3})$", RegexOptions.CultureInvariant);
    public bool IsConfigured => !String.IsNullOrWhiteSpace(config["wise_api_token"]);

    public async Task FetchAsync(CancellationToken cancellationToken = default)
    {
        if (!await importLock.WaitAsync(0, cancellationToken).ConfigureAwait(false)) return;
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(TimeSpan.FromMinutes(3));
        var stage = "initialization";
        try
        {
            if (!IsConfigured) throw Error("wise_api_token is not configured");
            var account = database.GetAccountByName("WISE");
            if (account.relativeBalance || account.isCredit || account._primaryAccount_Id.HasValue)
                throw Error("a primary absolute-balance cash account is required");
            stage = "load previous import";
            var latest = database.GetLatestStatementImport(Provider);
            var previous = latest is null ? null : JsonConvert.DeserializeObject<SyncState>(latest.sourceDataJson
                ?? throw Error("stored metadata missing"), JsonSettings) ?? throw Error("invalid stored metadata");
            stage = "load current holdings";
            var beginning = database.GetCurrentAccountHoldings(account);
            if (previous is null && (database.HasAccountHistory(account) || HoldingBalances(beginning).Values.Any(v => v != 0)))
                throw Error("existing Wise history requires explicit cleanup before migration");
            if (previous is not null && (previous.Version != 2 || previous.AccountId != account.Id))
                throw Error("stored account identity or metadata version changed");
            stage = "load fixed checkpoint";
            var checkpoint = database.GetStatementImportCheckpointTime(Provider)
                ?? database.GetStatementImportCheckpointTime(StatementImportProvider.WiseMail);
            DateTimeOffset? start = checkpoint.HasValue ? new DateTimeOffset(checkpoint.Value) : null;
            if (previous is not null && previous.Start != start) throw Error("fixed checkpoint changed");
            var scanTime = DateTimeOffset.UtcNow;
            if (previous?.ScannedThrough > scanTime.AddMinutes(5)) throw Error("invalid stored scan time");
            var since = previous?.ScannedThrough?.AddDays(-7) ?? start;
            if (start.HasValue && since < start) since = start;
            var oldEvents = (previous?.Events ?? []).ToDictionary(e => EventKey(e.Activity), StringComparer.Ordinal);

            using var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseProxy = false })
                { BaseAddress = new Uri("https://api.wise.com"), Timeout = TimeSpan.FromSeconds(15) };
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", config["wise_api_token"]);
            client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en");
            async Task<JToken> Get(string path, string endpoint)
            {
                stage = "GET " + endpoint;
                try
                {
                    using var response = await client.GetAsync(path, deadline.Token).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw Error(stage + $": HTTP {(int)response.StatusCode}");
                    using var reader = new JsonTextReader(new System.IO.StringReader(
                        await response.Content.ReadAsStringAsync(deadline.Token).ConfigureAwait(false)))
                        { FloatParseHandling = FloatParseHandling.Decimal, DateParseHandling = DateParseHandling.None };
                    return JToken.ReadFrom(reader);
                }
                catch (HttpRequestException) { throw Error(stage + ": transport failure"); }
                catch (JsonException) { throw Error(stage + ": invalid JSON"); }
            }

            async Task<ReceiptData> Receipt(string id)
            {
                stage = "GET /2026Q3/transfers/{id}/receipt.pdf";
                using var requestDeadline = CancellationTokenSource.CreateLinkedTokenSource(deadline.Token);
                requestDeadline.CancelAfter(TimeSpan.FromSeconds(15));
                using var response = await client.GetAsync($"/2026Q3/transfers/{id}/receipt.pdf",
                    HttpCompletionOption.ResponseHeadersRead, requestDeadline.Token).ConfigureAwait(false);
                if (response.StatusCode == HttpStatusCode.NotFound) return new(404, null, null);
                if (!response.IsSuccessStatusCode) throw Error(stage + $": HTTP {(int)response.StatusCode}");
                using var stream = await response.Content.ReadAsStreamAsync(requestDeadline.Token).ConfigureAwait(false);
                using var buffer = new System.IO.MemoryStream();
                var bytes = new byte[8192];
                int read;
                while ((read = await stream.ReadAsync(bytes, requestDeadline.Token).ConfigureAwait(false)) != 0)
                {
                    if (buffer.Length + read > 4_000_000) throw Error(stage + ": document too large");
                    buffer.Write(bytes, 0, read);
                }
                return ParseReceipt(buffer.ToArray(), id);
            }

            var profiles = Array(await Get("/v1/profiles", "/v1/profiles"));
            var personal = profiles.Where(p => Text(p, "type") == "personal").ToList();
            if (personal.Count != 1) throw Error("expected exactly one personal profile");
            var profile = Id(personal[0], "id");
            if (previous is not null && previous.Profile != profile) throw Error("personal profile changed");
            var balances = ParseBalances(await Get($"/v4/profiles/{profile}/balances?types=STANDARD,SAVINGS", "/v4/profiles/{profile}/balances"));
            if (previous is not null)
            {
                if (previous.Balances.Any(b => !balances.Any(n => n.Id == b.Id && n.Currency == b.Currency)))
                    throw Error("a previous balance account disappeared or changed currency");
                EqualBalances(HoldingBalances(beginning), SumBalances(previous.Balances), "previous import versus database");
            }

            async Task<List<JObject>> Activities()
            {
                var result = new List<JObject>();
                var cursors = new HashSet<string>(StringComparer.Ordinal);
                string? cursor = null;
                do
                {
                    var query = "?size=100";
                    if (since.HasValue) query += "&since=" + Uri.EscapeDataString(since.Value.UtcDateTime.ToString("O", CultureInfo.InvariantCulture));
                    if (cursor is not null) query += "&nextCursor=" + Uri.EscapeDataString(cursor);
                    var page = await Get($"/v1/profiles/{profile}/activities{query}", "/v1/profiles/{profile}/activities");
                    var rows = Array(page["activities"]);
                    result.AddRange(rows);
                    if (page["cursor"] is null) throw Error("activity pagination cursor missing");
                    cursor = Optional(page, "cursor");
                    if (cursor is not null && (rows.Count == 0 || !cursors.Add(cursor) || cursors.Count >= 1000))
                        throw Error("activity pagination did not terminate");
                } while (cursor is not null);
                if (result.Select(a => Text(a, "id")).Distinct(StringComparer.Ordinal).Count() != result.Count
                    || result.Select(EventKey).Distinct(StringComparer.Ordinal).Count() != result.Count)
                    throw Error("duplicate activity/resource");
                return result.OrderBy(EventKey, StringComparer.Ordinal).ToList();
            }

            var activities = await Activities();
            var currentKeys = activities.Select(EventKey).ToHashSet(StringComparer.Ordinal);
            if (oldEvents.Values.Any(e => Text(e.Activity, "status") == "COMPLETED"
                && (!since.HasValue || Stamp(Text(e.Activity, "createdOn")) > since.Value)
                && !currentKeys.Contains(EventKey(e.Activity))))
                throw Error("a previously completed activity disappeared; explicit reconciliation required");
            var events = new Dictionary<string, EventData>(oldEvents, StringComparer.Ordinal);
            var records = new List<Record>();
            // Only new/changed activities need enrichment. Saved source details remain private in the database.
            foreach (var activity in activities)
            {
                deadline.Token.ThrowIfCancellationRequested();
                var key = EventKey(activity);
                if (oldEvents.TryGetValue(key, out var old) && Fingerprint(old.Activity) == Fingerprint(activity))
                {
                    continue;
                }
                if (old is not null && Text(old.Activity, "status") == "COMPLETED")
                    throw Error("a completed activity changed; explicit reconciliation required (" + Hash(key)[..12] + ")");
                var data = new EventData(activity, null, null, null);
                if (Text(activity, "status") == "COMPLETED" && Text(activity["resource"]!, "type") == "TRANSFER")
                {
                    var id = Id(activity["resource"]!, "id");
                    var transfer = (JObject)await Get($"/v1/transfers/{id}", "/v1/transfers/{id}");
                    if (Id(transfer, "id") != id) throw Error("transfer identity mismatch");
                    JObject? quote = null;
                    if (Optional(transfer, "quoteUuid") is string quoteId)
                    {
                        if (!Guid.TryParse(quoteId, out _)) throw Error("invalid quote identifier");
                        quote = (JObject)await Get($"/v3/profiles/{profile}/quotes/{quoteId}", "/v3/profiles/{profile}/quotes/{id}");
                        if (Text(quote, "id") != quoteId) throw Error("quote identity mismatch");
                    }
                    var target = Id(transfer, "targetAccount");
                    var recipient = (JObject)await Get($"/v1/accounts/{target}", "/v1/accounts/{id}");
                    if (Id(recipient, "id") != target) throw Error("recipient identity mismatch");
                    stage = "GET /2026Q3/transfers/{id}/invoices/bankingpartner";
                    using var payoutResponse = await client.GetAsync($"/2026Q3/transfers/{id}/invoices/bankingpartner", deadline.Token).ConfigureAwait(false);
                    if (!payoutResponse.IsSuccessStatusCode && payoutResponse.StatusCode != HttpStatusCode.NotFound)
                        throw Error(stage + $": HTTP {(int)payoutResponse.StatusCode}");
                    JObject? payout = null;
                    if (payoutResponse.IsSuccessStatusCode)
                    {
                        using var reader = new JsonTextReader(new System.IO.StringReader(await payoutResponse.Content.ReadAsStringAsync(deadline.Token).ConfigureAwait(false)))
                            { FloatParseHandling = FloatParseHandling.Decimal, DateParseHandling = DateParseHandling.None };
                        payout = JObject.Load(reader);
                    }
                    data = new(activity, transfer, quote, recipient, payout, (int)payoutResponse.StatusCode, await Receipt(id));
                }
                stage = "activity parsing " + Hash(key)[..12];
                var parsed = ParseEvent(data, account);
                var evidence = FindCounterparty(data);
                ApplyCounterparty(parsed, account, evidence);
                // Keep the exact same-event relationship without hiding unsplit fees via matchedRecordId.
                data = data with { Match = evidence, RecordSources = parsed.Select(r => r.Source).ToList() };
                records.AddRange(parsed);
                events[key] = data;
            }
            var verification = await Activities();
            if (!activities.Select(Fingerprint).SequenceEqual(verification.Select(Fingerprint)))
                throw Error("activities changed during retrieval; retry next run");
            var end = ParseBalances(await Get($"/v4/profiles/{profile}/balances?types=STANDARD,SAVINGS", "/v4/profiles/{profile}/balances"));
            if (!balances.SequenceEqual(end)) throw Error("balances changed during retrieval; retry next run");
            stage = "balance reconciliation";
            var ending = SumBalances(end);
            var opening = HoldingBalances(beginning);
            if (previous is null)
            {
                // Same explicit first-import baseline as the former Plaid importer, not a recurring residual.
                opening = ending.ToDictionary(p => p.Key, p => p.Value - records.Where(r => r.t == p.Key).Sum(r => r.v));
                beginning = Holdings(account, opening);
            }
            var expected = opening.Keys.Union(records.Select(r => r.t)).ToDictionary(c => c,
                c => opening.GetValueOrDefault(c) + records.Where(r => r.t == c).Sum(r => r.v));
            EqualBalances(expected, ending, "opening plus records versus API balance");
            var state = new SyncState(2, profile, account.Id, start, end,
                events.OrderBy(e => e.Key, StringComparer.Ordinal).Select(e => e.Value).ToList(), scanTime);
            var source = JsonConvert.SerializeObject(state);
            Console.WriteLine($"Wise API: scanned activities={activities.Count}, new records={records.Count}, currencies={ending.Count}");
            if (previous is not null && records.Count == 0 && previous.Balances.SequenceEqual(end)
                && previous.ScannedThrough?.UtcDateTime.Date == scanTime.UtcDateTime.Date
                && events.Count == oldEvents.Count && events.Values.All(e => oldEvents.TryGetValue(EventKey(e.Activity), out var old)
                    && Fingerprint(e.Activity) == Fingerprint(old.Activity))) return;
            stage = "atomic database write";
            deadline.Token.ThrowIfCancellationRequested();
            database.SaveStatementRecordsAndHoldingsOnce([new(Provider, DateTime.Today, "WiseApi/" + Hash(source), account,
                records, Holdings(account, ending), ending.Select(p => new AccountBalance(account, new(p.Value, p.Key))).ToList(),
                opening.Select(p => new AccountBalance(account, new(p.Value, p.Key))).ToList(), beginning,
                sourceDataJson: source, forceValidateBeginningBalances: previous is not null)]);
        }
        catch (MailParseException e)
        {
            const string prefix = "Wise API: ";
            throw Error(stage + ": " + (e.Message.StartsWith(prefix, StringComparison.Ordinal)
                ? e.Message[prefix.Length..] : "source validation failed; private details suppressed"));
        }
        catch (OperationCanceledException) { throw Error(stage + ": cancelled or timeout"); }
        catch (Exception e)
        {
            var code = e.GetType().Name == "MySqlException" ? e.GetType().GetProperty("Number")?.GetValue(e) as int? : null;
            throw Error(stage + ": " + e.GetType().Name + (code.HasValue ? $" code={code}" : "") + "; private details suppressed");
        }
        finally { importLock.Release(); }
    }

    internal static List<Record> ParseEvent(EventData data, Account account)
    {
        var a = data.Activity;
        var type = Text(a, "type");
        var resourceType = Text(a["resource"] ?? throw Error("resource missing"), "type");
        var supported = type switch
        {
            "INTERBALANCE" => "BALANCE_TRANSACTION", "TRANSFER" or "BALANCE_DEPOSIT" => "TRANSFER",
            "CARD_PAYMENT" => "CARD_TRANSACTION", "DIRECT_DEBIT_TRANSACTION" => "DIRECT_DEBIT_TRANSACTION", "CARD_ORDER" => "CARD_ORDER",
            _ => throw Error("unsupported activity type")
        };
        if (resourceType != supported) throw Error("activity/resource type mismatch");
        var status = Text(a, "status");
        if (status == "CANCELLED") return [];
        if (status != "COMPLETED") throw Error("unsettled activity; retry after settlement");
        if (type == "CARD_ORDER") throw Error("completed card order accounting is not supported");
        var primary = ParseAmount(Text(a, "primaryAmount"));
        var secondary = Optional(a, "secondaryAmount") is string secondaryText ? ParseAmount(secondaryText) : null;
        var description = Plain(Text(a, "title"));
        if (Optional(a, "description") is string detail) description += "; " + Plain(detail);
        if (Optional(data.Transfer?["details"], "reference") is string reference) description += "; " + reference;
        // Counterparty evidence is resolved separately; parsing must never register account aliases.
        // createdOn is the activity time; updatedOn is not proof of a bank posting time.
        var time = Stamp(Text(a, "createdOn"));
        var result = new List<Record>();
        void Add(Currency amount, string reason, string suffix, Currency? original = null)
        {
            if (amount.v == 0) return;
            result.Add(new Record { Account = account, v = amount.v, t = amount.t, date = time.LocalDateTime,
                updateTime = DateTime.Now, Reason = reason, DestAccount = description.Length > 200 ? description[..200] : description,
                Source = "WiseApi/" + Hash(EventKey(a)) + "/" + suffix, DescCurrency = original, isInternal = false });
        }
        if (type == "INTERBALANCE")
        {
            if (secondary is null || primary.Money.t == secondary.Money.t || primary.Sign != "" || secondary.Sign != "")
                throw Error("unsupported conversion amounts");
            Add(new(-secondary.Money.v, secondary.Money.t), "换汇", "debit");
            Add(primary.Money, "换汇", "credit");
            foreach (var record in result)
            {
                record.isInternal = true;
                record.Source += "; code=BALANCE-WISE-" + Hash(EventKey(a));
            }
        }
        else if (type is "TRANSFER" or "BALANCE_DEPOSIT")
        {
            var transfer = data.Transfer ?? throw Error("transfer details missing");
            if (Text(transfer, "status") != "outgoing_payment_sent") throw Error("transfer is not settled");
            var sourceCurrency = ParseCurrency(Text(transfer, "sourceCurrency"));
            var targetCurrency = ParseCurrency(Text(transfer, "targetCurrency"));
            var targetValue = Money(transfer, "targetValue");
            if (primary.Money.t != targetCurrency || primary.Money.v != targetValue)
                throw Error("transfer target differs from activity");
            if (primary.Sign == "+")
            {
                if (secondary is not null) throw Error("unsupported incoming transfer secondary amount");
                Add(primary.Money, "转账", "credit");
            }
            else
            {
                if (type == "BALANCE_DEPOSIT" || primary.Sign != "") throw Error("unsupported transfer direction");
                var gross = secondary?.Money ?? primary.Money;
                if (secondary is not null && secondary.Sign != "" || gross.t != sourceCurrency)
                    throw Error("transfer source differs from activity");
                decimal? fee = null;
                if (data.Quote is JObject quote && Optional(quote, "preferredPayIn") == "BALANCE")
                {
                    if (Text(quote, "status") != "FUNDED") throw Error("balance quote is not funded");
                    if (ParseCurrency(Text(quote, "sourceCurrency")) != sourceCurrency
                        || ParseCurrency(Text(quote, "targetCurrency")) != targetCurrency)
                        throw Error("quote currency mismatch");
                    var options = Array(quote["paymentOptions"]).Where(o => Text(o, "payIn") == "BALANCE"
                        && Text(o, "payOut") == Text(quote, "payOut") && Money(o, "sourceAmount") == gross.v
                        && Money(o, "targetAmount") == targetValue).ToList();
                    if (options.Count != 1) throw Error("funded balance quote option is not unique");
                    fee = Money(options[0]["fee"] ?? throw Error("quote fee missing"), "total");
                    if (fee < 0 || Money(transfer, "sourceValue") + fee != gross.v)
                        throw Error("transfer principal plus quoted fee differs from gross debit");
                }
                if (fee.HasValue)
                {
                    Add(new(-(gross.v - fee.Value), gross.t), "转账", "principal", new(-targetValue, targetCurrency));
                    Add(new(-fee.Value, gross.t), "手续费", "fee");
                }
                else Add(new(-gross.v, gross.t), "转账", "gross", new(-targetValue, targetCurrency));
            }
        }
        else
        {
            if (primary.Sign != "" || secondary is not null) throw Error("unsupported card/direct debit amounts");
            Add(new(-primary.Money.v, primary.Money.t), type == "CARD_PAYMENT" ? "消费" : "转账", "debit");
        }
        return result;
    }

    internal static List<BalanceData> ParseBalances(JToken json)
    {
        var balances = new List<BalanceData>();
        foreach (var b in Array(json))
        {
            if (Text(b, "type") is not ("STANDARD" or "SAVINGS") || Text(b, "investmentState") != "NOT_INVESTED")
                throw Error("unsupported balance/investment type");
            var currency = ParseCurrency(Text(b, "currency"));
            decimal Read(string key)
            {
                var value = b[key] ?? throw Error("balance field missing");
                if (ParseCurrency(Text(value, "currency")) != currency) throw Error("balance currency mismatch");
                return Money(value, "value");
            }
            var amount = Read("amount");
            if (Read("reservedAmount") != 0) throw Error("reserved funds present; retry after settlement");
            if (Read("cashAmount") != amount || Read("totalWorth") != amount) throw Error("cash balance fields disagree");
            balances.Add(new(Id(b, "id"), currency, amount));
        }
        if (balances.Count == 0 || balances.Select(b => b.Id).Distinct().Count() != balances.Count) throw Error("empty or duplicate balances");
        return balances.OrderBy(b => b.Id, StringComparer.Ordinal).ToList();
    }

    private static Dictionary<CurrencyType, decimal> SumBalances(List<BalanceData> balances) =>
        balances.GroupBy(b => b.Currency).ToDictionary(g => g.Key, g => g.Sum(b => b.Amount));
    private static Dictionary<CurrencyType, decimal> HoldingBalances(List<Holding> holdings)
    {
        if (holdings.Any(h => h.holdingType != HoldingType.Cash || h.code != h.currentPrice.t.ToString())
            || holdings.Select(h => h.currentPrice.t).Distinct().Count() != holdings.Count) throw Error("unexpected local holdings");
        return holdings.ToDictionary(h => h.currentPrice.t, h => h.totalPrice.v);
    }
    private static List<Holding> Holdings(Account account, Dictionary<CurrencyType, decimal> balances) =>
        balances.Select(p => new Holding(p.Key.ToString(), HoldingType.Cash) { Account = account, currentPrice = new(p.Value, p.Key) }).ToList();
    private static void EqualBalances(Dictionary<CurrencyType, decimal> left, Dictionary<CurrencyType, decimal> right, string context)
    {
        foreach (var currency in left.Keys.Union(right.Keys))
            if (left.GetValueOrDefault(currency) != right.GetValueOrDefault(currency))
                throw Error(context + $": {currency} expected={left.GetValueOrDefault(currency)}, actual={right.GetValueOrDefault(currency)}; no data saved");
    }
    private static AmountData ParseAmount(string text)
    {
        var match = AmountPattern.Match(Plain(text));
        if (!match.Success) throw Error("unsupported activity amount format");
        var value = Decimal.Parse(match.Groups["value"].Value, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
        if (value <= 0 || Decimal.Round(value, 2) != value) throw Error("invalid activity amount precision/value");
        return new(new(value, ParseCurrency(match.Groups["currency"].Value)), match.Groups["sign"].Value);
    }
    private static CurrencyType ParseCurrency(string code) => code switch
    {
        "USD" => CurrencyType.USD, "CNY" => CurrencyType.RMB, "SGD" => CurrencyType.SGD, "EUR" => CurrencyType.EUR,
        "GBP" => CurrencyType.GBP, "HKD" => CurrencyType.HKD, "JPY" => CurrencyType.JPY, _ => throw Error("unsupported currency")
    };
    private static decimal Money(JToken node, string field)
    {
        var value = node[field];
        if (value?.Type is not (JTokenType.Integer or JTokenType.Float)) throw Error("numeric field missing or invalid");
        var number = value.Value<decimal>();
        if (Decimal.Round(number, 2) != number) throw Error("unsupported sub-cent precision");
        return number;
    }
    private static string Plain(string text) => WebUtility.HtmlDecode(Regex.Replace(text, "<[^>]*>", "")).Trim();
    private static List<JObject> Array(JToken? node) => node is JArray array && array.All(n => n is JObject)
        ? array.Cast<JObject>().ToList() : throw Error("expected array of objects");
    private static string? Optional(JToken? node, string field) => node?[field]?.Type is null or JTokenType.Null ? null
        : node[field]!.Type == JTokenType.String ? String.IsNullOrWhiteSpace(node[field]!.Value<string>()) ? null : node[field]!.Value<string>()
        : throw Error("invalid string field");
    private static string Text(JToken node, string field) => Optional(node, field) ?? throw Error("required string missing");
    private static string Id(JToken node, string field)
    {
        var token = node[field];
        if (token?.Type is not (JTokenType.Integer or JTokenType.String)
            || !Int64.TryParse(token.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id <= 0)
            throw Error("invalid numeric identifier");
        return id.ToString(CultureInfo.InvariantCulture);
    }
    private static string EventKey(JObject activity) => Text(activity["resource"] ?? throw Error("resource missing"), "type")
        + "/" + Id(activity["resource"]!, "id");
    internal static string Fingerprint(JObject activity)
    {
        JToken Ordered(JToken value) => value switch
        {
            JObject obj => new JObject(obj.Properties().OrderBy(p => p.Name, StringComparer.Ordinal)
                .Select(p => new JProperty(p.Name, Ordered(p.Value)))),
            JArray array => new JArray(array.Select(Ordered)),
            _ => value.DeepClone()
        };
        // MySQL's JSON storage reorders nested object properties; order is not a source revision.
        var source = new JObject(activity.Properties().Where(p => p.Name != "updatedOn")
            .Select(p => new JProperty(p.Name, p.Value.DeepClone())));
        return Ordered(source).ToString(Formatting.None);
    }
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    private static DateTimeOffset Stamp(string value) => DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
        && date <= DateTimeOffset.UtcNow.AddMinutes(5) ? date : throw Error("invalid activity date");
    private static MailParseException Error(string message) => new("Wise API: " + message);
    private sealed record AmountData(Currency Money, string Sign);
    internal sealed record BalanceData(string Id, CurrencyType Currency, decimal Amount);
    internal sealed record ReceiptData(int Status, string? Text, string? Sha256);
    internal sealed record AccountMatch(string Status, string[] Identifiers, string[] References, string? AccountName);
    internal sealed record EventData(JObject Activity, JObject? Transfer, JObject? Quote, JObject? Recipient,
        JObject? Payout = null, int? PayoutStatus = null, ReceiptData? Receipt = null, AccountMatch? Match = null,
        List<string>? RecordSources = null);
    private sealed record SyncState(int Version, string Profile, int AccountId, DateTimeOffset? Start, List<BalanceData> Balances,
        List<EventData> Events, DateTimeOffset? ScannedThrough = null);
}
