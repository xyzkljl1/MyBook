using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using SteamKit2;

namespace MyBook;

partial class WebUtil
{
    public async Task FetchSteamAsync(DateTime since)
    {
        var account = database.GetAccountByName("Steam");
        if (account.relativeBalance || account.isCredit)
            throw new InvalidOperationException("Steam import: an absolute-balance wallet account is required.");
        var beginning = database.GetAccountBalance(account, CurrencyType.RMB);
        var previous = database.GetStatementRecords(StatementImportProvider.SteamWeb, account);
        var bankPurchases = database.GetSteamBankPurchases(account.Id);
        var walletSince = since.Date.AddDays(-15);
        var purchaseSince = since.Date.AddDays(-60);
        await WithSteamSessionAsync(async (readInfo, accessToken, token) =>
        {
            var info = await readInfo().ConfigureAwait(false);
            if (!info.HasWallet || info.Currency != ECurrencyCode.CNY || info.PendingBalance != 0)
                throw new InvalidOperationException("Steam import: unsupported wallet currency or pending balance.");
            using var http = CreateSteamHistoryClient(info.SteamId, accessToken);
            var purchases = new List<SteamExternalPurchase>();
            var entries = await ReadSteamHistoryAsync(http, walletSince, token, purchases, purchaseSince).ConfigureAwait(false);
            var now = DateTime.Now;
            var records = BuildSteamWalletRecords(entries, account, previous, beginning.v, info.Balance, now);
            var supplements = BuildSteamPurchaseSupplements(purchases, bankPurchases);
            var supplemented = 0;
            if (records.Count > 0 || supplements.Count > 0)
                database.SaveStatementRecordsOnce(StatementImportProvider.SteamWeb, now.Date, records,
                    accountBalances: [new(account, new Currency(info.Balance, CurrencyType.RMB))],
                    statementKey: $"Steam/{account.Id}/{now:yyyyMMddHHmmssfffffff}",
                    beginningAccountBalances: [new(account, beginning)], forceValidateBeginningBalances: true,
                    afterSaveInTransaction: id =>
                    {
                        if (records.Count > 0)
                            database.MarkRecordsAsRefundMatched(
                                FindSteamRefundMatches(previous.Concat(database.GetRecordsByStatementImport(id))), withinImportTransaction: true);
                        supplemented = database.AppendSteamPurchaseSupplements(account.Id, purchases);
                    });
            Console.WriteLine($"Steam: {records.Count} wallet record(s); balance {info.Balance:0.00} RMB validated.");
            var pending = purchases.Select(p => p.Id).Distinct().Count(id => !bankPurchases.Any(b => HasSteamOrderCode(b.Record.Source, id)));
            Console.WriteLine($"Steam: {supplemented} bank record source(s) supplemented; {pending - supplemented} external purchase(s) unmatched.");
            return true;
        }).ConfigureAwait(false);
    }

    private HttpClient CreateSteamHistoryClient(ulong steamId, string accessToken)
    {
        var cookies = new CookieContainer();
        foreach (var host in new[] { "store.steampowered.com", "help.steampowered.com" })
            cookies.Add(new Uri("https://" + host), new Cookie("steamLoginSecure",
                Uri.EscapeDataString($"{steamId}||{accessToken}"), "/") { Secure = true, HttpOnly = true });
        var handler = CreateSteamHttpHandler();
        handler.CookieContainer = cookies;
        handler.AllowAutoRedirect = false;
        handler.AutomaticDecompression = DecompressionMethods.All;
        var http = new HttpClient(handler) { BaseAddress = new Uri("https://store.steampowered.com"), Timeout = TimeSpan.FromSeconds(45) };
        http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36");
        http.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.9");
        return http;
    }

    private static async Task<string> RequestSteamHistoryAsync(HttpClient http, string path,
        Dictionary<string, string>? form, CancellationToken token)
    {
        var uri = new Uri(http.BaseAddress!, path);
        for (int redirect = 0, attempt = 1; ; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var label = $"Steam {(form is null ? "GET" : "POST")} {uri.Host}{uri.AbsolutePath}";
            string failure;
            try
            {
                using var request = new HttpRequestMessage(form is null ? HttpMethod.Get : HttpMethod.Post, uri);
                if (form is not null) request.Content = new FormUrlEncodedContent(form);
                using var response = await http.SendAsync(request, token).ConfigureAwait(false);
                if ((int)response.StatusCode is 301 or 302 or 303 or 307 or 308)
                {
                    var next = response.Headers.Location is { } location ? new Uri(uri, location) : null;
                    if (redirect >= 3 || next is null || next.Scheme != "https"
                        || !next.IsDefaultPort || next.UserInfo.Length > 0
                        || (form is null
                            ? next.Host != "help.steampowered.com" || next.AbsolutePath is not ("/en/wizard/HelpWithTransaction" or "/en/wizard/HelpWithMyPurchase")
                            : (int)response.StatusCode is not (302 or 307 or 308) || next.Host != uri.Host
                                || next.AbsolutePath != uri.AbsolutePath || next.AbsolutePath != "/account/AjaxLoadMoreHistory/"))
                        throw new InvalidOperationException($"{label}: unexpected redirect; HTTP {(int)response.StatusCode}.");
                    uri = next;
                    redirect++;
                    attempt = 0;
                    continue;
                }
                if ((int)response.StatusCode is 408 or 429 or >= 500 and <= 599)
                    throw new HttpRequestException(null, null, response.StatusCode);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException($"{label}: HTTP {(int)response.StatusCode}.");
                return await response.Content.ReadAsStringAsync(token).ConfigureAwait(false);
            }
            catch (HttpRequestException error)
            { failure = error.StatusCode is { } status ? $"HTTP {(int)status}" : error.HttpRequestError.ToString(); }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            { failure = "request timed out"; }
            if (attempt == 3) throw new InvalidOperationException($"{label}: {failure}; failed after {attempt} attempts.");
            Console.WriteLine($"{label}: {failure}; attempt {attempt}/3 failed, retrying in 2 seconds.");
            await Task.Delay(TimeSpan.FromSeconds(2), token).ConfigureAwait(false);
        }
    }

    private static async Task<List<SteamWalletEntry>> ReadSteamHistoryAsync(HttpClient http, DateTime since, CancellationToken token,
        List<SteamExternalPurchase>? purchases = null, DateTime? purchaseSince = null)
    {
        var html = await RequestSteamHistoryAsync(http, "/account/history/?l=english", null, token).ConfigureAwait(false);
        var document = new HtmlDocument(); document.LoadHtml(html);
        if (document.DocumentNode.SelectSingleNode("//table[contains(@class,'wallet_history_table')]") is null)
            throw new InvalidOperationException("Steam history: purchase-history table missing; authorization or page format changed.");
        var cursorMatch = Regex.Match(html, @"\bg_historyCursor\s*=\s*(?<cursor>null|\{[^;]*?\})\s*;");
        var sessionMatch = Regex.Match(html, "\\bg_sessionID\\s*=\\s*[\"'](?<id>[a-zA-Z0-9]+)[\"']");
        if (!cursorMatch.Success)
            throw new InvalidOperationException("Steam history: pagination cursor missing.");
        var cursor = cursorMatch.Groups["cursor"].Value;
        var result = new List<SteamWalletEntry>();
        var seenCursors = new HashSet<string>();
        for (var page = 0; ; page++)
        {
            var (entries, oldest) = ParseSteamHistoryPage(html, since);
            result.AddRange(entries);
            if (purchases is not null) purchases.AddRange(ParseSteamExternalPurchases(html, purchaseSince ?? since));
            var searchSince = purchaseSince < since ? purchaseSince.Value : since;
            if (cursor == "null" || oldest.HasValue && oldest.Value < searchSince) break;
            if (!sessionMatch.Success || !seenCursors.Add(cursor) || page >= 100)
                throw new InvalidOperationException("Steam history: incomplete pagination.");
            var form = new Dictionary<string, string> { ["sessionid"] = sessionMatch.Groups["id"].Value };
            try
            {
                using var parsedCursor = JsonDocument.Parse(cursor);
                foreach (var property in parsedCursor.RootElement.EnumerateObject())
                {
                    if (property.Value.ValueKind is not (JsonValueKind.String or JsonValueKind.Number))
                        throw new InvalidOperationException("Steam history: unsupported cursor format.");
                    form[$"cursor[{property.Name}]"] = property.Value.ToString();
                }
                using var response = JsonDocument.Parse(await RequestSteamHistoryAsync(http,
                    "/account/AjaxLoadMoreHistory/", form, token).ConfigureAwait(false));
                html = response.RootElement.GetProperty("html").GetString() ?? "";
                if (response.RootElement.TryGetProperty("cursor", out var nextCursor))
                    cursor = nextCursor.GetRawText();
                else if (ParseSteamHistoryPage(html, since).Oldest is DateTime oldestDate && oldestDate < searchSince)
                    cursor = "null"; // The returned rows already cover the requested range.
                else
                    throw new InvalidOperationException("Steam history: pagination cursor missing before search start.");
                if (String.IsNullOrWhiteSpace(html) && cursor != "null")
                    throw new InvalidOperationException("Steam history: empty page before end of history.");
            }
            catch (Exception error) when (error is JsonException or KeyNotFoundException)
            { throw new InvalidOperationException("Steam history: invalid pagination response."); }
        }
        foreach (var entry in result.Where(entry => entry.Reason is "游戏" or "退款"))
        {
            var transactionId = entry.Source.Split('|')[1];
            var detail = await RequestSteamHistoryAsync(http,
                $"https://help.steampowered.com/en/wizard/HelpWithTransaction?transid={transactionId}&l=english", null, token).ConfigureAwait(false);
            entry.Items = ParseSteamOrderItems(detail, entry);
        }
        return result;
    }

    internal sealed record SteamWalletItem(string Id, string Name, decimal Amount);
    internal sealed record SteamWalletEntry(DateTime Date, string Source, string[] ItemNames, string Reason, decimal Change, decimal Balance, decimal Total)
    {
        public List<SteamWalletItem> Items { get; set; } = [];
    }

    internal static (List<SteamWalletEntry> Entries, DateTime? Oldest) ParseSteamHistoryPage(string html, DateTime since)
    {
        var document = new HtmlDocument(); document.LoadHtml(html);
        var entries = new List<SteamWalletEntry>();
        DateTime? oldest = null;
        foreach (var row in document.DocumentNode.SelectNodes("//tr[contains(concat(' ',normalize-space(@class),' '),' wallet_table_row ')]") ?? new HtmlNodeCollection(document.DocumentNode))
        {
            HtmlNode CellNode(string name) => row.SelectSingleNode($"./td[contains(concat(' ',normalize-space(@class),' '),' {name} ')]")
                ?? throw new InvalidOperationException($"Steam history: required {name} column missing.");
            string Cell(string name) => SteamHistoryText(CellNode(name));
            if (!DateTime.TryParse(Cell("wht_date"), CultureInfo.GetCultureInfo("en-US"), DateTimeStyles.None, out var date))
                throw new InvalidOperationException("Steam history: invalid transaction date.");
            oldest = !oldest.HasValue || date < oldest.Value ? date : oldest;
            if (date < since) continue;
            var changeText = Cell("wht_wallet_change");
            var balanceText = Cell("wht_wallet_balance");
            // Card/Alipay payments without wallet movement belong to their actual funding accounts.
            if (changeText.Length == 0) continue;
            var change = ReadSteamYuan(changeText);
            if (change == 0) continue;
            var balance = ReadSteamYuan(balanceText);
            var type = Cell("wht_type");
            var item = Cell("wht_items");
            var itemNames = CellNode("wht_items").SelectNodes("./div[not(contains(concat(' ',normalize-space(@class),' '),' wth_item_refunded '))]")
                ?.Select(SteamHistoryText).ToArray() ?? [item];
            if (itemNames.Length == 0 || itemNames.Any(String.IsNullOrWhiteSpace))
                throw new InvalidOperationException("Steam history: item names missing.");
            var total = ReadSteamYuan(Cell("wht_total"));
            if (total <= 0 || Math.Abs(change) > total)
                throw new InvalidOperationException("Steam history: wallet movement exceeds transaction total.");
            string reason;
            if (type.Contains("Refund", StringComparison.OrdinalIgnoreCase) && change > 0) reason = "退款";
            else if (type.Contains("Purchase", StringComparison.OrdinalIgnoreCase) && change < 0) reason = "游戏";
            else if (change > 0 && item.Contains("Wallet", StringComparison.OrdinalIgnoreCase)) reason = "转账";
            else throw new InvalidOperationException("Steam history: unsupported wallet transaction type or sign.");
            var link = HtmlEntity.DeEntitize(row.GetAttributeValue("onclick", "") + " "
                + row.SelectSingleNode(".//a[@href]")?.GetAttributeValue("href", "")) ?? "";
            var id = Regex.Match(link, @"[?&](?:transid|transactionid)=(\d+)(?:[&'""\s]|$)");
            if (!id.Success) throw new InvalidOperationException("Steam history: wallet transaction identifier missing.");
            entries.Add(new(date, $"Steam transaction|{id.Groups[1].Value}|{(change < 0 ? "debit" : "credit")}", itemNames, reason, change, balance, total));
        }
        return (entries, oldest);
    }

    internal static List<SteamWalletItem> ParseSteamOrderItems(string html, SteamWalletEntry entry)
    {
        var document = new HtmlDocument(); document.LoadHtml(html);
        var lines = document.DocumentNode.SelectNodes("//div[@class='purchase_line_items']/div")
            ?? throw new InvalidOperationException("Steam detail: item lines missing.");
        var items = lines.Select(line => new SteamWalletItem("",
            SteamHistoryText(line.SelectSingleNode(".//*[@class='purchase_detail_field']")),
            ReadSteamYuan(SteamHistoryText(line.SelectSingleNode(".//*[@class='refund_value']"))))).ToList();
        if (items.Any(item => String.IsNullOrWhiteSpace(item.Name) || item.Amount <= 0))
            throw new InvalidOperationException("Steam detail: invalid item name or price.");
        // A refund can refer to just one item from the original order.
        if (entry.Reason == "退款") items = items.Where(item => entry.ItemNames.Contains(item.Name, StringComparer.Ordinal)).ToList();
        if (!items.Select(item => item.Name).Order(StringComparer.Ordinal).SequenceEqual(entry.ItemNames.Order(StringComparer.Ordinal))
            || items.Sum(item => item.Amount) != entry.Total)
            throw new InvalidOperationException("Steam detail: item names or total disagree with history.");
        if (items.Count == 1)
            return [items[0] with { Amount = entry.Change }];
        // Mixed payments are explicitly kept as one wallet record because per-item funding is unavailable.
        if (Math.Abs(entry.Change) != entry.Total)
            return [new("", String.Join(" / ", entry.ItemNames), entry.Change)];
        var transactionId = entry.Source.Split('|')[1];
        var links = document.DocumentNode.SelectNodes("//a[contains(concat(' ',normalize-space(@class),' '),' help_purchase_button ')][.//*[@class='help_purchase_price']]")
            ?? throw new InvalidOperationException("Steam detail: item identifiers missing.");
        var result = new List<SteamWalletItem>();
        foreach (var item in items)
        {
            var candidates = links.Where(link => SteamHistoryText(link.SelectSingleNode(".//*[@class='help_purchase_content']")) == item.Name).ToList();
            if (candidates.Count != 1) throw new InvalidOperationException("Steam detail: item identifier is ambiguous.");
            var href = HtmlEntity.DeEntitize(candidates[0].GetAttributeValue("href", "")) ?? "";
            var id = Regex.Match(href, @"[?&]line_item=(\d+)(?:&|$)");
            var transaction = Regex.Match(href, @"[?&]transid=(\d+)(?:&|$)");
            if (!id.Success || !transaction.Success || transaction.Groups[1].Value != transactionId)
                throw new InvalidOperationException("Steam detail: item identifier disagrees with receipt.");
            result.Add(item with { Id = id.Groups[1].Value, Amount = Math.Sign(entry.Change) * item.Amount });
        }
        if (result.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != result.Count)
            throw new InvalidOperationException("Steam detail: duplicate item identifier.");
        return result;
    }

    private static string SteamHistoryText(HtmlNode? node) => Regex.Replace(HtmlEntity.DeEntitize(node?.InnerText ?? "") ?? "", @"\s+", " ").Trim();

    private static decimal ReadSteamYuan(string text)
    {
        var match = Regex.Match(text, @"^([+\-]?)\s*[¥￥]\s*((?:\d+|\d{1,3}(?:,\d{3})+)\.\d{2})$");
        if (!match.Success) throw new InvalidOperationException("Steam history: invalid CNY wallet amount.");
        var value = Decimal.Parse(match.Groups[2].Value, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
        return match.Groups[1].Value == "-" ? -value : value;
    }

    internal static List<Record> FindSteamRefundMatches(IEnumerable<Record> records)
    {
        var candidates = records.Where(record => !record.isRefundMatched && !record.isInternal && record.matchedRecordId is null)
            .Select(record => (Record: record, Key: Regex.Match(record.Source,
                @"^Steam transaction\|(?<transaction>\d+)\|(?<side>debit|credit)(?:\|item/(?<item>\d+))?; (?<name>.+)$")))
            .Where(item => item.Key.Success).ToList();
        var purchases = candidates.Where(item => item.Record.Reason == "游戏" && item.Record.v < 0 && item.Key.Groups["side"].Value == "debit").ToList();
        var refunds = candidates.Where(item => item.Record.Reason == "退款" && item.Record.v > 0 && item.Key.Groups["side"].Value == "credit").ToList();
        static bool Matches((Record Record, Match Key) purchase, (Record Record, Match Key) refund) =>
            purchase.Record._account_Id == refund.Record._account_Id && purchase.Record.t == refund.Record.t
            && purchase.Record.date <= refund.Record.date && purchase.Record.v == -refund.Record.v
            && purchase.Key.Groups["transaction"].Value == refund.Key.Groups["transaction"].Value
            && (purchase.Key.Groups["item"].Success && refund.Key.Groups["item"].Success
                ? purchase.Key.Groups["item"].Value == refund.Key.Groups["item"].Value
                // A single-item refund page can omit the item ID; require a unique name within the same order.
                : purchase.Key.Groups["name"].Value == refund.Key.Groups["name"].Value);
        var matched = new List<Record>();
        foreach (var refund in refunds)
        {
            var expenses = purchases.Where(purchase => Matches(purchase, refund)).ToList();
            if (expenses.Count != 1 || refunds.Count(candidate => Matches(expenses[0], candidate)) != 1)
                continue;
            matched.Add(expenses[0].Record);
            matched.Add(refund.Record);
        }
        return matched;
    }

    internal static List<Record> BuildSteamWalletRecords(List<SteamWalletEntry> entries, Account account,
        List<Record> previous, decimal beginning, decimal ending, DateTime now)
    {
        var known = previous.ToLookup(record => record.Source.Split(';', 2)[0].Split("|item/", 2)[0], StringComparer.Ordinal);
        var seen = new Dictionary<string, SteamWalletEntry>(StringComparer.Ordinal);
        var balance = beginning;
        var records = new List<Record>();
        // Steam returns newest first, including same-day wallet movements.
        foreach (var entry in entries.AsEnumerable().Reverse())
        {
            if (seen.TryGetValue(entry.Source, out var duplicate))
            {
                if (duplicate.Date != entry.Date || duplicate.Reason != entry.Reason || duplicate.Change != entry.Change
                    || duplicate.Balance != entry.Balance || duplicate.Total != entry.Total
                    || !duplicate.ItemNames.SequenceEqual(entry.ItemNames) || !duplicate.Items.SequenceEqual(entry.Items))
                    throw new InvalidOperationException("Steam history: conflicting duplicate transaction.");
                continue;
            }
            seen.Add(entry.Source, entry);
            var items = entry.Reason == "转账" ? new List<SteamWalletItem> { new("", String.Join(" / ", entry.ItemNames), entry.Change) } : entry.Items;
            if (items.Count == 0 || items.Sum(item => item.Amount) != entry.Change)
                throw new InvalidOperationException("Steam history: item amounts do not match wallet movement.");
            var expected = items.Select(item => new Record
            {
                Account = account, date = entry.Date, postingDate = now, updateTime = now,
                t = CurrencyType.RMB, v = item.Amount, Reason = entry.Reason,
                Source = entry.Source + (item.Id.Length == 0 ? "" : "|item/" + item.Id) + "; "
                    + (item.Name.Length > 180 ? item.Name[..180] : item.Name)
            }).ToList();
            var old = known[entry.Source].ToList();
            if (old.Count > 0)
            {
                if (old.Count != expected.Count || expected.Any(record => !old.Any(saved => saved.Source == record.Source
                    && saved.v == record.v && saved.t == record.t && saved.date.Date == record.date)))
                    throw new InvalidOperationException("Steam history: previously imported wallet transaction changed.");
                continue;
            }
            balance += entry.Change;
            if (balance != entry.Balance)
                throw new InvalidOperationException("Steam history: wallet running balance mismatch.");
            records.AddRange(expected);
        }
        if (balance != ending) throw new InvalidOperationException("Steam history: wallet closing balance mismatch.");
        return records;
    }
}
