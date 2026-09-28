using System.Globalization;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace MyBook;

partial class WebUtil
{
    internal sealed record SteamExternalPurchase(string Id, DateTime Date, decimal Amount, string Items, string Payment, string CardSuffix);
    internal sealed record SteamBankPurchase(Record Record, string[] Cards);

    internal static bool HasSteamOrderCode(string source, string? id = null) =>
        Regex.IsMatch(source, @"(?:^|;\s*)code=SteamOrder-" + (id is null ? @"\d+" : Regex.Escape(id)) + @"(?=;|$)", RegexOptions.IgnoreCase);

    internal static bool IsSteamBankPurchase(Record record) => record.v < 0 && !record.Fake
        && !record.isInternal && !record.isRefundMatched && record.matchedRecordId is null && record.backup is null
        && record.Reason is not ("转账" or "换汇" or "手续费" or "税费" or "还款")
        && Regex.IsMatch(record.DestAccount + " " + record.Source,
            @"\b(?:steam(?:games|powered)?|valve)\b", RegexOptions.IgnoreCase);

    private static string SteamPaymentMethod(string text)
    {
        var match = Regex.Match(text, @"Visa|Master\s*Card|American Express|Amex|JCB|UnionPay|Alipay|WeChat|PayPal", RegexOptions.IgnoreCase);
        return Regex.Replace(match.Value, @"\s", "").ToUpperInvariant().Replace("AMERICANEXPRESS", "AMEX");
    }

    internal static List<SteamExternalPurchase> ParseSteamExternalPurchases(string html, DateTime since)
    {
        var document = new HtmlDocument(); document.LoadHtml(html);
        var result = new List<SteamExternalPurchase>();
        foreach (var row in document.DocumentNode.SelectNodes("//tr[contains(concat(' ',normalize-space(@class),' '),' wallet_table_row ')]") ?? new HtmlNodeCollection(document.DocumentNode))
        {
            HtmlNode? Cell(string name) => row.SelectSingleNode($"./td[contains(concat(' ',normalize-space(@class),' '),' {name} ')]");
            var type = SteamHistoryText(Cell("wht_type"));
            if (!type.Contains("Purchase", StringComparison.OrdinalIgnoreCase)) continue;
            if (!DateTime.TryParse(SteamHistoryText(Cell("wht_date")), CultureInfo.GetCultureInfo("en-US"), DateTimeStyles.None, out var date))
                throw new InvalidOperationException("Steam history: invalid purchase date.");
            if (date < since) continue;
            var payment = SteamPaymentMethod(type);
            if (payment.Length == 0) continue;
            var changeText = SteamHistoryText(Cell("wht_wallet_change"));
            var change = changeText.Length == 0 ? 0 : ReadSteamYuan(changeText);
            if (change > 0) continue; // Wallet top-ups are transfers, not game purchases.
            var amount = ReadSteamYuan(SteamHistoryText(Cell("wht_total"))) + change;
            if (amount <= 0) continue;
            var link = HtmlEntity.DeEntitize(row.GetAttributeValue("onclick", "") + " "
                + row.SelectSingleNode(".//a[@href]")?.GetAttributeValue("href", "")) ?? "";
            var id = Regex.Match(link, @"[?&](?:transid|transactionid)=(\d+)(?:[&'""\s]|$)");
            if (!id.Success) throw new InvalidOperationException("Steam history: purchase transaction identifier missing.");
            var names = Cell("wht_items")?.SelectNodes("./div[not(contains(concat(' ',normalize-space(@class),' '),' wth_item_refunded '))]")
                ?.Select(SteamHistoryText).ToArray() ?? [];
            if (names.Length == 0 || names.Any(String.IsNullOrWhiteSpace))
                throw new InvalidOperationException("Steam history: purchase item names missing.");
            var suffix = Regex.Match(type, @"(?:Visa|Master\s*Card|American Express|Amex|JCB|UnionPay)\s*[*xX•·]+\s*(\d{2,4})(?!\d)", RegexOptions.IgnoreCase);
            result.Add(new(id.Groups[1].Value, date, amount, String.Join(" / ", names), payment, suffix.Groups[1].Value));
        }
        return result;
    }

    internal static List<RecordSourceSupplement> BuildSteamPurchaseSupplements(
        IReadOnlyList<SteamExternalPurchase> purchases, IReadOnlyList<SteamBankPurchase> bankPurchases)
    {
        var orders = purchases.GroupBy(p => p.Id).Select(g => g.Distinct().ToList()).Where(g => g.Count == 1)
            .Select(g => g[0]).Where(p => !bankPurchases.Any(b => HasSteamOrderCode(b.Record.Source, p.Id))).ToList();
        var candidates = bankPurchases.Where(b => IsSteamBankPurchase(b.Record) && !HasSteamOrderCode(b.Record.Source)).ToList();
        var matches = orders.Select(p => (Purchase: p, Banks: candidates.Where(b =>
        {
            var record = b.Record;
            var amount = record.t == CurrencyType.RMB ? record.v : record.DescCurrency is { t: CurrencyType.RMB } original ? original.v : (decimal?)null;
            var method = SteamPaymentMethod(record.DestAccount + " " + record.Source);
            return amount == -p.Amount && Math.Abs((record.date.Date - p.Date.Date).TotalDays) <= 3
                && (method.Length == 0 || method == p.Payment)
                && (p.CardSuffix.Length == 0 || b.Cards.Any(card =>
                    Regex.IsMatch(card, @"^[0-9*xX•·]+$") && card.EndsWith(p.CardSuffix, StringComparison.Ordinal)));
        }).ToList())).ToList();
        var result = new List<RecordSourceSupplement>();
        foreach (var match in matches)
        {
            if (match.Banks.Count != 1 || matches.Count(m => m.Banks.Any(b => b.Record.Id == match.Banks[0].Record.Id)) != 1) continue;
            var purchase = match.Purchase;
            var code = "SteamOrder-" + purchase.Id;
            var append = FormattableString.Invariant($"Steam purchase; code={code}; date={purchase.Date:yyyy-MM-dd}; paid={purchase.Amount:0.00} RMB; payment={purchase.Payment}; {purchase.Items}");
            var record = match.Banks[0].Record;
            // Preserve the bank's full source, including identifiers used by its own importer.
            if (record.Source.Length + 2 + append.Length > 1024) continue;
            result.Add(new(record.Id, code, append));
        }
        return result;
    }
}
