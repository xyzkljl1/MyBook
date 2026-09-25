using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Microsoft.VisualBasic.FileIO;

namespace MyBook;

partial class WebUtil
{
    private sealed record FirstTradeExportRow(DateTime Date, DateTime? Settlement, string Type, string Symbol,
        decimal Quantity, decimal Price, decimal Amount, string Description, string? Subaccount = null,
        string? FitId = null, decimal? Commission = null, decimal? Fees = null);

    private static List<FirstTradeTransaction> ReconcileFirstTradeSources(FirstTradeAccountCapture item)
    {
        var csv = ReadFirstTradeCsv(item.Csv);
        var (ofx, asOf) = ReadFirstTradeOfx(item);
        var result = new List<FirstTradeTransaction>();
        var occurrences = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var row in item.HistoryPages.SelectMany(p => p.GetProperty("items").EnumerateArray()))
        {
            var date = FirstTradeDate(FirstTradeText(row, "report_date"), "yyyy-MM-dd");
            if (date < item.HistoryFrom.Date || date > item.HistoryThrough.Date)
                throw new FirstTradeException("out-of-range history transaction");
            var type = FirstTradeText(row, "trans_str");
            var description = FirstTradeText(row, "description");
            var subaccount = FirstTradeText(row, "account_type");
            var symbol = FirstTradeText(row, "symbol", allowEmpty: true);
            var quantity = FirstTradeNumber(row, "quantity");
            var price = FirstTradeNumber(row, "trade_price");
            var amount = FirstTradeNumber(row, "amount");
            if (subaccount is not ("Cash" or "Margin")) throw new FirstTradeException("unsupported history subaccount");
            var tx = new FirstTradeTransaction("", date, type, description, subaccount, symbol, quantity, price, amount);
            var csvRow = Match(csv, tx, "CSV");
            var ofxRow = Match(ofx, tx, "OFX");
            // Exports can lag the live API. Settled rows must reconcile; security transfers are absent from OFX.
            if (date <= asOf && csvRow is null) throw new FirstTradeException($"CSV missing history transaction on {date:yyyy-MM-dd}");
            if (date <= asOf && ofxRow is null && !(type == "OTHER" && quantity != 0))
                throw new FirstTradeException($"OFX missing history transaction on {date:yyyy-MM-dd}");
            if (csvRow?.Settlement is DateTime csvSettlement && ofxRow?.Settlement is DateTime ofxSettlement && csvSettlement != ofxSettlement)
                throw new FirstTradeException("CSV/OFX settlement dates disagree");
            if (csvRow?.Commission is decimal c && ofxRow?.Commission is decimal oc && c != oc
                || csvRow?.Fees is decimal f && ofxRow?.Fees is decimal of && f != of)
                throw new FirstTradeException("CSV/OFX charge details disagree");
            var commission = csvRow?.Commission ?? ofxRow?.Commission ?? 0;
            var fees = csvRow?.Fees ?? ofxRow?.Fees;
            var assumed = false;
            if (commission < 0 || fees < 0 || Decimal.Round(commission, 2) != commission
                || fees.HasValue && Decimal.Round(fees.Value, 2) != fees.Value)
                throw new FirstTradeException("invalid commission or fee precision");
            if (type is "BOUGHT" or "SOLD")
            {
                var gross = Decimal.Round(Math.Abs(quantity) * price, 2, MidpointRounding.AwayFromZero);
                if (fees is null)
                {
                    fees = type == "SOLD" ? FirstTradeSecFee(date, gross) : 0;
                    assumed = type == "SOLD";
                }
                FirstTradeEqual((type == "BOUGHT" ? -gross : gross) - commission - fees.Value, amount,
                    $"trade amount after {(assumed ? "assumed SEC" : "reported")} charges on {date:yyyy-MM-dd}");
            }
            else if (commission != 0 || fees.GetValueOrDefault() != 0)
                throw new FirstTradeException("non-trade has unsupported embedded charges");
            // Description is mutable. Decimal strings avoid JSON scale differences changing identity.
            var hash = FirstTradeHash(JsonSerializer.Serialize(new { date, type, subaccount, symbol,
                quantity = Math.Abs(quantity).ToString("G29", CultureInfo.InvariantCulture),
                price = price.ToString("G29", CultureInfo.InvariantCulture), amount = amount.ToString("G29", CultureInfo.InvariantCulture) }));
            occurrences.TryGetValue(hash, out var occurrence);
            occurrences[hash] = ++occurrence;
            var identity = hash + ":" + occurrence;
            result.Add(tx with { Key = ofxRow?.FitId is string id ? "ofx:" + id : "v2:" + identity,
                Identity = identity, SettlementDate = csvRow?.Settlement ?? ofxRow?.Settlement,
                Commission = commission, Fees = fees ?? 0, AssumedFee = assumed });
        }
        if (csv.Any(InRange) || ofx.Any(InRange)) throw new FirstTradeException("export contains transactions missing from API history");
        return result;

        bool InRange(FirstTradeExportRow row) => row.Date >= item.HistoryFrom.Date && row.Date <= item.HistoryThrough.Date;

        static FirstTradeExportRow? Match(List<FirstTradeExportRow> rows, FirstTradeTransaction tx, string source)
        {
            // CSV labels cash deposits OTHER; the API supplies DEPOSIT and OFX confirms the credit.
            var candidates = rows.Where(r => r.Date == tx.Date
                && (r.Type == tx.Type || source == "CSV" && r.Type == "OTHER" && tx.Type == "DEPOSIT") && r.Amount == tx.Amount
                && Math.Abs(r.Quantity) == Math.Abs(tx.Quantity) && r.Price == tx.Price
                && (r.Quantity == 0 || r.Symbol == tx.Symbol)
                && (r.Subaccount is null || r.Subaccount.Equals(tx.Subaccount, StringComparison.OrdinalIgnoreCase))).ToList();
            if (candidates.Count > 1)
                candidates = candidates.Where(r => Clean(r.Description) == Clean(tx.Description)).ToList();
            if (candidates.Count > 1 && candidates.Distinct().Count() == 1) candidates = [candidates[0]];
            if (candidates.Count > 1) throw new FirstTradeException($"ambiguous {source} transaction match");
            var found = candidates.SingleOrDefault();
            if (found is not null) rows.Remove(found);
            return found;
        }
        static string Clean(string value) => Regex.Replace(value.ToUpperInvariant(), @"\s+", " ").Trim();
    }

    internal static decimal FirstTradeSecFee(DateTime date, decimal gross)
    {
        // User-authorized assumption: round each sale's SEC fee to cents, midpoint away from zero.
        // Firstrade published USD 20.60 per million effective 2026-04-06; update on a new fee advisory.
        // https://www.firstrade.com/trading/pricing
        if (date < new DateTime(2026, 4, 6) || gross <= 0)
            throw new FirstTradeException("SEC fee rate is not verified for this trade date or amount");
        return Decimal.Round(gross * 0.00002060m, 2, MidpointRounding.AwayFromZero);
    }

    private static bool IsFirstTradeSubaccountTransfer(FirstTradeTransaction tx) => tx.Type == "OTHER" && tx.Quantity != 0
        && ((tx.Subaccount == "Cash" && tx.Description.EndsWith("TFR to Type 2", StringComparison.Ordinal))
            || (tx.Subaccount == "Margin" && tx.Description.EndsWith("TFR from Type 1", StringComparison.Ordinal)));

    private static HashSet<string> MatchFirstTradePrevious(List<FirstTradeTransaction> transactions, List<Record> previous,
        DateTime from, DateTime through)
    {
        var known = new HashSet<string>(StringComparer.Ordinal);
        foreach (var group in previous.Where(r => r.Source.StartsWith("FirstTrade transaction|", StringComparison.Ordinal))
            .GroupBy(r => r.Source.Split('|')[1]))
        {
            var record = group.FirstOrDefault(r => r.Source.Contains("|cash|", StringComparison.Ordinal)) ?? group.First();
            var identity = Regex.Match(record.Source, @"\|identity=([A-F0-9]+:\d+)").Groups[1].Value;
            var matches = transactions.Where(tx => tx.Key == group.Key || (identity.Length > 0 && tx.Identity == identity)
                || MatchesLegacy(tx, record, group.Key)).ToList();
            if (matches.Count > 1) throw new FirstTradeException("ambiguous previously imported transaction");
            if (matches.Count == 1)
            {
                var tx = matches[0];
                if (identity.Length > 0 && (identity != tx.Identity
                    || group.Where(r => r.Source.Contains("|fee|", StringComparison.Ordinal)).Sum(r => r.v) != -tx.Fees
                    || group.Where(r => r.Source.Contains("|commission|", StringComparison.Ordinal)).Sum(r => r.v) != -tx.Commission))
                    throw new FirstTradeException($"previously imported transaction details revised on {tx.Date:yyyy-MM-dd}");
                if (!known.Add(matches[0].Key)) throw new FirstTradeException("multiple previous transactions match one incoming transaction");
            }
            else if (record.date.Date >= from.Date && record.date.Date <= through.Date)
                throw new FirstTradeException($"previously imported transaction removed or revised on {record.date:yyyy-MM-dd}");
        }
        return known;

        static bool MatchesLegacy(FirstTradeTransaction tx, Record record, string key)
        {
            if (key.StartsWith("v2:") || key.StartsWith("ofx:") || record.date.Date != tx.Date || record.v != tx.Amount) return false;
            var prefix = "FirstTrade transaction|" + key + "|cash|" + tx.Type + " " + tx.Subaccount + " ";
            if (!record.Source.StartsWith(prefix, StringComparison.Ordinal)) return false;
            var date = tx.Date; var type = tx.Type; var description = record.Source[prefix.Length..];
            var subaccount = tx.Subaccount; var symbol = tx.Symbol; var quantity = tx.Quantity; var price = tx.Price; var amount = tx.Amount;
            var hash = FirstTradeHash(JsonSerializer.Serialize(new { date, type, description, subaccount, symbol, quantity, price, amount }));
            return key == hash + ":" + tx.Identity.Split(':')[1];
        }
    }

    private static List<FirstTradeExportRow> ReadFirstTradeCsv(string text)
    {
        using var reader = new TextFieldParser(new StringReader(text)) { HasFieldsEnclosedInQuotes = true, TrimWhiteSpace = true };
        reader.SetDelimiters(",");
        var header = reader.ReadFields();
        string[] expected = ["Symbol", "Quantity", "Price", "Action", "Description", "TradeDate", "SettledDate", "Interest", "Amount", "Commission", "Fee", "CUSIP", "RecordType"];
        if (header is null || !header.SequenceEqual(expected)) throw new FirstTradeException("invalid CSV export header; web login may be required");
        var result = new List<FirstTradeExportRow>();
        while (!reader.EndOfData)
        {
            var fields = reader.ReadFields()!;
            if (fields.Length != expected.Length || FirstTradeDecimal(fields[7]) != 0 || fields[12] is not ("Trade" or "Financial"))
                throw new FirstTradeException("unsupported CSV row or accrued interest");
            var type = fields[3].ToUpperInvariant() switch { "BUY" => "BOUGHT", "SELL" => "SOLD", var value => value };
            result.Add(new(FirstTradeDate(fields[5], "yyyy-MM-dd"), FirstTradeDate(fields[6], "yyyy-MM-dd"),
                type, fields[0], FirstTradeDecimal(fields[1]), fields[2].Length == 0 ? 0 : FirstTradeDecimal(fields[2]),
                FirstTradeDecimal(fields[8]), fields[4], Commission: FirstTradeDecimal(fields[9]), Fees: FirstTradeDecimal(fields[10])));
        }
        return result;
    }

    private static (List<FirstTradeExportRow> Rows, DateTime AsOf) ReadFirstTradeOfx(FirstTradeAccountCapture item)
    {
        var start = item.Ofx.IndexOf("<OFX>", StringComparison.Ordinal);
        if (start < 0) throw new FirstTradeException("invalid OFX export; web login may be required");
        using var reader = XmlReader.Create(new StringReader(item.Ofx[start..]), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var root = XDocument.Load(reader).Root!;
        if (root.Descendants("STATUS").Any(s => s.Element("CODE")?.Value != "0")) throw new FirstTradeException("OFX reports failure status");
        var statement = root.Descendants("INVSTMTRS").SingleOrDefault() ?? throw new FirstTradeException("OFX investment statement missing");
        if (statement.Element("INVACCTFROM")?.Element("ACCTID")?.Value != item.Account || statement.Element("CURDEF")?.Value != "USD")
            throw new FirstTradeException("OFX account or currency mismatch");
        var list = statement.Element("INVTRANLIST") ?? throw new FirstTradeException("OFX transactions missing");
        var asOf = Date(statement.Element("DTASOF")?.Value);
        if (Date(list.Element("DTSTART")?.Value) > item.HistoryFrom.Date || Date(list.Element("DTEND")?.Value) < item.HistoryThrough.Date
            || asOf > item.HistoryThrough.Date)
            throw new FirstTradeException("OFX date range does not cover requested range");
        var symbols = root.Descendants("SECINFO").ToDictionary(s => s.Element("SECID")!.Element("UNIQUEID")!.Value,
            s => s.Element("TICKER")?.Value ?? "", StringComparer.Ordinal);
        var result = new List<FirstTradeExportRow>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in list.Elements().Where(e => e.Name.LocalName is not ("DTSTART" or "DTEND")))
        {
            string Text(string name) => row.Descendants(name).SingleOrDefault()?.Value ?? "";
            decimal Number(string name) => Text(name) is { Length: > 0 } value ? FirstTradeDecimal(value) : 0;
            var id = Text("FITID");
            if (id.Length == 0 || id.Length > 180 || id.Contains('|') || !ids.Add(id)) throw new FirstTradeException("invalid or duplicate OFX transaction ID");
            var description = Text("MEMO");
            var type = row.Name.LocalName switch
            {
                "BUYSTOCK" when Text("BUYTYPE") == "BUY" => "BOUGHT",
                "SELLSTOCK" when Text("SELLTYPE") == "SELL" => "SOLD",
                "INCOME" when Text("INCOMETYPE") == "DIV" => "DIVIDEND",
                "INCOME" when Text("INCOMETYPE") == "INTEREST" || description.StartsWith("FULLYPAID LENDING REBATE", StringComparison.Ordinal) => "INTEREST",
                "INVBANKTRAN" when Text("NAME") is "XFER CASH TO MARGIN" or "XFER MARGIN TO CASH" => "OTHER",
                "INVBANKTRAN" when Text("TRNTYPE") is "DEP" or "CREDIT" => "DEPOSIT",
                "INVBANKTRAN" when Text("TRNTYPE") == "FEE" => "FEE",
                _ => throw new FirstTradeException("unsupported OFX transaction type")
            };
            var trade = type is "BOUGHT" or "SOLD";
            if (trade && (Text("COMMISSION").Length == 0 || Text("FEES").Length == 0)) throw new FirstTradeException("OFX trade charges missing");
            result.Add(new(Date(Text("DTTRADE").Length > 0 ? Text("DTTRADE") : Text("DTPOSTED")),
                Text("DTSETTLE").Length > 0 ? Date(Text("DTSETTLE")) : null, type,
                symbols.GetValueOrDefault(Text("UNIQUEID"), ""), Number("UNITS"), Number("UNITPRICE"),
                Number(row.Name.LocalName == "INVBANKTRAN" ? "TRNAMT" : "TOTAL"),
                description.Length > 0 ? description : Text("NAME"), Text("SUBACCTFUND"), id,
                trade ? Number("COMMISSION") : null, trade ? Number("FEES") : null));
        }
        return (result, asOf);
        static DateTime Date(string? value) => value?.Length >= 8 ? FirstTradeDate(value[..8], "yyyyMMdd")
            : throw new FirstTradeException("missing OFX date");
    }

    private static DateTime FirstTradeDate(string text, string format) => DateTime.TryParseExact(text, format,
        CultureInfo.InvariantCulture, DateTimeStyles.None, out var value) ? value : throw new FirstTradeException("invalid source date");
    private static decimal FirstTradeDecimal(string text)
    {
        if (!Decimal.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var value))
            throw new FirstTradeException("invalid export number");
        MySqlDecimalColumnTypes.ValidateCurrencyValue(value, "FirstTrade export");
        return value;
    }
    private static string FirstTradeHash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
}
