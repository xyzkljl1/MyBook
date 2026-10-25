using System.IO;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook;

partial class PubWebUtil
{
    private async Task<decimal> FetchYahooFinanceStock(string code)
    {
        var symbol = Uri.EscapeDataString(code.Trim().ToUpperInvariant());
        var path = $"query1.finance.yahoo.com/v8/finance/chart/{symbol}";
        var request = $"Yahoo Finance GET {path}";
        try
        {
            // Yahoo uses the existing pubweb_google_proxy route, also used for Google FX history.
            using var response = await googleHttpClient.GetAsync($"https://{path}?interval=1m&range=1d&includePrePost=true").ConfigureAwait(false);
            if (response.StatusCode != HttpStatusCode.OK)
                throw new InvalidOperationException($"{request}: HTTP {(int)response.StatusCode}.");
            var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            try { return ParseYahooFinanceStockPrice(json, code.Trim().ToUpperInvariant()); }
            catch (Exception e) when (e is JsonException or InvalidOperationException or ArgumentException or FormatException or OverflowException)
            { throw new InvalidOperationException($"{request}: HTTP 200; invalid stock quote ({e.GetType().Name})."); }
        }
        catch (HttpRequestException e) { throw new InvalidOperationException($"{request}: transport {e.HttpRequestError}."); }
        catch (OperationCanceledException) { throw new InvalidOperationException($"{request}: timeout."); }
    }

    private static decimal ParseYahooFinanceStockPrice(string json, string symbol)
    {
        using var reader = new JsonTextReader(new StringReader(json)) { FloatParseHandling = FloatParseHandling.Decimal };
        var chart = JObject.Load(reader)["chart"];
        if (chart?["error"] is not { Type: JTokenType.Null }
            || chart["result"] is not JArray { Count: 1 } results
            || results[0]["meta"] is not JObject meta
            || (string?)meta["symbol"] != symbol || (string?)meta["currency"] != "USD")
            throw new InvalidOperationException("Missing or mismatched Yahoo quote.");

        var price = (decimal?)meta["regularMarketPrice"];
        var time = (long?)meta["regularMarketTime"];
        if (price is not > 0 || time is not > 0)
            throw new InvalidOperationException("Missing positive Yahoo price or quote timestamp.");
        _ = DateTimeOffset.FromUnixTimeSeconds(time.Value);

        var result = results[0];
        if (result["timestamp"] is JArray times && result.SelectToken("indicators.quote[0].close") is JArray closes)
        {
            if (times.Count != closes.Count)
                throw new InvalidOperationException("Yahoo price/timestamp count mismatch.");
            for (var i = 0; i < times.Count; i++)
            {
                if (closes[i].Type == JTokenType.Null) continue;
                var close = (decimal?)closes[i];
                var timestamp = (long?)times[i];
                if (close is not > 0 || timestamp is not > 0)
                    throw new InvalidOperationException("Invalid Yahoo chart quote.");
                _ = DateTimeOffset.FromUnixTimeSeconds(timestamp.Value);
                // Include newer pre/post-market points; stale chart points must not replace a newer regular quote.
                if (timestamp > time) { price = close; time = timestamp; }
            }
        }
        return price.Value;
    }
}
