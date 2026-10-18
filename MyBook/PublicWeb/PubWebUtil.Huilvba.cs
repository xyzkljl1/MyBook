using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook;

partial class PubWebUtil
{
    internal static readonly RateSource[] HuilvbaSources =
        [RateSource.HlbCcb, RateSource.HlbIcbc, RateSource.HlbCib, RateSource.HlbHfBank];

    private static DateOnly BeijingQuoteDate(DateTime localTime) => DateOnly.FromDateTime(
        new DateTimeOffset(DateTime.SpecifyKind(localTime, DateTimeKind.Local)).ToOffset(TimeSpan.FromHours(8)).DateTime);

    internal async Task FetchHuilvbaRates(RateSource source)
    {
        var through = BeijingQuoteDate(DateTime.Now).AddDays(-1);
        await FetchScheduledRatePairs(source, [CurrencyType.USD, CurrencyType.HKD, CurrencyType.GBP, CurrencyType.EUR],
            (currency, last) => BeijingQuoteDate(last) < through
                ? ReadHuilvbaRates(source, currency, BeijingQuoteDate(last).AddDays(1), through)
                : Task.FromResult(new List<RateHistory>())).ConfigureAwait(false);
    }

    internal async Task<List<RateHistory>> ReadHuilvbaRates(RateSource source, CurrencyType currency,
        DateOnly from, DateOnly through)
    {
        var bankCode = source switch
        {
            RateSource.HlbCcb => "ccb", RateSource.HlbIcbc => "icbc",
            RateSource.HlbCib => "cib", RateSource.HlbHfBank => "hfbank",
            _ => throw new ArgumentOutOfRangeException(nameof(source))
        };
        if (currency is not (CurrencyType.USD or CurrencyType.HKD or CurrencyType.GBP or CurrencyType.EUR))
            throw new ArgumentOutOfRangeException(nameof(currency));
        if (from > through) return [];
        if (through >= BeijingQuoteDate(DateTime.Now))
            throw new ArgumentOutOfRangeException(nameof(through), "Huilvba history requires completed Beijing dates.");
        var (client, route) = currency switch
        {
            CurrencyType.HKD => (httpClient, "pubweb_proxy"),
            CurrencyType.GBP => (secondaryHttpClient, "pubweb_proxy2"),
            _ => (directHttpClient, "direct")
        };
        var path = $"www.huilv8.com/waihuipaijia/{bankCode}/{currency.ToString().ToLowerInvariant()}";
        var request = $"Huilvba GET {path} ({from:yyyy-MM-dd} to {through:yyyy-MM-dd}; route={route})";
        try
        {
            using var response = await client.GetAsync($"https://{path}").ConfigureAwait(false);
            if (response.StatusCode == HttpStatusCode.TooManyRequests)
                throw new InvalidOperationException($"{request}: HTTP 429; rate limited.");
            if (response.StatusCode != HttpStatusCode.OK)
                throw new InvalidOperationException($"{request}: HTTP {(int)response.StatusCode}.");
            var html = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            try { return ParseHuilvbaHistory(html, source, bankCode, currency, from, through, DateTime.Now); }
            catch (InvalidOperationException e)
            { throw new InvalidOperationException($"{request}: HTTP 200; {e.Message}"); }
            catch (Exception e) when (e is JsonException or FormatException or OverflowException or ArgumentException)
            { throw new InvalidOperationException($"{request}: HTTP 200; invalid historical quote data ({e.GetType().Name})."); }
        }
        catch (HttpRequestException e)
        { throw new InvalidOperationException($"{request}: {e.HttpRequestError}; HTTP {e.StatusCode?.ToString() ?? "unavailable"}."); }
        catch (OperationCanceledException)
        { throw new InvalidOperationException($"{request}: timeout; HTTP unavailable."); }
    }

    internal static List<RateHistory> ParseHuilvbaHistory(string html, RateSource source, string bankCode,
        CurrencyType currency, DateOnly from, DateOnly through, DateTime fetchedAt)
    {
        // Decode the server-rendered data, without executing page scripts or using its current quote.
        var flight = string.Concat(Regex.Matches(html, @"self\.__next_f\.push\((\[.*?\])\)\s*;?\s*</script>", RegexOptions.Singleline)
            .SelectMany(m => JArray.Parse(m.Groups[1].Value).Where(t => t.Type == JTokenType.String).Select(t => (string?)t)));
        var objects = new List<JObject>();
        foreach (var line in flight.Split('\n'))
        {
            var colon = line.IndexOf(':');
            if (colon < 0 || colon + 1 == line.Length || line[colon + 1] is not ('[' or '{')) continue;
            using var reader = new JsonTextReader(new StringReader(line[(colon + 1)..]))
                { FloatParseHandling = FloatParseHandling.Decimal, DateParseHandling = DateParseHandling.None };
            var root = JToken.ReadFrom(reader);
            if (root is JContainer container) objects.AddRange(container.DescendantsAndSelf().OfType<JObject>());
        }
        if (!objects.Any(o => (string?)o["bank_code"] == bankCode && (string?)o["currency_code"] == currency.ToString()))
            throw new InvalidOperationException("Missing or unexpected bank/currency identity.");
        var histories = objects.Where(o => o["points"] is JArray).ToList();
        if (histories.Count != 1 || (bool?)histories[0]["unavailable"] != false)
            throw new InvalidOperationException("Historical quotes unavailable.");
        var points = new SortedDictionary<DateTime, (decimal? Buy, decimal? Sell)>();
        foreach (var row in (JArray)histories[0]["points"]!)
        {
            if (!DateTime.TryParseExact((string?)row["time"], "yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var time))
                throw new InvalidOperationException("Invalid historical sampling time.");
            var buy = (decimal?)row["buy_rate"];
            var sell = (decimal?)row["sell_rate"];
            if (buy is <= 0 || sell is <= 0 || !points.TryAdd(time, (buy, sell)))
                throw new InvalidOperationException("Invalid or duplicate historical quote.");
        }
        if (points.Count > 0 && from < DateOnly.FromDateTime(points.First().Key))
            throw new InvalidOperationException("Requested start precedes the available historical window.");

        // Use each completed Beijing day's last sample. Midnight remains a daily marker, not publication time.
        // The page quotes CNY per 100 foreign units; reverse rates store foreign units per CNY.
        return points.GroupBy(p => DateOnly.FromDateTime(p.Key)).Where(g => g.Key >= from && g.Key <= through)
            .Select(g => (Day: g.Key, Quote: g.Last().Value)).Where(p => p.Quote.Buy.HasValue || p.Quote.Sell.HasValue)
            .Select(p => new RateHistory
            {
                source = source, currency = currency, fetchedAt = fetchedAt,
                rateDate = new DateTimeOffset(p.Day.ToDateTime(TimeOnly.MinValue), TimeSpan.FromHours(8)).LocalDateTime,
                exchangeRateToRmb = p.Quote.Buy / 100m,
                exchangeRateFromRmb = p.Quote.Sell.HasValue ? decimal.Round(100m / p.Quote.Sell.Value, 18, MidpointRounding.AwayFromZero) : null
            }).ToList();
    }
}
