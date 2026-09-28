using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook
{
    // Google Finance source for US stocks, ETFs, and FX rates to CNY.
    partial class PubWebUtil
    {
        private async Task<List<RateHistory>> ReadGoogleCurrencyRates(CurrencyType currency, DateTime sinceUtc, DateTime throughExclusiveUtc)
        {
            async Task<List<RateHistory>> Read(bool reverse)
            {
                var pair = reverse ? $"CNY-{currency}" : $"{currency}-CNY";
                var html = await ReadGoogleFinancePage(currency, history: true, reverse).ConfigureAwait(false);
                try { return ParseGoogleFinanceHistory(html, currency, sinceUtc, throughExclusiveUtc, DateTime.Now, reverse); }
                catch (Exception e) when (e is JsonException or ArgumentException or InvalidOperationException or FormatException or OverflowException)
                { throw new InvalidOperationException($"Google Finance GET www.google.com/finance/quote/{pair}: HTTP 200; invalid daily history ({e.GetType().Name})."); }
            }
            var sides = await Task.WhenAll(Read(false), Read(true)).ConfigureAwait(false);
            return PairExchangeRates(sides[0], sides[1]);
        }

        private static DateTime GetGoogleHistoryStart(DateTime lastLocalTime, DateTime utcToday)
        {
            var since = DateTime.SpecifyKind(lastLocalTime, DateTimeKind.Local).ToUniversalTime().Date.AddDays(1);
            if (since > utcToday)
                throw new InvalidOperationException("Google history start is after the current UTC date.");
            if (utcToday > since.AddMonths(1))
                throw new InvalidOperationException($"Google history range {since:yyyy-MM-dd} to {utcToday:yyyy-MM-dd} exceeds one calendar month.");
            return since;
        }

        private static List<RateHistory> ParseGoogleFinanceHistory(string html, CurrencyType currency,
            DateTime sinceUtc, DateTime utcToday, DateTime fetchedAt, bool reverse = false)
        {
            var quotes = new Dictionary<DateTime, decimal>();
            var found = false;
            foreach (Match match in Regex.Matches(html, @"AF_initDataCallback\(\{key: '[^']+',.*?data:(?<data>.*?), sideChannel:", RegexOptions.Singleline))
            {
                using var reader = new JsonTextReader(new StringReader(match.Groups["data"].Value)) { FloatParseHandling = FloatParseHandling.Decimal };
                var data = JToken.ReadFrom(reader);
                if (data.SelectToken("[0][0]") is not JArray chart
                    || !chart.Descendants().OfType<JValue>().Any(v => v.Type == JTokenType.String && (string?)v == (reverse ? $"CNY / {currency}" : $"{currency} / CNY"))
                    || chart.SelectToken("[3][0][0]") is not JArray period || period.Count != 1 || (int?)period[0] != 1
                    || chart.SelectToken("[3][0][1]") is not JArray points) continue;
                found = true;
                foreach (var point in points)
                {
                    if (point[0] is not JArray time || time.Count < 8 || time[7] is not JArray offset
                        || offset.Any(v => v.Type != JTokenType.Null && (decimal)v != 0))
                        throw new InvalidOperationException("Historical quote must provide a zero UTC offset.");
                    var utc = new DateTime((int)time[0], (int)time[1], (int)time[2], (int?)time[3] ?? 0,
                        (int?)time[4] ?? 0, (int?)time[5] ?? 0, DateTimeKind.Utc).AddTicks(((long?)time[6] ?? 0) / 100);
                    // The last chart point can be a live quote for the unfinished UTC day.
                    if (utc.Date < sinceUtc || utc.Date >= utcToday) continue;
                    var rate = (decimal?)point[1]?[0];
                    if (rate is null or <= 0) throw new InvalidOperationException("Missing positive historical exchange rate.");
                    if (quotes.TryGetValue(utc, out var existing) && existing != rate.Value)
                        throw new InvalidOperationException("Conflicting historical exchange rates.");
                    quotes[utc] = rate.Value;
                }
            }
            if (!found) throw new InvalidOperationException("Daily currency history is missing.");
            if (quotes.Keys.GroupBy(t => t.Date).Any(g => g.Count() != 1))
                throw new InvalidOperationException("Multiple historical quotes for one UTC day.");
            return quotes.OrderBy(p => p.Key).Select(p => new RateHistory
            {
                source = RateSource.GoogleFinance, currency = currency, rateDate = p.Key.ToLocalTime(),
                fetchedAt = fetchedAt, exchangeRateToRmb = reverse ? null : p.Value,
                exchangeRateFromRmb = reverse ? p.Value : null
            }).ToList();
        }

        public async Task<Currency?> FetchCurrencyToRmb(CurrencyType currencyType)
        {
            if (currencyType == CurrencyType.RMB)
                return new Currency(1, CurrencyType.RMB);

            try
            {
                return await ReadGoogleFinanceExchangeRate(currencyType).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch currency exchange rate {currencyType}/CNY: {e.Message}");
            }

            return null;
        }

        private async Task<Currency> ReadGoogleFinanceExchangeRate(CurrencyType currency)
        {
            var html = await ReadGoogleFinancePage(currency, history: false).ConfigureAwait(false);
            var rate = ParseGoogleFinanceExchangeRate(html, currency.ToString());
            if (rate is null or <= 0)
                throw new InvalidOperationException($"Google Finance GET www.google.com/finance/quote/{currency}-CNY: HTTP 200; missing valid rate in HTML.");
            return new Currency(rate.Value, CurrencyType.RMB);
        }

        private async Task<string> ReadGoogleFinancePage(CurrencyType currency, bool history, bool reverse = false)
        {
            var pair = reverse ? $"CNY-{currency}" : $"{currency}-CNY";
            var request = $"Google Finance GET www.google.com/finance/quote/{pair}";
            try
            {
                using var response = await googleHttpClient.GetAsync($"https://www.google.com/finance/quote/{pair}?" + (history ? "window=1M" : "hl=en")).ConfigureAwait(false);
                if (response.StatusCode != HttpStatusCode.OK)
                    throw new InvalidOperationException($"{request}: HTTP {(int)response.StatusCode}.");
                return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (HttpRequestException e)
            { throw new InvalidOperationException($"{request}: {e.HttpRequestError}; HTTP {e.StatusCode?.ToString() ?? "unavailable"}."); }
            catch (OperationCanceledException)
            { throw new InvalidOperationException($"{request}: timeout; HTTP unavailable."); }
        }

        private static decimal? ParseGoogleFinanceExchangeRate(string html, string fromCurrency)
        {
            var match = Regex.Match(html, $@"""{Regex.Escape(fromCurrency)} / CNY""\s*,\s*3\s*,\s*null\s*,\s*\[(?<rate>\d+(?:\.\d+)?)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                return null;

            return decimal.TryParse(match.Groups["rate"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var rate) ? rate : null;
        }

        public async Task<decimal> FetchGoogleFinanceStock(string code, string exchange = "NASDAQ")
        {
            var symbol = code.Trim().ToUpperInvariant();
            var market = exchange.Trim().ToUpperInvariant();
            var url = $"https://www.google.com/finance/quote/{Uri.EscapeDataString(symbol)}:{market}?hl=en";
            try
            {
                var html = await HttpGetString(url, googleHttpClient).ConfigureAwait(false);
                if (String.IsNullOrWhiteSpace(html))
                    return -1;

                var price = ParseGoogleFinanceStockPrice(html, symbol, market);
                if (price is null)
                    return -1;

                Console.WriteLine($"{symbol}:{market}:{price.Value}");
                return price.Value;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {symbol}:{market}: {e}");
            }

            return -1;
        }

        private static decimal? ParseGoogleFinanceStockPrice(string html, string symbol, string exchange)
        {
            var entityMatch = Regex.Match(
                html,
                $@"\[""[^""]+"",\[""{Regex.Escape(symbol)}"",""{Regex.Escape(exchange)}""\],""[^""]+"",\d+,""[A-Z]{{3}}"",\[-?\d+(?:\.\d+)?",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!entityMatch.Success)
                return null;

            var symbolMarker = $@"""{symbol}:{exchange}""";
            var entityEnd = html.IndexOf(symbolMarker, entityMatch.Index, StringComparison.OrdinalIgnoreCase);
            var entityLength = entityEnd > entityMatch.Index ? entityEnd - entityMatch.Index : Math.Min(1200, html.Length - entityMatch.Index);
            var entity = html.Substring(entityMatch.Index, entityLength);
            var quoteMatches = Regex.Matches(entity, @"\[(?<price>-?\d+(?:\.\d+)?),-?\d+(?:\.\d+)?,-?\d+(?:\.\d+)?,\d+,\d+,\d+\]");
            if (quoteMatches.Count == 0)
                return null;

            var match = quoteMatches[quoteMatches.Count - 1];
            return decimal.TryParse(match.Groups["price"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) ? price : null;
        }
    }
}
