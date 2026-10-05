using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net.Http;

namespace MyBook
{
    sealed class KrakenPubUtil
    {
        private const string KrakenOhlcUrl = "https://api.kraken.com/0/public/OHLC";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(20);
        private static readonly HttpClient sharedHttpClient = CreateHttpClient();
        private readonly SemaphoreSlim cacheLock = new(1, 1);
        private readonly Dictionary<(string Asset, DateTime Date), KrakenDailyPrice> cache = new();

        internal async Task<List<MarketPrice>> FetchLatestUsdPricesAsync(IEnumerable<string> assets)
        {
            var symbols = assets.Select(GetBaseAsset).Distinct(StringComparer.Ordinal).ToList();
            if (symbols.Count == 0) return [];
            foreach (var symbol in symbols) ValidateSupportedAsset(symbol);
            const string request = "Kraken GET api.kraken.com/0/public/Ticker";
            try
            {
                var pairs = String.Join(",", symbols.Select(symbol => symbol + "USD"));
                using var response = await sharedHttpClient.GetAsync(
                    $"https://api.kraken.com/0/public/Ticker?pair={Uri.EscapeDataString(pairs)}&assetVersion=1").ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException($"{request}: HTTP {(int)response.StatusCode}.");
                var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                return ParseTickerPrices(text, symbols, DateTimeOffset.Now);
            }
            catch (HttpRequestException e) { throw new InvalidOperationException($"{request}: transport {e.HttpRequestError}."); }
            catch (TaskCanceledException) { throw new InvalidOperationException($"{request}: timeout."); }
            catch (Exception e) when (e is Newtonsoft.Json.JsonException or FormatException or OverflowException)
            { throw new InvalidOperationException($"{request}: HTTP 200; invalid ticker response ({e.GetType().Name})."); }
        }

        internal static List<MarketPrice> ParseTickerPrices(string text, IReadOnlyList<string> symbols, DateTimeOffset fetchedAt)
        {
            const string request = "Kraken GET api.kraken.com/0/public/Ticker: HTTP 200";
            var json = JObject.Parse(text);
            if (json["error"] is not JArray errors || errors.Count != 0 || json["result"] is not JObject result)
                throw new InvalidOperationException($"{request}; ticker API error or missing result.");
            return symbols.Select(symbol =>
            {
                var value = result[symbol + "/USD"]?["c"]?[0]?.Value<string>();
                if (!Decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var price) || price <= 0)
                    throw new InvalidOperationException($"{request}; missing positive {symbol}/USD last trade price.");
                return new MarketPrice(symbol, HoldingType.Crypto, price, CurrencyType.USD, fetchedAt);
            }).ToList();
        }

        public async Task<KrakenDailyPriceSet> FetchDailyUsdPricesAsync(
            IEnumerable<string> assets,
            DateTime firstDate,
            DateTime lastDate,
            CancellationToken cancellationToken = default)
        {
            firstDate = firstDate.Date;
            lastDate = lastDate.Date;
            if (firstDate > lastDate)
                throw new ArgumentException("Crypto price start date must not be after end date.", nameof(firstDate));

            var baseAssets = assets
                .Select(GetBaseAsset)
                .Where(asset => asset != "USD")
                .Distinct(StringComparer.Ordinal)
                .OrderBy(asset => asset, StringComparer.Ordinal)
                .ToList();
            foreach (var asset in baseAssets)
                ValidateSupportedAsset(asset);

            await cacheLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                foreach (var asset in baseAssets)
                {
                    var missing = EnumerateDates(firstDate, lastDate)
                        .Any(date => !cache.ContainsKey((asset, date)));
                    if (missing)
                        await FetchAndCacheAssetAsync(asset, firstDate, lastDate, cancellationToken).ConfigureAwait(false);
                }

                var prices = new Dictionary<(string Asset, DateTime Date), KrakenDailyPrice>();
                foreach (var asset in baseAssets)
                {
                    foreach (var date in EnumerateDates(firstDate, lastDate))
                    {
                        if (!cache.TryGetValue((asset, date), out var price))
                            throw new InvalidOperationException($"Missing Kraken daily close: asset={asset}, date={date:yyyy-MM-dd}.");
                        prices[(asset, date)] = price;
                    }
                }
                return new KrakenDailyPriceSet(prices);
            }
            finally
            {
                cacheLock.Release();
            }
        }

        public static bool IsCryptoAsset(string asset)
        {
            return GetBaseAsset(asset) is "BTC" or "ETH" or "USDT" or "BABY";
        }

        public static string GetBaseAsset(string asset)
        {
            var normalized = asset.Trim().ToUpperInvariant();
            var suffixIndex = normalized.IndexOf('.', StringComparison.Ordinal);
            return suffixIndex < 0 ? normalized : normalized[..suffixIndex];
        }

        private async Task FetchAndCacheAssetAsync(
            string asset,
            DateTime firstDate,
            DateTime lastDate,
            CancellationToken cancellationToken)
        {
            var requestStart = firstDate.AddDays(-7);
            var since = new DateTimeOffset(DateTime.SpecifyKind(requestStart, DateTimeKind.Utc)).ToUnixTimeSeconds();
            var pair = asset + "USD";
            var url = $"{KrakenOhlcUrl}?pair={Uri.EscapeDataString(pair)}&assetVersion=1&interval=1440&since={since.ToString(CultureInfo.InvariantCulture)}";
            using var response = await sharedHttpClient.GetAsync(url, cancellationToken).ConfigureAwait(false);
            var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException($"Kraken OHLC request failed for {pair}: {(int)response.StatusCode} {response.ReasonPhrase}.");

            var json = JObject.Parse(text);
            var errors = json["error"] as JArray;
            if (errors is { Count: > 0 })
                throw new InvalidOperationException($"Kraken OHLC request failed for {pair}: {String.Join("; ", errors.Values<string>())}.");
            var result = json["result"] as JObject
                ?? throw new InvalidOperationException($"Kraken OHLC response has no result for {pair}.");
            var rows = result.Properties()
                .Where(property => property.Name != "last")
                .Select(property => property.Value)
                .OfType<JArray>()
                .SingleOrDefault()
                ?? throw new InvalidOperationException($"Kraken OHLC response has no candle array for {pair}.");

            var committed = new SortedDictionary<DateTime, KrakenDailyPrice>();
            var currentUtcDate = DateTime.UtcNow.Date;
            foreach (var row in rows.OfType<JArray>())
            {
                if (row.Count < 5)
                    throw new InvalidOperationException($"Kraken OHLC row is incomplete for {pair}.");
                var unixTime = row[0]!.Value<long>();
                var candleDate = DateTimeOffset.FromUnixTimeSeconds(unixTime).UtcDateTime.Date;
                if (candleDate >= currentUtcDate)
                    continue;
                var close = Decimal.Parse(row[4]!.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture);
                if (close <= 0)
                    throw new InvalidOperationException($"Kraken OHLC close must be positive: pair={pair}, date={candleDate:yyyy-MM-dd}, close={close}.");
                MySqlDecimalColumnTypes.ValidateCurrencyValue(close, $"Kraken {pair} close");
                if (!committed.TryAdd(candleDate, new KrakenDailyPrice(asset, candleDate, close, candleDate)))
                    throw new InvalidOperationException($"Kraken OHLC returned duplicate daily candle: pair={pair}, date={candleDate:yyyy-MM-dd}.");
            }

            KrakenDailyPrice? previous = null;
            foreach (var date in EnumerateDates(requestStart, lastDate))
            {
                if (committed.TryGetValue(date, out var candle))
                    previous = candle;
                if (date < firstDate)
                    continue;
                if (previous is null)
                    throw new InvalidOperationException($"Kraken OHLC has no close at or before {date:yyyy-MM-dd} for {pair}.");
                cache[(asset, date)] = previous with { Date = date };
            }
        }

        private static IEnumerable<DateTime> EnumerateDates(DateTime firstDate, DateTime lastDate)
        {
            for (var date = firstDate.Date; date <= lastDate.Date; date = date.AddDays(1))
                yield return date;
        }

        private static void ValidateSupportedAsset(string asset)
        {
            if (!IsCryptoAsset(asset))
                throw new InvalidOperationException($"Unsupported crypto asset: {asset}.");
        }

        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient { Timeout = RequestTimeout };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 KrakenPubUtil");
            return client;
        }
    }

    sealed class KrakenDailyPriceSet
    {
        private readonly IReadOnlyDictionary<(string Asset, DateTime Date), KrakenDailyPrice> prices;

        public KrakenDailyPriceSet(IReadOnlyDictionary<(string Asset, DateTime Date), KrakenDailyPrice> prices)
        {
            this.prices = prices;
        }

        public KrakenDailyPrice Get(string asset, DateTime date)
        {
            var baseAsset = KrakenPubUtil.GetBaseAsset(asset);
            if (baseAsset == "USD")
                return new KrakenDailyPrice("USD", date.Date, 1m, date.Date);
            return prices.TryGetValue((baseAsset, date.Date), out var price)
                ? price
                : throw new InvalidOperationException($"Crypto daily close was not loaded: asset={baseAsset}, date={date:yyyy-MM-dd}.");
        }
    }

    sealed record KrakenDailyPrice(string Asset, DateTime Date, decimal CloseUsd, DateTime SourceCandleDate);
}
