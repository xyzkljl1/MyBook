using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net.Http;

namespace MyBook
{
    sealed class CryptoPriceUtil
    {
        private const string KrakenOhlcUrl = "https://api.kraken.com/0/public/OHLC";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(20);
        private static readonly HttpClient sharedHttpClient = CreateHttpClient();
        private readonly SemaphoreSlim cacheLock = new(1, 1);
        private readonly Dictionary<(string Asset, DateTime Date), CryptoDailyPrice> cache = new();

        public async Task<CryptoPriceSet> FetchDailyUsdPricesAsync(
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

                var prices = new Dictionary<(string Asset, DateTime Date), CryptoDailyPrice>();
                foreach (var asset in baseAssets)
                {
                    foreach (var date in EnumerateDates(firstDate, lastDate))
                    {
                        if (!cache.TryGetValue((asset, date), out var price))
                            throw new InvalidOperationException($"Missing Kraken daily close: asset={asset}, date={date:yyyy-MM-dd}.");
                        prices[(asset, date)] = price;
                    }
                }
                return new CryptoPriceSet(prices);
            }
            finally
            {
                cacheLock.Release();
            }
        }

        public static bool IsCryptoAsset(string asset)
        {
            return GetBaseAsset(asset) is "BTC" or "ETH" or "USDT";
        }

        public static string GetBaseAsset(string asset)
        {
            var normalized = asset.Trim().ToUpperInvariant();
            var suffixIndex = normalized.IndexOf('.', StringComparison.Ordinal);
            return suffixIndex < 0 ? normalized : normalized[..suffixIndex];
        }

        public static decimal CalculateMarketValue(decimal quantity, decimal unitPrice)
        {
            ValidateDecimal30_18(quantity, "crypto quantity");
            ValidateDecimal30_18(unitPrice, "crypto unit price");
            try
            {
                var value = Decimal.Round(checked(quantity * unitPrice), 2, MidpointRounding.AwayFromZero);
                ValidateDecimal30_18(value, "crypto market value");
                return value;
            }
            catch (OverflowException exception)
            {
                throw new InvalidOperationException(
                    $"Crypto market value exceeds decimal(30,18): quantity={quantity}, unitPrice={unitPrice}.",
                    exception);
            }
        }

        public static void ValidateDecimal30_18(decimal value, string name)
        {
            MySqlDecimalColumnTypes.ValidateCurrencyValue(value, name);
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

            var committed = new SortedDictionary<DateTime, CryptoDailyPrice>();
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
                ValidateDecimal30_18(close, $"Kraken {pair} close");
                if (!committed.TryAdd(candleDate, new CryptoDailyPrice(asset, candleDate, close, candleDate)))
                    throw new InvalidOperationException($"Kraken OHLC returned duplicate daily candle: pair={pair}, date={candleDate:yyyy-MM-dd}.");
            }

            CryptoDailyPrice? previous = null;
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
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 CryptoPriceUtil");
            return client;
        }
    }

    sealed class CryptoPriceSet
    {
        private readonly IReadOnlyDictionary<(string Asset, DateTime Date), CryptoDailyPrice> prices;

        public CryptoPriceSet(IReadOnlyDictionary<(string Asset, DateTime Date), CryptoDailyPrice> prices)
        {
            this.prices = prices;
        }

        public CryptoDailyPrice Get(string asset, DateTime date)
        {
            var baseAsset = CryptoPriceUtil.GetBaseAsset(asset);
            if (baseAsset == "USD")
                return new CryptoDailyPrice("USD", date.Date, 1m, date.Date);
            return prices.TryGetValue((baseAsset, date.Date), out var price)
                ? price
                : throw new InvalidOperationException($"Crypto daily close was not loaded: asset={baseAsset}, date={date:yyyy-MM-dd}.");
        }
    }

    sealed record CryptoDailyPrice(string Asset, DateTime Date, decimal CloseUsd, DateTime SourceCandleDate);
}
