using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // Shared helpers for public web market data fetches.
    partial class PubWebUtil : IDisposable
    {
        private static readonly TimeSpan HttpRequestTimeout = TimeSpan.FromSeconds(20);
        private readonly DatabaseUtil? database;
        private readonly HttpClient httpClient;
        private readonly HttpClient googleHttpClient;
        private readonly HttpClient directHttpClient;
        private readonly HttpClient secondaryHttpClient;

        public PubWebUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.database = database;
            httpClient = CreateHttpClient(ParsePubWebProxyConfig(config["pubweb_proxy"]));
            googleHttpClient = CreateHttpClient(ParsePubWebProxyConfig(config["pubweb_google_proxy"], "pubweb_google_proxy"));
            directHttpClient = CreateHttpClient(null);
            secondaryHttpClient = CreateHttpClient(ParsePubWebProxyConfig(config["pubweb_proxy2"], "pubweb_proxy2"));
        }

        public Task<MarketPrice?> Fetch(Holding holding)
        {
            return Fetch(holding.code, holding.holdingType);
        }

        public async Task<MarketPrice?> Fetch(string code, HoldingType holdingType)
        {
            Currency? ret = null;
            switch (holdingType)
            {
                case HoldingType.NASDAQ:
                case HoldingType.ARCA:
                    ret = new Currency(await FetchYahooFinanceStock(code).ConfigureAwait(false), CurrencyType.USD);
                    break;
                case HoldingType.UST:
                    Console.WriteLine("skip UST price: fetcher is not configured");
                    break;
                case HoldingType.SHANGHAI:
                    ret = new Currency(await FetchShanghaiStock(code).ConfigureAwait(false), CurrencyType.RMB);
                    break;
                case HoldingType.CNFUND:
                    ret = new Currency(await FetchCNFund(code).ConfigureAwait(false), CurrencyType.RMB);
                    break;
                case HoldingType.Cash:
                    break;
                case HoldingType.Accrued:
                    break;
                case HoldingType.Crypto:
                    break;
            }

            return ret is null || ret.v <= 0 ? null
                : new MarketPrice(code, holdingType, ret.v, ret.t, DateTimeOffset.Now);
        }

        internal async Task FetchScheduledExchangeRates()
        {
            var db = database ?? throw new InvalidOperationException("Scheduled rates require a database.");
            await FetchScheduledRatePairs(RateSource.GoogleFinance,
                Enum.GetValues<CurrencyType>().Where(c => c != CurrencyType.RMB),
                (currency, last) => ReadGoogleCurrencyRates(currency, GetGoogleHistoryStart(last, DateTime.UtcNow.Date), DateTime.UtcNow.Date)).ConfigureAwait(false);
            db.CacheExchangeLosses();
        }

        private async Task FetchScheduledRatePairs(RateSource source, IEnumerable<CurrencyType> currencies,
            Func<CurrencyType, DateTime, Task<List<RateHistory>>> fetch)
        {
            var db = database ?? throw new InvalidOperationException("Scheduled rates require a database.");
            var progress = db.GetLatestCompleteExchangeRateTimes(source);
            foreach (var currency in currencies)
            {
                DateTime? latest = progress.TryGetValue(currency, out var last) ? last : null;
                await ImportSchedule.RunAsync($"{source}/{currency}", 1, 7, () => latest, async since =>
                {
                    var batch = await fetch(currency, since).ConfigureAwait(false);
                    db.SaveRateHistory(batch);
                    latest = batch.Where(r => r.exchangeRateToRmb is > 0 && r.exchangeRateFromRmb is > 0)
                        .Max(r => (DateTime?)r.rateDate) ?? latest;
                }).ConfigureAwait(false);
            }
        }

        private static List<RateHistory> PairExchangeRates(List<RateHistory> forward, List<RateHistory> reverse)
        {
            DateOnly QuoteDate(RateHistory rate) => rate.source == RateSource.GoogleFinance
                ? DateOnly.FromDateTime(DateTime.SpecifyKind(rate.rateDate, DateTimeKind.Local).ToUniversalTime())
                : BeijingQuoteDate(rate.rateDate);
            var reverseByDate = reverse.ToDictionary(QuoteDate);
            var result = new List<RateHistory>();
            foreach (var rate in forward)
            {
                if (reverseByDate.Remove(QuoteDate(rate), out var other))
                {
                    if (rate.rateDate != other.rateDate || rate.currency != other.currency || rate.source != other.source)
                        throw new InvalidOperationException($"Exchange rate pair timestamp mismatch: {rate.source}/{rate.currency} {rate.rateDate:O} / {other.rateDate:O}.");
                    rate.exchangeRateFromRmb = other.exchangeRateFromRmb;
                    rate.fetchedAt = rate.fetchedAt > other.fetchedAt ? rate.fetchedAt : other.fetchedAt;
                }
                result.Add(rate);
            }
            return result.Concat(reverseByDate.Values).OrderBy(r => r.rateDate).ToList();
        }

        internal async Task FetchMarketPricesAsync(IReadOnlyList<(string Code, HoldingType HoldingType)> targets,
            KrakenPubUtil kraken, Action<MarketPrice> onPrice, bool scheduled = false)
        {
            var startedAt = DateTimeOffset.Now;
            var elapsed = Stopwatch.StartNew();
            var results = new Dictionary<(string Code, HoldingType Type), MarketPrice>();
            var errors = new List<Exception>();
            void Report(MarketPrice quote)
            {
                results[(quote.Code, quote.HoldingType)] = quote;
                onPrice(quote);
            }
            foreach (var target in targets.Where(target => target.HoldingType != HoldingType.Crypto))
            {
                try
                {
                    var quote = await FetchWithRetry(async () =>
                        await Fetch(target.Code, target.HoldingType).ConfigureAwait(false)
                        ?? throw new InvalidOperationException($"Missing market price: {target.Code} ({target.HoldingType}).")).ConfigureAwait(false);
                    Report(quote);
                }
                catch (Exception e) { errors.Add(e); }
            }
            try
            {
                var crypto = await FetchWithRetry(() => kraken.FetchLatestUsdPricesAsync(targets
                    .Where(target => target.HoldingType == HoldingType.Crypto).Select(target => target.Code))).ConfigureAwait(false);
                foreach (var quote in crypto)
                    Report(quote);
            }
            catch (Exception e) { errors.Add(e); }
            WriteMarketPriceLog(startedAt, elapsed.Elapsed, targets, results, errors, scheduled);
            if (errors.Count > 0)
                throw new AggregateException("Market price requests failed.", errors);

            static async Task<T> FetchWithRetry<T>(Func<Task<T>> fetch)
            {
                // Three attempts in total; retry only the failed quote request.
                for (var attempt = 1; ; attempt++)
                {
                    try { return await fetch().ConfigureAwait(false); }
                    catch (Exception) when (attempt < 3)
                    { await Task.Delay(TimeSpan.FromSeconds(1)).ConfigureAwait(false); }
                }
            }
        }

        private static void WriteMarketPriceLog(DateTimeOffset startedAt, TimeSpan elapsed,
            IReadOnlyList<(string Code, HoldingType HoldingType)> targets,
            IReadOnlyDictionary<(string Code, HoldingType Type), MarketPrice> results, List<Exception> errors, bool scheduled)
        {
            try
            {
                var content = new StringBuilder().AppendLine(FormattableString.Invariant(
                    $"[{startedAt:O}] trigger={(scheduled ? "scheduled" : "manual-test")}; end={DateTimeOffset.Now:O}; elapsed={elapsed.TotalSeconds:F3}s; result={(errors.Count == 0 ? "OK" : "FAIL")}; succeeded={results.Count}; missing={targets.Count - results.Count}; errors={errors.Count}"));
                foreach (var target in targets)
                    content.AppendLine(results.TryGetValue(target, out var quote)
                        ? FormattableString.Invariant($"OK {target.Code}/{target.HoldingType}: {quote.Price} {quote.Currency}; fetchedAt={quote.FetchedAt:O}")
                        : $"MISSING {target.Code}/{target.HoldingType}");
                foreach (var error in errors.SelectMany(error => error is AggregateException aggregate
                    ? aggregate.Flatten().InnerExceptions.AsEnumerable() : [error]))
                {
                    // These providers construct public request summaries without credentials or response bodies.
                    var detail = error is InvalidOperationException &&
                        (error.Message.StartsWith("Yahoo Finance GET query1.finance.yahoo.com/v8/finance/chart/", StringComparison.Ordinal)
                        || error.Message.StartsWith("Kraken GET api.kraken.com/0/public/Ticker", StringComparison.Ordinal))
                        ? error.Message.Replace('\r', ' ').Replace('\n', ' ') : error.GetType().FullName;
                    content.AppendLine($"ERROR {detail}");
                }
                File.AppendAllText(Path.Combine(AppContext.BaseDirectory, "MyBook.market-prices.log"),
                    content.AppendLine().ToString(), new UTF8Encoding(false));
            }
            catch (Exception)
            {
                Console.Error.WriteLine("Could not write market price log.");
            }
        }

        public Task<List<Holding>> Fetch(Account account)
        {
            Console.WriteLine("skip account holding fetch: fetcher is not configured");
            return Task.FromResult(new List<Holding>());
        }

        public void Dispose()
        {
            httpClient.Dispose();
            googleHttpClient.Dispose();
            directHttpClient.Dispose();
            secondaryHttpClient.Dispose();
        }

        public async Task<JObject?> HttpGetJson(string url)
        {
            try
            {
                using HttpResponseMessage response = await httpClient.GetAsync(url).ConfigureAwait(false);
                if (response.StatusCode != HttpStatusCode.OK)
                    return null;

                var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                text = text.Substring(text.IndexOf('{'), text.LastIndexOf('}') - text.IndexOf('{') + 1);
                return (JObject?)JsonConvert.DeserializeObject(text);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch {url}: {e}");
            }

            return null;
        }

        public async Task<string?> HttpGetString(string url, HttpClient? client = null)
        {
            try
            {
                using HttpResponseMessage response = await (client ?? httpClient).GetAsync(url).ConfigureAwait(false);
                if (response.StatusCode != HttpStatusCode.OK)
                    return null;

                return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch {url}: {e}");
            }

            return null;
        }

        private static HttpClient CreateHttpClient(IWebProxy? proxy)
        {
            var handler = new HttpClientHandler
            {
                UseProxy = proxy is not null,
                Proxy = proxy
            };
            var client = new HttpClient(handler)
            {
                Timeout = HttpRequestTimeout
            };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36");
            client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.9");
            return client;
        }

        private static IWebProxy? ParsePubWebProxyConfig(string? value, string configKey = "pubweb_proxy")
        {
            if (String.IsNullOrWhiteSpace(value))
                return null;

            if (!Uri.TryCreate(value.Trim(), UriKind.Absolute, out var uri)
                || !String.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                || String.IsNullOrWhiteSpace(uri.Host)
                || uri.IsDefaultPort)
            {
                throw new InvalidOperationException($"Invalid {configKey} config. Expected http://host:port.");
            }

            if (!String.IsNullOrEmpty(uri.UserInfo)
                || !String.IsNullOrEmpty(uri.Query)
                || !String.IsNullOrEmpty(uri.Fragment)
                || uri.AbsolutePath != "/")
            {
                throw new InvalidOperationException($"Invalid {configKey} config. Proxy credentials, path, query, and fragment are not supported.");
            }

            return new WebProxy(uri);
        }
    }
}
