using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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

        public PubWebUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.database = database;
            httpClient = CreateHttpClient(ParsePubWebProxyConfig(config["pubweb_proxy"]));
            googleHttpClient = CreateHttpClient(ParsePubWebProxyConfig(config["pubweb_google_proxy"], "pubweb_google_proxy"));
        }

        public Task<Currency?> Fetch(Holding holding)
        {
            return Fetch(Finance.FromHolding(holding));
        }

        public async Task<Currency?> Fetch(Finance finance)
        {
            Currency? ret = null;
            switch (finance.holdingType)
            {
                case HoldingType.NASDAQ:
                    ret = new Currency(await FetchGoogleFinanceStock(finance.code, "NASDAQ").ConfigureAwait(false), CurrencyType.USD);
                    break;
                case HoldingType.ARCA:
                    ret = new Currency(await FetchGoogleFinanceStock(finance.code, "NYSEARCA").ConfigureAwait(false), CurrencyType.USD);
                    break;
                case HoldingType.UST:
                    Console.WriteLine("skip UST price: fetcher is not configured");
                    break;
                case HoldingType.SHANGHAI:
                    ret = new Currency(await FetchShanghaiStock(finance.code).ConfigureAwait(false), CurrencyType.RMB);
                    break;
                case HoldingType.CNFUND:
                    ret = new Currency(await FetchCNFund(finance.code).ConfigureAwait(false), CurrencyType.RMB);
                    break;
                case HoldingType.Cash:
                    var currencyType = Enum.TryParse<CurrencyType>(finance.code, out var parsedCurrencyType)
                        ? parsedCurrencyType
                        : finance.currentPrice.t;
                    ret = await FetchCurrencyToRmb(currencyType).ConfigureAwait(false);
                    break;
                case HoldingType.Accrued:
                    break;
                case HoldingType.Crypto:
                    break;
            }

            ret = ret is null || ret.v < 0 ? null : ret;
            if (ret is not null)
            {
                finance.currentPrice = ret;
                finance.currentPriceTime = DateTimeOffset.Now.ToUnixTimeSeconds();
                database?.SaveFinance(finance);
            }

            return ret;
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
                await ImportSchedule.RunAsync($"{source}/{currency}", 1, 5, () => latest, async since =>
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
                : KylcDate(rate.rateDate);
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

        public async Task FetchExchangeRates(IEnumerable<CurrencyType> currencyTypes)
        {
            var distinctCurrencyTypes = currencyTypes.Distinct().ToList();
            var rates = await Task.WhenAll(distinctCurrencyTypes.Select(async currencyType =>
                (CurrencyType: currencyType, Rate: await FetchCurrencyToRmb(currencyType).ConfigureAwait(false)))).ConfigureAwait(false);

            foreach (var (currencyType, rate) in rates)
            {
                if (rate is null || rate.v < 0)
                    continue;

                var finance = new Finance(currencyType.ToString(), HoldingType.Cash)
                {
                    currentPrice = rate,
                    currentPriceTime = DateTimeOffset.Now.ToUnixTimeSeconds()
                };
                database?.SaveFinance(finance);
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
