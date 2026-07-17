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

        public PubWebUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.database = database;
            httpClient = CreateHttpClient(ParsePubWebProxyConfig(config["pubweb_proxy"]));
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

        public async Task FetchExchangeRates()
        {
            await FetchExchangeRates(Enum.GetValues<CurrencyType>()).ConfigureAwait(false);
        }

        public async Task FetchExchangeRates(IEnumerable<CurrencyType> currencyTypes)
        {
            var distinctCurrencyTypes = currencyTypes.Distinct().ToList();
            var fiatCurrencies = distinctCurrencyTypes.Where(currencyType => !IsCryptoCurrency(currencyType)).ToList();
            var rates = await Task.WhenAll(fiatCurrencies.Select(async currencyType =>
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

        internal static bool IsCryptoCurrency(CurrencyType currencyType)
        {
            return currencyType is CurrencyType.BTC or CurrencyType.ETH or CurrencyType.USDT;
        }

        public Task<List<Holding>> Fetch(Account account)
        {
            Console.WriteLine("skip account holding fetch: fetcher is not configured");
            return Task.FromResult(new List<Holding>());
        }

        public void Dispose()
        {
            httpClient.Dispose();
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

        public async Task<string?> HttpGetString(string url)
        {
            try
            {
                using HttpResponseMessage response = await httpClient.GetAsync(url).ConfigureAwait(false);
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

        private static IWebProxy? ParsePubWebProxyConfig(string? value)
        {
            if (String.IsNullOrWhiteSpace(value))
                return null;

            if (!Uri.TryCreate(value.Trim(), UriKind.Absolute, out var uri)
                || !String.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                || String.IsNullOrWhiteSpace(uri.Host)
                || uri.IsDefaultPort)
            {
                throw new InvalidOperationException("Invalid pubweb_proxy config. Expected http://host:port.");
            }

            if (!String.IsNullOrEmpty(uri.UserInfo)
                || !String.IsNullOrEmpty(uri.Query)
                || !String.IsNullOrEmpty(uri.Fragment)
                || uri.AbsolutePath != "/")
            {
                throw new InvalidOperationException("Invalid pubweb_proxy config. Proxy credentials, path, query, and fragment are not supported.");
            }

            return new WebProxy(uri);
        }
    }
}
