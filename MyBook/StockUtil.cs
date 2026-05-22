using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // 最新行情/汇率获取的调度入口与各来源共用的 HTTP 辅助逻辑。
    partial class StockUtil
    {
        private readonly DatabaseUtil? database;

        public StockUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            // 为了支持 gbk 编码。
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.database = database;
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
                    ret = new Currency(await FetchGoogleFinanceStock(finance.code, "NASDAQ"), CurrencyType.USD);
                    break;
                case HoldingType.ARCA:
                    ret = new Currency(await FetchGoogleFinanceStock(finance.code, "NYSEARCA"), CurrencyType.USD);
                    break;
                case HoldingType.UST:
                    Console.WriteLine("skip UST price: IB Gateway fetcher is marked Not used");
                    break;
                case HoldingType.SHANGHAI:
                    ret = new Currency(await FetchShanghaiStock(finance.code), CurrencyType.RMB);
                    break;
                case HoldingType.CNFUND:
                    ret = new Currency(await FetchCNFund(finance.code), CurrencyType.RMB);
                    break;
                case HoldingType.Cash:
                    var currencyType = Enum.TryParse<CurrencyType>(finance.code, out var parsedCurrencyType)
                        ? parsedCurrencyType
                        : finance.currentPrice.t;
                    ret = await FetchCurrencyToRmb(currencyType);
                    break;
                case HoldingType.Accrued:
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
            await FetchExchangeRates(Enum.GetValues<CurrencyType>());
        }

        public async Task FetchExchangeRates(IEnumerable<CurrencyType> currencyTypes)
        {
            foreach (var currencyType in currencyTypes.Distinct())
                await Fetch(new Finance(currencyType.ToString(), HoldingType.Cash));
        }

        public Task<List<Holding>> Fetch(Account account)
        {
            Console.WriteLine("skip account holding fetch: IB Gateway fetcher is marked Not used");
            return Task.FromResult(new List<Holding>());
        }

        public async Task<JObject?> HttpGetJson(string url)
        {
            try
            {
                using HttpClient client = new();
                using HttpResponseMessage response = await client.GetAsync(url);
                if (response.StatusCode != HttpStatusCode.OK)
                    return null;

                var text = await response.Content.ReadAsStringAsync();
                // 有些接口返回 JSONP，只取最外层大括号中的 JSON。
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
                using HttpClient client = new();
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36");
                client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.9");
                using HttpResponseMessage response = await client.GetAsync(url);
                if (response.StatusCode != HttpStatusCode.OK)
                    return null;

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch {url}: {e}");
            }

            return null;
        }
    }
}
