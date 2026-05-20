using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // 持仓价格获取的调度入口与各来源共用的 HTTP 辅助逻辑。
    partial class StockUtil
    {
        public StockUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            // 为了支持 gbk 编码。
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public async Task<Currency?> Fetch(Stock stock)
        {
            Currency? ret = null;
            switch (stock.stockType)
            {
                case StockType.NASDAQ:
                    ret = new Currency(await FetchGoogleFinanceStock(stock.code, "NASDAQ"), CurrencyType.USD);
                    break;
                case StockType.ARCA:
                    ret = new Currency(await FetchGoogleFinanceStock(stock.code, "NYSEARCA"), CurrencyType.USD);
                    break;
                case StockType.UST:
                    Console.WriteLine("skip UST price: IB Gateway fetcher is marked Not used");
                    break;
                case StockType.SHANGHAI:
                    ret = new Currency(await FetchShanghaiStock(stock.code), CurrencyType.RMB);
                    break;
                case StockType.CNFUND:
                    ret = new Currency(await FetchCNFund(stock.code), CurrencyType.RMB);
                    break;
                case StockType.Cash:
                    ret = await FetchCurrencyToRmb(stock.currentPrice.t);
                    break;
            }

            ret = ret is null || ret.v < 0 ? null : ret;
            if (ret is not null)
            {
                stock.currentPrice = ret;
                stock.currentPriceTime = DateTimeOffset.Now.ToUnixTimeSeconds();
            }

            return ret;
        }

        public Task<List<Stock>> Fetch(Account account)
        {
            Console.WriteLine("skip account stock fetch: IB Gateway fetcher is marked Not used");
            return Task.FromResult(new List<Stock>());
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
