using MailKit.Net.Proxy;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook
{

    class StockUtil
    {
        string key;
        public StockUtil(IConfigurationRoot config)
        {
            // 为了支持gbk编码
            System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            key = config["alphavantage_key"]!;
        }
        public async Task<Currency?> Fetch(Stock stock)
        {
            Currency? ret = null;
            switch(stock.t)
            {
                case StockType.US:
                    ret = new Currency(await FetchStock(stock.code), CurrencyType.USD);
                    break;
                case StockType.SHANGHAI:
                    ret = new Currency(await FetchStock(stock.code + ".SHH"), CurrencyType.RMB);
                    break;
                case StockType.CNFUND:
                    ret = new Currency(await FetchCNFund(stock.code), CurrencyType.RMB);
                    break;
            }
            ret = ret==null||ret.v < 0 ? null : ret;
            return ret;
        }
        // 返回小于0表示错误
        // 股价api不返回币种，caller处理币种
        public async Task<decimal> FetchStock(string code = "QQQ")
        {            
            var url = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&datatype=json&symbol={code}&apikey={key}";
            try
            {
                var doc = await HttpGetJson(url);
                if (doc is null)
                    return -1;
                var meta = doc["Meta Data"]!.ToObject<JObject>()!;
                var prices = doc["Time Series (Daily)"]!.ToObject<JObject>()!;
                var date = meta["3. Last Refreshed"]!.ToString();
                var price = decimal.Parse(prices[date]!["4. close"]!.ToString());
                Console.WriteLine($"{meta["2. Symbol"]!.ToString()}({meta["3. Last Refreshed"]!.ToString()}):{price}");
                return price;
        }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {code}: {e}");
            }
            return 0;
        }
        public async Task<decimal> FetchCNFund(string code = "021282")
        {
            var url = $"http://fundgz.1234567.com.cn/js/{code}.js";
            try
            {
                var doc = await HttpGetJson(url);
                if (doc is null)
                    return -1;
                var price = decimal.Parse(doc["dwjz"]!.ToString()!); // 单位净值
                var name = doc["name"]!.ToString()!;
                var date = doc["gztime"]!.ToString();
                Console.WriteLine($"{name}({date}):{price}");
                return price;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {code}: {e}");
            }
            return 0;
        }
        public async Task<JObject?> HttpGetJson(string url)
        {
            try
            {
                using (HttpClient client = new())
                using (HttpResponseMessage response = await client.GetAsync(url))
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                        return null;
                    var text = await response.Content.ReadAsStringAsync();
                    // 只选最外层的大括号里面的东西
                    text = text.Substring(text.IndexOf('{'), text.LastIndexOf('}') - text.IndexOf('{')+1);
                    return (JObject?)JsonConvert.DeserializeObject(text);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch {url}: {e}");
            }
            return null;
        }
    }
}
