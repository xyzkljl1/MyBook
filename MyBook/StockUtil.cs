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
        // 返回小于0表示错误
        public async Task<decimal> FetchUS(string code = "QQQ")
        {            
            var url = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&datatype=json&symbol={code}&apikey={key}";
            try
            {
                using (HttpClient client= new())
                using (HttpResponseMessage response = await client.GetAsync(url))
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                        return -1;
                    var text = await response.Content.ReadAsStringAsync();
                    var doc = (JObject?)JsonConvert.DeserializeObject(text);
                    if (doc is null)
                        return -1;
                    var meta = doc["Meta Data"]!.ToObject<JObject>()!;
                    var prices = doc["Time Series (Daily)"]!.ToObject<JObject>()!;
                    var date = meta["3. Last Refreshed"]!.ToString();
                    var price = decimal.Parse(prices[date]!["4. close"]!.ToString());
                    Console.WriteLine($"{meta["2. Symbol"]!.ToString()}({meta["3. Last Refreshed"]!.ToString()}):{price}");
                    return price;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch us stock {e}");
            }
            return 0;
        }
    }
}
