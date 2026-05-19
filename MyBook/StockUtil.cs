using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;

namespace MyBook
{
    class StockUtil
    {
        public StockUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            // 为了支持gbk编码
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

        public async Task<Currency?> FetchCurrencyToRmb(CurrencyType currencyType)
        {
            if (currencyType == CurrencyType.RMB)
                return new Currency(1, CurrencyType.RMB);

            var fromCurrency = currencyType.ToString();
            var url = $"https://www.google.com/finance/quote/{fromCurrency}-CNY?hl=en";
            try
            {
                var html = await HttpGetString(url);
                var rate = String.IsNullOrWhiteSpace(html) ? null : ParseGoogleFinanceExchangeRate(html, fromCurrency);
                if (rate is null)
                    return null;

                Console.WriteLine($"{fromCurrency}/CNY:{rate.Value}");
                return new Currency(rate.Value, CurrencyType.RMB);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch currency exchange rate {fromCurrency}/CNY: {e}");
            }

            return null;
        }

        private static decimal? ParseGoogleFinanceExchangeRate(string html, string fromCurrency)
        {
            var match = Regex.Match(html, $@"""{Regex.Escape(fromCurrency)} / CNY""\s*,\s*3\s*,\s*null\s*,\s*\[(?<rate>\d+(?:\.\d+)?)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                return null;

            return decimal.TryParse(match.Groups["rate"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var rate) ? rate : null;
        }

        // Google Finance 会同时返回常规交易价格和盘前/盘后价格；如有扩展时段价格，优先使用扩展时段价格。
        public async Task<decimal> FetchGoogleFinanceStock(string code, string exchange = "NASDAQ")
        {
            var symbol = code.Trim().ToUpperInvariant();
            var market = exchange.Trim().ToUpperInvariant();
            var url = $"https://www.google.com/finance/quote/{Uri.EscapeDataString(symbol)}:{market}?hl=en";
            try
            {
                var html = await HttpGetString(url);
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

        // 返回小于0表示错误。新浪行情字段 3 是当前价，停牌或未开盘时可能为 0，此时回退到昨收。
        public async Task<decimal> FetchShanghaiStock(string code)
        {
            var symbol = $"sh{code.Trim().ToLowerInvariant().Replace("sh", "")}";
            var url = $"https://hq.sinajs.cn/list={symbol}";
            try
            {
                var text = await HttpGetString(url);
                if (String.IsNullOrWhiteSpace(text))
                    return -1;

                var match = Regex.Match(text, "=\"(?<data>[^\"]*)\"");
                if (!match.Success)
                    return -1;

                var fields = match.Groups["data"].Value.Split(',');
                if (fields.Length < 4)
                    return -1;

                var priceText = fields[3];
                if (!Decimal.TryParse(priceText, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) || price <= 0)
                    Decimal.TryParse(fields[2], NumberStyles.Number, CultureInfo.InvariantCulture, out price);

                if (price <= 0)
                    return -1;

                Console.WriteLine($"{fields[0]}({symbol}):{price}");
                return price;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {code}: {e}");
            }

            return -1;
        }

        public async Task<decimal> FetchCNFund(string code = "021282")
        {
            var url = $"http://fundgz.1234567.com.cn/js/{code}.js";
            try
            {
                var doc = await HttpGetJson(url);
                if (doc is null)
                    return -1;

                var priceText = doc["gsz"]?.ToString();
                if (String.IsNullOrWhiteSpace(priceText) || !Decimal.TryParse(priceText, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) || price <= 0)
                    price = decimal.Parse(doc["dwjz"]!.ToString()!, CultureInfo.InvariantCulture);
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
                    text = text.Substring(text.IndexOf('{'), text.LastIndexOf('}') - text.IndexOf('{') + 1);
                    return (JObject?)JsonConvert.DeserializeObject(text);
                }
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
                using (HttpClient client = new())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36");
                    client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.9");
                    using (HttpResponseMessage response = await client.GetAsync(url))
                    {
                        if (response.StatusCode != HttpStatusCode.OK)
                            return null;
                        return await response.Content.ReadAsStringAsync();
                    }
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
