using System.Globalization;
using System.Text.RegularExpressions;

namespace MyBook
{
    // 新浪行情来源：上海交易所股票实时价格。
    partial class StockUtil
    {
        // 返回小于 0 表示错误。新浪行情字段 3 是当前价，停牌或未开盘时可能为 0，此时回退到昨收。
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
    }
}
