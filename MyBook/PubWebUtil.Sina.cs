using System.Globalization;
using System.Text.RegularExpressions;

namespace MyBook
{
    // Sina source for Shanghai Stock Exchange prices.
    partial class PubWebUtil
    {
        public async Task<decimal> FetchShanghaiStock(string code)
        {
            var symbol = $"sh{code.Trim().ToLowerInvariant().Replace("sh", "")}";
            var url = $"https://hq.sinajs.cn/list={symbol}";
            try
            {
                var text = await HttpGetString(url).ConfigureAwait(false);
                if (String.IsNullOrWhiteSpace(text))
                    return -1;

                var match = Regex.Match(text, "=\"(?<data>[^\"]*)\"");
                if (!match.Success)
                    return -1;

                var fields = match.Groups["data"].Value.Split(',');
                if (fields.Length <= 3)
                    return -1;

                var currentPriceText = fields[3];
                var previousCloseText = fields.Length > 2 ? fields[2] : "";
                var priceText = currentPriceText;
                if (!Decimal.TryParse(priceText, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) || price <= 0)
                    priceText = previousCloseText;
                if (!Decimal.TryParse(priceText, NumberStyles.Number, CultureInfo.InvariantCulture, out price) || price <= 0)
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
