using System.Globalization;

namespace MyBook
{
    // Tiantian Fund source for domestic fund prices.
    partial class PubWebUtil
    {
        public async Task<decimal> FetchCNFund(string code = "021282")
        {
            var url = $"http://fundgz.1234567.com.cn/js/{code}.js";
            try
            {
                var doc = await HttpGetJson(url).ConfigureAwait(false);
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
    }
}
