using System.Globalization;
using System.Text.RegularExpressions;

namespace MyBook
{
    // Google Finance source for US stocks, ETFs, and FX rates to CNY.
    partial class PubWebUtil
    {
        public async Task<Currency?> FetchCurrencyToRmb(CurrencyType currencyType)
        {
            if (currencyType == CurrencyType.RMB)
                return new Currency(1, CurrencyType.RMB);

            var fromCurrency = currencyType.ToString();
            var url = $"https://www.google.com/finance/quote/{fromCurrency}-CNY?hl=en";
            try
            {
                var html = await HttpGetString(url).ConfigureAwait(false);
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

        public async Task<decimal> FetchGoogleFinanceStock(string code, string exchange = "NASDAQ")
        {
            var symbol = code.Trim().ToUpperInvariant();
            var market = exchange.Trim().ToUpperInvariant();
            var url = $"https://www.google.com/finance/quote/{Uri.EscapeDataString(symbol)}:{market}?hl=en";
            try
            {
                var html = await HttpGetString(url).ConfigureAwait(false);
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
    }
}
