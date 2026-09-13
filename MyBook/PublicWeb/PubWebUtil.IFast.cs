using System.Text.Json;

namespace MyBook;

partial class PubWebUtil
{
    internal const string IFastInterestRateUrl = "https://www.ifastgb.com/api/current-account-setup/active-current-account-setup";

    public async Task<Dictionary<CurrencyType, decimal>> FetchIFastInterestRates()
    {
        var json = await HttpGetString(IFastInterestRateUrl).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Cannot fetch IFast interest rates.");
        return ParseIFastInterestRates(json);
    }

    internal static Dictionary<CurrencyType, decimal> ParseIFastInterestRates(string json)
    {
        using var document = JsonDocument.Parse(json);
        var products = document.RootElement.EnumerateArray().ToList();
        var result = new Dictionary<CurrencyType, decimal>();
        foreach (var code in new[] { "GBP", "USD", "EUR", "HKD", "SGD", "CNY" })
        {
            var matches = products.Where(product => product.GetProperty("currency").GetString() == code).ToList();
            if (matches.Count != 1)
                throw new InvalidOperationException($"IFast Gross rate missing or ambiguous: {code}.");
            if (matches[0].GetProperty("grossRateFrequency").GetString() != "Monthly")
                throw new InvalidOperationException($"Unexpected IFast interest frequency: {code}.");
            var rate = matches[0].GetProperty("grossRate").GetDecimal() / 100m;
            if (rate < 0 || rate >= 1)
                throw new InvalidOperationException($"Invalid IFast Gross rate: {code}.");
            result.Add(new Currency(0, code).t, rate);
        }
        return result;
    }
}
