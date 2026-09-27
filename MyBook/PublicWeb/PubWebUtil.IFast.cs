using System.Text.Json;

namespace MyBook;

partial class PubWebUtil
{
    internal const string IFastInterestRateUrl = "https://www.ifastgb.com/api/current-account-setup/active-current-account-setup";

    internal sealed record IFastRateQuote(decimal Gross, decimal Aer);

    internal async Task<List<RateHistory>> ReadIFastInterestRates()
    {
        var json = await HttpGetString(IFastInterestRateUrl).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Cannot fetch IFast interest rates.");
        var quotes = ParseIFastRateQuotes(json);
        var observedAt = DateTime.Now;
        // The website does not provide an effective time. This is an observation, not a backdated rate.
        return quotes.Select(pair => new RateHistory
        {
            source = RateSource.IFastWebsite, currency = pair.Key,
            rateDate = observedAt, fetchedAt = observedAt,
            grossRate = pair.Value.Gross, aer = pair.Value.Aer
        }).ToList();
    }

    internal static Dictionary<CurrencyType, IFastRateQuote> ParseIFastRateQuotes(string json)
    {
        using var document = JsonDocument.Parse(json);
        var products = document.RootElement.EnumerateArray().ToList();
        var result = new Dictionary<CurrencyType, IFastRateQuote>();
        foreach (var code in new[] { "GBP", "USD", "EUR", "HKD", "SGD", "CNY" })
        {
            var matches = products.Where(product => product.GetProperty("currency").GetString() == code).ToList();
            if (matches.Count != 1)
                throw new InvalidOperationException($"IFast Gross rate missing or ambiguous: {code}.");
            if (matches[0].GetProperty("grossRateFrequency").GetString() != "Monthly")
                throw new InvalidOperationException($"Unexpected IFast interest frequency: {code}.");
            var rate = matches[0].GetProperty("grossRate").GetDecimal() / 100m;
            var aer = matches[0].GetProperty("aer").GetDecimal() / 100m;
            if (rate < 0 || rate >= 1 || aer < rate || aer >= 1)
                throw new InvalidOperationException($"Invalid IFast Gross rate: {code}.");
            result.Add(new Currency(0, code).t, new IFastRateQuote(rate, aer));
        }
        return result;
    }
}
