using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace MyBook;

partial class PubWebUtil
{
    internal static readonly RateSource[] KylcSources =
        [RateSource.KylcCcb, RateSource.KylcIcbc, RateSource.KylcCib, RateSource.KylcHfBank];

    // Kylc supplies a calendar date, not an intraday publication timestamp.
    // Buy quotes are CNY per foreign unit; sell quotes are inverted to foreign units per CNY.
    private static DateOnly KylcDate(DateTime localTime) => DateOnly.FromDateTime(
        new DateTimeOffset(DateTime.SpecifyKind(localTime, DateTimeKind.Local)).ToOffset(TimeSpan.FromHours(8)).DateTime);

    internal async Task FetchKylcRates(RateSource source)
    {
        var through = KylcDate(DateTime.Now).AddDays(-1);
        await FetchScheduledRatePairs(source, [CurrencyType.USD, CurrencyType.HKD, CurrencyType.GBP, CurrencyType.EUR],
            (currency, last) => KylcDate(last) < through
                ? ReadKylcRates(source, currency, KylcDate(last).AddDays(1), through)
                : Task.FromResult(new List<RateHistory>())).ConfigureAwait(false);
    }

    internal async Task<List<RateHistory>> ReadKylcRates(RateSource source, CurrencyType currency,
        DateOnly from, DateOnly through)
    {
        var sides = await Task.WhenAll(ReadKylcRateSide(source, currency, from, through, false),
            ReadKylcRateSide(source, currency, from, through, true)).ConfigureAwait(false);
        return PairExchangeRates(sides[0], sides[1]);
    }

    private async Task<List<RateHistory>> ReadKylcRateSide(RateSource source, CurrencyType currency,
        DateOnly from, DateOnly through, bool reverse)
    {
        var bankCode = source switch
        {
            RateSource.KylcCcb => "ccb", RateSource.KylcIcbc => "icbc",
            RateSource.KylcCib => "cib", RateSource.KylcHfBank => "hfbank",
            _ => throw new ArgumentOutOfRangeException(nameof(source))
        };
        if (currency is not (CurrencyType.USD or CurrencyType.HKD or CurrencyType.GBP or CurrencyType.EUR))
            throw new ArgumentOutOfRangeException(nameof(currency));
        var today = DateOnly.FromDateTime(DateTimeOffset.UtcNow.ToOffset(TimeSpan.FromHours(8)).DateTime);
        if (from > through || through >= today)
            throw new ArgumentOutOfRangeException(nameof(through), "Kylc history requires a completed date range in Beijing time.");

        var result = new List<RateHistory>();
        var path = $"www.kylc.com/huilv/d-{bankCode}-{currency.ToString().ToLowerInvariant()}" + (reverse ? ".html" : "/hui_buy.html");
        for (var start = from; start <= through;)
        {
            var end = through.DayNumber - start.DayNumber > 364 ? start.AddDays(364) : through;
            // The site's date picker requires two distinct dates, even for a one-day lookup.
            var queryStart = start == end ? start.AddDays(-1) : start;
            var request = $"Kylc GET {path} ({start:yyyy-MM-dd} to {end:yyyy-MM-dd})";
            try
            {
                using var response = await httpClient.GetAsync($"https://{path}?datefrom={queryStart.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}&dateto={end.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}").ConfigureAwait(false);
                if (response.StatusCode != HttpStatusCode.OK)
                    throw new InvalidOperationException($"{request}: HTTP {(int)response.StatusCode}.");
                var html = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                try { result.AddRange(ParseKylcHistory(html, source, bankCode, currency, start, end, DateTime.Now, reverse)); }
                catch (InvalidOperationException e)
                { throw new InvalidOperationException($"{request}: HTTP 200; {e.Message}"); }
            }
            catch (HttpRequestException e)
            { throw new InvalidOperationException($"{request}: {e.HttpRequestError}; HTTP {e.StatusCode?.ToString() ?? "unavailable"}."); }
            catch (OperationCanceledException)
            { throw new InvalidOperationException($"{request}: timeout; HTTP unavailable."); }
            start = end.AddDays(1);
        }
        return result;
    }

    internal static List<RateHistory> ParseKylcHistory(string html, RateSource source, string bankCode,
        CurrencyType currency, DateOnly from, DateOnly through, DateTime fetchedAt, bool reverse = false)
    {
        foreach (var (key, expected) in new[] { ("bank", bankCode.ToUpperInvariant()), ("ccy", currency.ToString()), ("mode", reverse ? "hui_sell" : "hui_buy") })
            if (!Regex.IsMatch(html, $@"\bvar\s+g_{key}\s*=\s*'{Regex.Escape(expected)}'\s*;"))
                throw new InvalidOperationException("Unexpected bank, currency or quote type.");
        var document = new HtmlDocument();
        document.LoadHtml(html);
        var table = document.DocumentNode.SelectSingleNode("//table[contains(concat(' ', normalize-space(@class), ' '), ' huilv-history-data-table ')]");
        if (table is null)
        {
            var empty = document.DocumentNode.SelectSingleNode("//*[contains(concat(' ', normalize-space(@class), ' '), ' huilv-history-empty-alert ')]");
            if (empty?.InnerText.Trim() == "没有相关汇率记录!") return [];
            throw new InvalidOperationException("Missing historical quote table.");
        }
        var headers = table.SelectNodes(".//th");
        if (headers is null || headers.Count < 2 || headers[0].InnerText.Trim() != "日期" || headers[1].InnerText.Trim() != (reverse ? "现汇卖出价" : "现汇买入价"))
            throw new InvalidOperationException("Unexpected historical quote columns.");

        var quotes = new Dictionary<DateOnly, RateHistory>();
        foreach (var row in table.SelectNodes(".//tr[td]") ?? Enumerable.Empty<HtmlNode>())
        {
            var cells = row.SelectNodes("./td");
            if (cells is null || cells.Count < 2
                || !DateOnly.TryParseExact(cells[0].InnerText.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var day)
                || !decimal.TryParse(cells[1].InnerText.Trim(), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var rate) || rate <= 0)
                throw new InvalidOperationException("Invalid historical quote row.");
            if (day < from || day > through) continue;
            if (!quotes.TryAdd(day, new RateHistory
            {
                source = source, currency = currency, exchangeRateToRmb = reverse ? null : rate,
                exchangeRateFromRmb = reverse ? Decimal.Round(1m / rate, 18, MidpointRounding.AwayFromZero) : null,
                fetchedAt = fetchedAt,
                rateDate = new DateTimeOffset(day.ToDateTime(TimeOnly.MinValue), TimeSpan.FromHours(8)).LocalDateTime
            }))
                throw new InvalidOperationException("Duplicate historical quote date.");
        }
        return quotes.Values.OrderBy(quote => quote.rateDate).ToList();
    }
}
