using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MailKit.Search;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    // PreviousAer is used only to validate a received notice, never persisted.
    internal sealed record IFastRateNotice(CurrencyType Currency, DateTime Date, decimal PreviousAer, decimal Aer);
    internal sealed record IFastScheduledRate(DateTime Date, decimal? Gross);

    private async Task UpdateIFastInterestRates()
    {
        // Independent of transaction-mail progress: notices can arrive before their effective dates.
        var messages = await SearchMessagesFromMailbox(CreateYahooMailbox() with { Proxy = null }, "IFast rate notices",
            SearchQuery.FromContains("@ifastgb.com").And(SearchQuery.SubjectContains("Interest rate update")),
            null, GetMailDateTime).ConfigureAwait(false);
        var notices = messages.SelectMany(ParseIFastRateNotice).ToList();
        var history = database.GetRateHistory(RateSource.IFastWebsite, RateSource.IFastMail);
        using var web = new PubWebUtil(config);
        var observations = await web.ReadIFastInterestRates().ConfigureAwait(false);
        // Keep at most one copy of each observed rate per bank day, preserving its first observation time.
        var rates = observations.Where(rate => !history.Any(item => item.source == rate.source
            && item.currency == rate.currency && IFastRateBankDate(item) == IFastRateBankDate(rate)
            && item.grossRate == rate.grossRate && item.aer == rate.aer)).ToList();
        ValidateIFastRateNotices(history.Concat(rates).ToList(), notices);
        var fetchedAt = DateTime.Now;
        rates.AddRange(notices.Select(notice => new RateHistory
        {
            source = RateSource.IFastMail,
            currency = notice.Currency,
            rateDate = IFastLocalTime(notice.Date),
            fetchedAt = fetchedAt,
            aer = notice.Aer
        }));
        BuildIFastRateSchedule(history.Concat(rates).ToList());
        database.SaveRateHistory(rates);
    }

    internal static List<IFastRateNotice> ParseIFastRateNotice(MimeMessage message)
    {
        var senders = message.From.Mailboxes.ToList();
        if (senders.Count != 1 || !senders[0].Address.EndsWith("@ifastgb.com", StringComparison.OrdinalIgnoreCase)
            || !message.Subject.Contains("Interest rate update", StringComparison.OrdinalIgnoreCase))
            throw new MailParseException("Invalid IFast rate notice sender or subject.");
        var doc = new HtmlDocument();
        doc.LoadHtml(message.HtmlBody ?? "");
        var text = NormalizeMailText(WebUtility.HtmlDecode(message.TextBody ?? doc.DocumentNode.InnerText));
        var matches = Regex.Matches(text,
            @"interest rate for (?<currency>[A-Z]{3}) Multi-Currency Current Account will change from (?<old>\d+(?:\.\d+)?)% AER to (?<new>\d+(?:\.\d+)?)% AER\.\s*This change will take effect on (?<date>\d{1,2} [A-Za-z]{3} \d{4})\.",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (matches.Count == 0 || matches.Count != Regex.Matches(text, "will change from", RegexOptions.IgnoreCase).Count)
            throw new MailParseException("Unsupported or incomplete IFast rate notice.");
        var result = matches.Select(match => new IFastRateNotice(new Currency(0, match.Groups["currency"].Value.ToUpperInvariant()).t,
            DateTime.ParseExact(match.Groups["date"].Value, "d MMM yyyy", CultureInfo.InvariantCulture),
            Decimal.Parse(match.Groups["old"].Value, CultureInfo.InvariantCulture) / 100m,
            Decimal.Parse(match.Groups["new"].Value, CultureInfo.InvariantCulture) / 100m)).ToList();
        if (result.Any(item => !IFastCurrencies.Contains(item.Currency) || item.PreviousAer < 0 || item.PreviousAer >= 1 || item.Aer < 0 || item.Aer >= 1)
            || result.GroupBy(item => item.Currency).Any(group => group.Count() != 1))
            throw new MailParseException("Invalid or duplicate IFast rate notice currency/rate.");
        return result;
    }

    private Dictionary<CurrencyType, List<IFastScheduledRate>> ReadIFastRateSchedule()
        => BuildIFastRateSchedule(database.GetRateHistory(RateSource.IFastWebsite, RateSource.IFastMail));

    private static DateTime IFastRateBankDate(RateHistory rate)
        => TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(rate.rateDate, DateTimeKind.Local), IFastTimeZone).Date;

    // Published AER is displayed to two percentage decimal places (e.g. 1.8048% becomes 1.80%).
    private static bool SameIFastAer(decimal left, decimal right)
        => Decimal.Round(left * 100, 2, MidpointRounding.AwayFromZero)
            == Decimal.Round(right * 100, 2, MidpointRounding.AwayFromZero);

    internal static void ValidateIFastRateNotices(List<RateHistory> history, List<IFastRateNotice> notices)
    {
        foreach (var notice in notices)
        {
            var first = history.Where(item => item.source == RateSource.IFastWebsite && item.currency == notice.Currency)
                .MinBy(item => item.rateDate)
                ?? throw new InvalidOperationException($"Missing IFast rate baseline: {notice.Currency}.");
            if (notice.Date < IFastRateBankDate(first)) continue;
            var earlier = history.Where(item => item.source == RateSource.IFastMail && item.currency == notice.Currency)
                .Select(item => (Date: IFastRateBankDate(item), Aer: item.aer!.Value))
                .Concat(notices.Where(item => item.Currency == notice.Currency).Select(item => (item.Date, item.Aer)))
                .Where(item => item.Date >= IFastRateBankDate(first) && item.Date < notice.Date)
                .OrderBy(item => item.Date).ToList();
            var previousAer = earlier.Count == 0 ? first.aer!.Value : earlier[^1].Aer;
            if (!SameIFastAer(previousAer, notice.PreviousAer)
                && !(notice.Date == IFastRateBankDate(first) && SameIFastAer(first.aer!.Value, notice.Aer)))
                throw new InvalidOperationException($"Broken IFast rate notice chain: {notice.Currency}, {notice.Date:yyyy-MM-dd}.");
        }
    }

    internal static Dictionary<CurrencyType, List<IFastScheduledRate>> BuildIFastRateSchedule(List<RateHistory> history)
    {
        var result = new Dictionary<CurrencyType, List<IFastScheduledRate>>();
        foreach (var currency in IFastCurrencies)
        {
            var snapshots = history.Where(item => item.source == RateSource.IFastWebsite && item.currency == currency)
                .OrderBy(item => item.rateDate).ToList();
            var first = snapshots.FirstOrDefault()
                ?? throw new InvalidOperationException($"Missing IFast rate baseline: {currency}.");
            var firstDate = IFastRateBankDate(first);
            var baseline = first.aer!.Value;
            var changes = history.Where(item => item.source == RateSource.IFastMail && item.currency == currency
                    && IFastRateBankDate(item) >= firstDate)
                .GroupBy(IFastRateBankDate).OrderBy(group => group.Key).Select(group =>
                {
                    var distinct = group.Select(item => item.aer!.Value).Distinct().ToList();
                    return distinct.Count == 1 ? (Date: group.Key, Aer: distinct[0])
                        : throw new InvalidOperationException($"Conflicting IFast rate notices: {currency}, {group.Key:yyyy-MM-dd}.");
                }).ToList();
            if (changes.Count == 0 || changes[0].Date > firstDate)
                changes.Insert(0, (firstDate, baseline));
            var schedule = new List<IFastScheduledRate>();
            for (var index = 0; index < changes.Count; index++)
            {
                var change = changes[index];
                var until = index + 1 < changes.Count ? changes[index + 1].Date : DateTime.MaxValue;
                var previousAer = index == 0 ? baseline : changes[index - 1].Aer;
                var observations = snapshots.Where(item => IFastRateBankDate(item) >= change.Date && IFastRateBankDate(item) < until).ToList();
                if (observations.Any(item => !SameIFastAer(item.aer!.Value, change.Aer)
                    && !(IFastRateBankDate(item) == change.Date && SameIFastAer(item.aer!.Value, previousAer))))
                    throw new InvalidOperationException($"Unannounced IFast rate change: {currency}, {change.Date:yyyy-MM-dd}.");
                var gross = observations.Where(item => SameIFastAer(item.aer!.Value, change.Aer))
                    .Select(item => item.grossRate).Distinct().ToList();
                if (gross.Count > 1) throw new InvalidOperationException($"Ambiguous IFast Gross rate: {currency}, {change.Date:yyyy-MM-dd}.");
                schedule.Add(new IFastScheduledRate(change.Date, gross.Count == 1 ? gross[0] : null));
            }
            result.Add(currency, schedule);
        }
        return result;
    }

    internal static decimal GetIFastDailyRate(List<IFastScheduledRate> schedule, DateTime day)
    {
        var rate = schedule.LastOrDefault(item => item.Date <= day);
        return rate?.Gross is decimal gross && gross >= 0 && gross < 1 ? gross
            : throw new InvalidOperationException($"Missing confirmed IFast Gross rate for {day:yyyy-MM-dd}; statement required for uncovered dates.");
    }
}
