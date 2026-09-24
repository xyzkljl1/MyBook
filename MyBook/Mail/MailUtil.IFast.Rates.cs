using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MailKit.Search;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    private const string IFastRateSnapshotPrefix = "IFast-rate-snapshot-";
    private const string IFastRateMailPrefix = "IFast-rate-mail-";
    private static readonly JsonSerializerOptions IFastRateJsonOptions = new()
        { Converters = { new JsonStringEnumConverter() } };
    internal sealed record IFastRateNotice(CurrencyType Currency, DateTime Date, decimal PreviousAer, decimal Aer);
    internal sealed record IFastRateSnapshot(DateTimeOffset RetrievedAt, Dictionary<CurrencyType, PubWebUtil.IFastRateQuote> Rates)
    {
        public DateTime Date => TimeZoneInfo.ConvertTime(RetrievedAt, IFastTimeZone).Date;
    }
    internal sealed record IFastScheduledRate(DateTime Date, decimal? Gross);

    private async Task UpdateIFastInterestRates()
    {
        // Independent of transaction-mail progress: notices can arrive before their effective dates.
        var messages = await SearchMessagesFromMailbox(CreateYahooMailbox() with { Proxy = null }, "IFast rate notices",
            SearchQuery.FromContains("@ifastgb.com").And(SearchQuery.SubjectContains("Interest rate update")),
            null, GetMailDateTime).ConfigureAwait(false);
        foreach (var message in messages)
        {
            var key = IFastRateMailPrefix + IFastMessageKey(message.MessageId);
            if (database.IsStatementKeyImported(IFastProvider, key)) continue;
            var notices = ParseIFastRateNotice(message);
            database.MarkStatementProcessedOnce(IFastProvider, GetMailDateTime(message), key,
                sourceDataJson: JsonSerializer.Serialize(notices, IFastRateJsonOptions));
        }
        using var web = new PubWebUtil(config, database);
        var rates = await web.FetchIFastRateQuotes().ConfigureAwait(false);
        var snapshot = new IFastRateSnapshot(DateTimeOffset.UtcNow, rates);
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(rates, IFastRateJsonOptions))));
        // At most one copy of each observed quote set per bank day; keep earlier observations intact.
        database.MarkStatementProcessedOnce(IFastProvider, IFastLocalTime(snapshot.Date),
            $"{IFastRateSnapshotPrefix}{snapshot.Date:yyyy-MM-dd}-{hash}",
            sourceDataJson: JsonSerializer.Serialize(snapshot, IFastRateJsonOptions));
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
    {
        var imports = database.GetStatementImports(IFastProvider);
        T Read<T>(StatementImport item) => JsonSerializer.Deserialize<T>(item.sourceDataJson
            ?? throw new InvalidOperationException("Missing IFast rate source data."), IFastRateJsonOptions)
            ?? throw new InvalidOperationException("Invalid IFast rate source data.");
        return BuildIFastRateSchedule(imports.Where(item => item.statementKey.StartsWith(IFastRateSnapshotPrefix, StringComparison.Ordinal))
                .Select(Read<IFastRateSnapshot>).ToList(),
            imports.Where(item => item.statementKey.StartsWith(IFastRateMailPrefix, StringComparison.Ordinal))
                .SelectMany(Read<List<IFastRateNotice>>).ToList());
    }

    internal static Dictionary<CurrencyType, List<IFastScheduledRate>> BuildIFastRateSchedule(
        List<IFastRateSnapshot> snapshots, List<IFastRateNotice> notices)
    {
        var first = snapshots.OrderBy(item => item.RetrievedAt).FirstOrDefault()
            ?? throw new InvalidOperationException("Missing IFast rate baseline.");
        // Published AER is displayed to two percentage decimal places (e.g. 1.8048% becomes 1.80%).
        static bool SameAer(decimal left, decimal right) => Decimal.Round(left * 100, 2, MidpointRounding.AwayFromZero)
            == Decimal.Round(right * 100, 2, MidpointRounding.AwayFromZero);
        var result = new Dictionary<CurrencyType, List<IFastScheduledRate>>();
        foreach (var currency in IFastCurrencies)
        {
            var changes = notices.Where(item => item.Currency == currency && item.Date >= first.Date)
                .GroupBy(item => item.Date).OrderBy(group => group.Key).Select(group =>
                {
                    var distinct = group.Distinct().ToList();
                    return distinct.Count == 1 ? distinct[0] : throw new InvalidOperationException($"Conflicting IFast rate notices: {currency}, {group.Key:yyyy-MM-dd}.");
                }).ToList();
            var baseline = first.Rates[currency].Aer;
            var previous = baseline;
            foreach (var change in changes)
            {
                if (!SameAer(previous, change.PreviousAer) && !(change.Date == first.Date && SameAer(baseline, change.Aer)))
                    throw new InvalidOperationException($"Broken IFast rate notice chain: {currency}, {change.Date:yyyy-MM-dd}.");
                previous = change.Aer;
            }
            if (changes.Count == 0 || changes[0].Date > first.Date)
                changes.Insert(0, new IFastRateNotice(currency, first.Date, baseline, baseline));
            var schedule = new List<IFastScheduledRate>();
            for (var index = 0; index < changes.Count; index++)
            {
                var change = changes[index];
                var until = index + 1 < changes.Count ? changes[index + 1].Date : DateTime.MaxValue;
                var observations = snapshots.Where(item => item.Date >= change.Date && item.Date < until).ToList();
                if (observations.Any(item => !SameAer(item.Rates[currency].Aer, change.Aer)
                    && !(item.Date == change.Date && SameAer(item.Rates[currency].Aer, change.PreviousAer))))
                    throw new InvalidOperationException($"Unannounced IFast rate change: {currency}, {change.Date:yyyy-MM-dd}.");
                var gross = observations.Where(item => SameAer(item.Rates[currency].Aer, change.Aer))
                    .Select(item => item.Rates[currency].Gross).Distinct().ToList();
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
