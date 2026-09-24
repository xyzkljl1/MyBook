using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace MyBook;

partial class MailUtil
{
    private async Task GenerateIFastInterest()
    {
        var range = GetIFastInitializationRange();
        var thisMonth = FirstDayOfMonth(TimeZoneInfo.ConvertTime(DateTimeOffset.Now, IFastTimeZone).Date);
        var statements = ReadIFastStatements();
        var rates = ReadIFastRateSchedule();
        var baselineDate = rates.Values.Min(schedule => schedule[0].Date);
        var calculatedAt = DateTimeOffset.Now;
        for (var month = range.End.AddMonths(-1); month < thisMonth; month = month.AddMonths(1))
        {
            var postingDate = month.AddMonths(1);
            var key = $"IFast-interest-{month:yyyy-MM}";
            // Interest is listed in the following month's statement, on its posting date.
            var statement = statements.SingleOrDefault(item => item.Month == postingDate);
            if (statement is not null)
            {
                ImportIFastStatementInterest(statement);
                continue;
            }
            if (database.IsStatementKeyImported(IFastProvider, key)
                || database.IsStatementKeyImported(IFastProvider, $"IFast-statement-{postingDate:yyyy-MM}"))
                continue;
            // Forward-only rate tracking: do not reconstruct unimported historical interest.
            if (postingDate <= baselineDate) continue;
            var all = GetIFastAccountRecords();
            var expected = CalculateIFastInterest(month, all, rates, calculatedAt);
            if (all.Any(record => record.Reason == "利息" && IFastBankDate(record) == postingDate))
                throw new InvalidOperationException($"IFast interest already exists without its monthly calculation marker: {month:yyyy-MM}.");
            database.SaveStatementRecordsOnce(IFastProvider, IFastLocalTime(postingDate), expected, statementKey: key,
                afterSaveInTransaction: _ =>
                {
                    var recalculated = CalculateIFastInterest(month, GetIFastAccountRecords(), rates, calculatedAt);
                    if (recalculated.Count != expected.Count || recalculated.Where((record, index) => record.v != expected[index].v).Any())
                        throw new InvalidOperationException("IFast interest inputs changed during import.");
                });
            Console.WriteLine($"IFast calculated interest {month:yyyy-MM}: {expected.Count} records using dated Gross rates, pending statement validation.");
        }
    }

    private List<Record> CalculateIFastInterest(DateTime month, List<Record> records,
        Dictionary<CurrencyType, List<IFastScheduledRate>> rates, DateTimeOffset calculatedAt)
    {
        var result = new List<Record>();
        var end = month.AddMonths(1);
        var account = GetIFastAccount();
        if (records.Any(record => record._account_Id != account.Id || !IFastCurrencies.Contains(record.t)))
            throw new InvalidOperationException("IFast interest inputs contain an unexpected account or currency.");
        foreach (var currency in IFastCurrencies)
        {
            if (!rates.TryGetValue(currency, out var schedule))
                throw new InvalidOperationException($"Missing IFast rate schedule: {currency}.");
            var entries = records.Where(record => record.t == currency).ToList();
            var balance = entries.Where(record => IFastBankDate(record) < month).Sum(record => record.v);
            var movements = entries.Where(record => IFastBankDate(record) >= month && IFastBankDate(record) < end)
                .GroupBy(IFastBankDate).ToDictionary(group => group.Key, group => group.Sum(record => record.v));
            decimal interest = 0;
            var inputs = new StringBuilder();
            for (var day = month; day < end; day = day.AddDays(1))
            {
                balance += movements.GetValueOrDefault(day);
                if (balance < 0) throw new InvalidOperationException($"Negative IFast daily balance: {currency}, {day:yyyy-MM-dd}.");
                if (balance == 0) continue;
                var rate = GetIFastDailyRate(schedule, day);
                // User-approved provisional convention: Actual/365, no daily cent rounding,
                // monthly midpoint rounding away from zero; no fractional-cent carry.
                interest += balance * rate / 365m;
                inputs.Append(CultureInfo.InvariantCulture, $"{day:yyyy-MM-dd}:{balance}:{rate};");
            }
            var amount = Decimal.Round(interest, 2, MidpointRounding.AwayFromZero);
            if (amount == 0) continue;
            var code = $"IFast-interest-{month:yyyy-MM}-{currency}";
            var record = BuildIFastRecord(new Currency(amount, currency), IFastLocalTime(end), "利息", $"{month:yyyy-MM} 利息（计算）", code);
            var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(inputs.ToString())));
            var appliedRates = String.Join(",", schedule.Where(item => item.Date < end)
                .Select(item => FormattableString.Invariant($"{item.Date:yyyy-MM-dd}:{item.Gross}")));
            record.Source = FormattableString.Invariant($"code={code}; calculation=IFast-interest-{month:yyyy-MM}; bankDate={end:yyyy-MM-dd}; grossSchedule={appliedRates}; calculatedAt={calculatedAt:O}; rateSource={PubWebUtil.IFastInterestRateUrl}; Gross/365; monthly AwayFromZero(2); no carry; inputs={hash}");
            result.Add(record);
        }
        return result;
    }
}
