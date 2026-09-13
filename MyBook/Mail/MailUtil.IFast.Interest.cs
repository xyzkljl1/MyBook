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
        Dictionary<CurrencyType, decimal>? latestRates = null;
        DateTimeOffset retrievedAt = default;
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
            if (latestRates is null)
            {
                using var web = new PubWebUtil(config, database);
                latestRates = await web.FetchIFastInterestRates().ConfigureAwait(false);
                retrievedAt = DateTimeOffset.Now;
            }
            var all = GetIFastAccountRecords();
            var expected = CalculateIFastInterest(month, all, latestRates, retrievedAt);
            if (all.Any(record => record.Reason == "利息" && IFastBankDate(record) == postingDate))
                throw new InvalidOperationException($"IFast interest already exists without its monthly calculation marker: {month:yyyy-MM}.");
            database.SaveStatementRecordsOnce(IFastProvider, IFastLocalTime(postingDate), expected, statementKey: key,
                afterSaveInTransaction: _ =>
                {
                    var recalculated = CalculateIFastInterest(month, GetIFastAccountRecords(), latestRates, retrievedAt);
                    if (recalculated.Count != expected.Count || recalculated.Where((record, index) => record.v != expected[index].v).Any())
                        throw new InvalidOperationException("IFast interest inputs changed during import.");
                });
            Console.WriteLine($"IFast calculated interest {month:yyyy-MM}: {expected.Count} records using latest Gross rates, pending statement validation.");
        }
    }

    private List<Record> CalculateIFastInterest(DateTime month, List<Record> records,
        Dictionary<CurrencyType, decimal> rates, DateTimeOffset retrievedAt)
    {
        var result = new List<Record>();
        var end = month.AddMonths(1);
        foreach (var currency in IFastCurrencies)
        {
            if (!rates.TryGetValue(currency, out var rate) || rate < 0 || rate >= 1)
                throw new InvalidOperationException($"Missing or invalid IFast latest Gross rate: {currency}.");
            var account = database.GetAccountByName(IFastAccountName(currency));
            var entries = records.Where(record => record._account_Id == account.Id).ToList();
            if (entries.Any(record => record.t != currency))
                throw new InvalidOperationException($"IFast account contains an unexpected currency: {currency}.");
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
            record.Source = FormattableString.Invariant($"code={code}; calculation=IFast-interest-{month:yyyy-MM}; bankDate={end:yyyy-MM-dd}; latestGross={rate}; rateRetrievedAt={retrievedAt:O}; rateSource={PubWebUtil.IFastInterestRateUrl}; Gross/365; monthly AwayFromZero(2); no carry; inputs={hash}");
            result.Add(record);
        }
        return result;
    }
}
