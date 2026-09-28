namespace MyBook;

partial class DatabaseUtil
{
    internal int CacheExchangeLosses()
    {
        return ExecuteLockedTransaction(() =>
        {
            var last = db.Queryable<Record>().Where(r => r.exchangeLossCache != null)
                .OrderByDescending(r => r.Id).First();
            var lastId = last?.Id ?? 0;
            var records = db.Queryable<Record>().Where(r => r.Id > lastId && r.exchangeLossCache == null
                && !r.Fake && !r.isInternal && r.matchedRecordId == null && !r.isRefundMatched)
                .ToList();
            var updates = BuildExchangeLossCache(records, GetRateHistory(RateSource.GoogleFinance));
            foreach (var update in updates)
                db.Updateable<Record>().SetColumns(r => r.exchangeLossCache == update.Value)
                    .Where(r => r.Id == update.Key && r.exchangeLossCache == null).ExecuteCommand();
            return updates.Count;
        });
    }

    internal static Dictionary<int, decimal> BuildExchangeLossCache(IEnumerable<Record> records, IEnumerable<RateHistory> history)
    {
        var rates = history.Where(r => r.source == RateSource.GoogleFinance && r.exchangeRateToRmb > 0)
            .OrderBy(r => r.rateDate)
            .GroupBy(r => (r.currency, DateTime.SpecifyKind(r.rateDate, DateTimeKind.Local).ToUniversalTime().Date))
            .ToDictionary(g => g.Key, g => g.Last().exchangeRateToRmb!.Value);
        var result = new Dictionary<int, decimal>();
        foreach (var record in records.Where(r => r.exchangeLossCache is null && !r.Fake && !r.isInternal
                     && r.matchedRecordId is null && !r.isRefundMatched && IsForeignFundedRmbPurchase(r)))
        {
            // Exact transaction-day quote only; no previous-date fallback and no recalculation.
            if (rates.TryGetValue((record.t, record.date.Date), out var rate))
                result.Add(record.Id, -record.v * rate + record.DescCurrency!.v);
        }
        return result;
    }

    // Display calculations only; never writes Records or cached statistics.
    internal List<MonthlyRmbExpenseCalculation> CalculateMonthlyRmbExpenses(DateTime firstMonth, int months)
    {
        if (months <= 0 || months > 120) throw new ArgumentOutOfRangeException(nameof(months));
        firstMonth = new DateTime(firstMonth.Year, firstMonth.Month, 1);
        var end = firstMonth.AddMonths(months);
        var lifeIds = db.Queryable<Account>().Where(a => a.usage == AccountUsage.Life).Select(a => a.Id).ToList();
        var records = db.Queryable<Record>().Where(r => lifeIds.Contains(r._account_Id)
            && !r.Fake && !r.isInternal && r.matchedRecordId == null && !r.isRefundMatched
            && r.date >= firstMonth && r.date < end).ToList();
        return CalculateMonthlyRmbExpenses(records, GetRateHistory(RateSource.GoogleFinance), firstMonth, months, DateTime.Now);
    }

    internal static List<MonthlyRmbExpenseCalculation> CalculateMonthlyRmbExpenses(
        IReadOnlyList<Record> records, IReadOnlyList<RateHistory> history, DateTime firstMonth, int months, DateTime now, bool allowUnavailableMonths = false, Dictionary<CurrencyType, decimal>? fallbackRates = null)
    {
        if (months <= 0 || months > 120) throw new ArgumentOutOfRangeException(nameof(months));
        firstMonth = new DateTime(firstMonth.Year, firstMonth.Month, 1);
        var currentMonth = new DateTime(now.Year, now.Month, 1);
        var utcToday = DateTime.SpecifyKind(now, DateTimeKind.Local).ToUniversalTime().Date;
        var ratesByCurrency = history.Where(r => r.source == RateSource.GoogleFinance && r.exchangeRateToRmb > 0)
            .GroupBy(r => r.currency).ToDictionary(g => g.Key,
                g => g.OrderBy(r => r.rateDate).ToList());
        var recordsByMonth = records.Where(r => !r.Fake && !r.isInternal && r.matchedRecordId == null && !r.isRefundMatched && r.date <= now)
            .ToLookup(r => new DateTime(r.date.Year, r.date.Month, 1));
        var result = new List<MonthlyRmbExpenseCalculation>();
        for (var i = 0; i < months; i++)
        {
            var month = firstMonth.AddMonths(i);
            if (month > currentMonth) break;
            var cutoff = month == currentMonth ? utcToday : month.AddMonths(1).AddDays(-1);
            var monthlyRecords = recordsByMonth[month].ToList();
            // Pending RMB purchases follow ordinary expense conversion until the scheduler fills their cache.
            var cachedPurchases = monthlyRecords.Where(r => IsForeignFundedRmbPurchase(r) && r.exchangeLossCache.HasValue).ToList();
            var ordinaryRecords = monthlyRecords.Where(r => !IsForeignFundedRmbPurchase(r) || !r.exchangeLossCache.HasValue).ToList();
            var rates = new Dictionary<CurrencyType, decimal> { [CurrencyType.RMB] = 1m };
            var selectedRates = new Dictionary<CurrencyType, MonthlyReferenceRate>();
            var missing = new List<string>();
            foreach (var currency in ordinaryRecords.Select(r => r.t).Distinct().Where(c => c != CurrencyType.RMB))
            {
                // Google quote dates are UTC dates, even though RateHistory stores local timestamps.
                var quote = ratesByCurrency.GetValueOrDefault(currency)?.LastOrDefault(r =>
                    DateTime.SpecifyKind(r.rateDate, DateTimeKind.Local).ToUniversalTime().Date <= cutoff);
                if (quote is null)
                {
                    if (fallbackRates is not null && fallbackRates.TryGetValue(currency, out var fallback))
                    {
                        rates[currency] = fallback;
                        continue;
                    }
                    if (!allowUnavailableMonths)
                        throw new InvalidOperationException($"Monthly expense: GoogleFinance rate missing for {currency} at {cutoff:yyyy-MM-dd}.");
                    missing.Add($"{currency} {cutoff:yyyy-MM-dd}");
                    continue;
                }
                rates[currency] = quote.exchangeRateToRmb!.Value;
                selectedRates[currency] = new(quote.exchangeRateToRmb.Value,
                    DateTime.SpecifyKind(quote.rateDate, DateTimeKind.Local).ToUniversalTime().Date);
            }

            if (missing.Count > 0)
            {
                result.Add(new(new ReasonFlowSeries
                {
                    Currency = CurrencyType.RMB, Month = month, MonthLabel = month.ToString("yyyy年MM月"),
                    IsAvailable = false,
                    RateDescription = $"缺少指定日期或此前的 Google 汇率：{String.Join("、", missing)}"
                }, selectedRates));
                continue;
            }
            var series = BuildRmbReasonFlowSeries(ordinaryRecords,
                month, month.AddMonths(1), rates);
            decimal lossTotal = 0;
            decimal debitTotal = 0;
            foreach (var record in cachedPurchases)
            {
                var principal = -record.DescCurrency!.v;
                var loss = record.exchangeLossCache!.Value;
                lossTotal += loss;
                debitTotal += principal + loss;
                AddItem(String.IsNullOrWhiteSpace(record.Reason) ? "未分类" : record.Reason, false, principal);
            }
            // Net gains reduce the same expense item; a net gain remains a negative exchange loss.
            AddItem("汇损", false, lossTotal);
            series.Items = series.Items.OrderBy(item => item.IsIncome).ThenByDescending(item => item.Total)
                .ThenBy(item => item.Reason).ToList();
            series.TotalExpense = Currency.RoundMoney(series.Items.Where(item => !item.IsIncome).Sum(item => item.Total));
            series.TotalIncome = Currency.RoundMoney(series.Items.Where(item => item.IsIncome).Sum(item => item.Total));
            series.RateDescription = selectedRates.Count == 0 ? "未使用月末历史报价；汇损直接使用缓存。"
                : "Google 基准（RMB / 单位外币，报价日期 UTC）：" + String.Join("；", selectedRates.OrderBy(p => p.Key)
                    .Select(p => $"{p.Key} {p.Value.RmbPerUnit:0.########}（{p.Value.SourceDateUtc:yyyy-MM-dd}）"));
            foreach (var item in series.Items.Where(item => item.Reason == "汇损"))
                item.CurrencyDetails = debitTotal == 0 ? "" : $"汇损{item.Total / debitTotal * 100:0.00}%";
            result.Add(new(series, selectedRates));

            void AddItem(string reason, bool income, decimal amount)
            {
                if (amount == 0) return;
                var item = series.Items.FirstOrDefault(item => item.Reason == reason && item.IsIncome == income);
                if (item is null) series.Items.Add(new ReasonFlowItem { Reason = reason, IsIncome = income, Total = amount });
                else item.Total += amount;
            }
        }
        return result;
    }

    private static bool IsForeignFundedRmbPurchase(Record record) => record.v < 0 && record.t != CurrencyType.RMB
        && record.DescCurrency is { t: CurrencyType.RMB, v: < 0 }
        && record.HoldingQuantity == 0
        && record.Reason is not ("换汇" or "转账" or "内部转账" or "手续费" or "交易");
}

internal sealed record MonthlyReferenceRate(decimal RmbPerUnit, DateTime SourceDateUtc);
internal sealed record MonthlyRmbExpenseCalculation(ReasonFlowSeries Series,
    IReadOnlyDictionary<CurrencyType, MonthlyReferenceRate> Rates);
