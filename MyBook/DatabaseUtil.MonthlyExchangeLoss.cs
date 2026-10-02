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
                && !r.Fake && !r.isRefundMatched)
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
        var rates = GetExchangeLossRates(history);
        // The first inserted Google quote for each currency is its fixed import starting point.
        var startDates = history.Where(r => r.source == RateSource.GoogleFinance).GroupBy(r => r.currency)
            .ToDictionary(g => g.Key, g => DateTime.SpecifyKind(g.MinBy(r => r.Id)!.rateDate, DateTimeKind.Local).ToUniversalTime().Date);
        var result = new Dictionary<int, decimal>();
        // Stop at the first missing quote so the last cached ID cannot advance past it.
        foreach (var record in records.Where(r => r.exchangeLossCache is null && !r.Fake && !r.isRefundMatched).OrderBy(r => r.Id))
        {
            var purchase = !record.isInternal && record.matchedRecordId is null && IsForeignFundedRmbPurchase(record);
            if (!purchase && !IsCurrencyConversion(record)) continue;
            var quoteDate = GetExchangeLossQuoteDate(record);
            var currency = record.t == CurrencyType.RMB ? record.DescCurrency!.t : record.t;
            if (startDates.TryGetValue(currency, out var start) && quoteDate < start
                || record.t != CurrencyType.RMB && record.DescCurrency!.t != CurrencyType.RMB
                    && startDates.TryGetValue(record.DescCurrency.t, out var targetStart) && quoteDate < targetStart)
                continue;
            if (GetExchangeLossRate(rates, currency, quoteDate,
                fromRmb: record.t == CurrencyType.RMB || record.DescCurrency!.t != CurrencyType.RMB) is not { } rate) break;
            if (purchase)
            {
                result.Add(record.Id, -record.v * rate + record.DescCurrency!.v);
            }
            // Conversion/repayment principal may be internal and matched; only this record's amounts are used.
            else
            {
                if (record.t == CurrencyType.RMB)
                    result.Add(record.Id, -record.v - Math.Abs(record.DescCurrency!.v) / rate);
                else if (record.DescCurrency!.t == CurrencyType.RMB)
                    result.Add(record.Id, -record.v * rate - Math.Abs(record.DescCurrency!.v));
                else if (GetExchangeLossRate(rates, record.DescCurrency.t, quoteDate, fromRmb: false) is { } targetRate)
                {
                    // rmb->外币->外币->rmb，其中rmb和外币之间使用实际基准汇率，以换一圈无损为标准计算中间的基准汇率。
                    // 中间基准汇率 = 1 / (转出币种 exchangeRateFromRmb * 转入币种 exchangeRateToRmb)。
                    // 直接比较两侧人民币价值，避免先计算中间汇率再换算造成额外舍入。
                    result.Add(record.Id, -record.v / rate - Math.Abs(record.DescCurrency.v) * targetRate);
                }
                else break;
            }
        }
        return result;
    }

    private static Dictionary<(CurrencyType, DateTime), RateHistory> GetExchangeLossRates(IEnumerable<RateHistory> history) =>
        history.Where(r => r.source == RateSource.GoogleFinance).OrderBy(r => r.rateDate)
            .GroupBy(r => (r.currency, DateTime.SpecifyKind(r.rateDate, DateTimeKind.Local).ToUniversalTime().Date))
            .ToDictionary(g => g.Key, g => g.Last());

    private static decimal? GetExchangeLossRate(IReadOnlyDictionary<(CurrencyType, DateTime), RateHistory> rates,
        CurrencyType currency, DateTime quoteDate, bool fromRmb)
    {
        decimal? previous = null;
        var previousDate = DateTime.MinValue;
        var hasLaterQuote = false;
        foreach (var ((rateCurrency, date), quote) in rates)
        {
            if (rateCurrency != currency) continue;
            var rate = fromRmb ? quote.exchangeRateFromRmb : quote.exchangeRateToRmb;
            if (rate is not > 0) continue;
            if (date == quoteDate) return rate;
            if (date > quoteDate) hasLaterQuote = true;
            else if (date > previousDate) { previousDate = date; previous = rate; }
        }
        // Only fill an interior gap in the same currency and direction; a missing tail may be delayed data.
        return hasLaterQuote ? previous : null;
    }

    private static DateTime GetExchangeLossQuoteDate(Record record) => record.date.Date;

    private static decimal? GetConversionRmbDebit(Record record, IReadOnlyDictionary<(CurrencyType, DateTime), RateHistory> rates)
    {
        if (record.t == CurrencyType.RMB) return -record.v;
        if (record.DescCurrency!.t == CurrencyType.RMB) return Math.Abs(record.DescCurrency.v) + record.exchangeLossCache;
        return -record.v / GetExchangeLossRate(rates, record.t, GetExchangeLossQuoteDate(record), fromRmb: true);
    }

    // Display calculations only; never writes Records or cached statistics.
    internal List<MonthlyRmbExpenseCalculation> CalculateMonthlyRmbExpenses(DateTime firstMonth, int months)
    {
        if (months <= 0 || months > 120) throw new ArgumentOutOfRangeException(nameof(months));
        firstMonth = new DateTime(firstMonth.Year, firstMonth.Month, 1);
        var end = firstMonth.AddMonths(months);
        var lifeIds = db.Queryable<Account>().Where(a => a.usage == AccountUsage.Life).Select(a => a.Id).ToList();
        var records = db.Queryable<Record>().Where(r => lifeIds.Contains(r._account_Id)
            && !r.Fake && !r.isRefundMatched
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
        var exchangeLossRates = GetExchangeLossRates(history);
        var ratesByCurrency = history.Where(r => r.source == RateSource.GoogleFinance && r.exchangeRateToRmb > 0)
            .GroupBy(r => r.currency).ToDictionary(g => g.Key,
                g => g.OrderBy(r => r.rateDate).ToList());
        var recordsByMonth = records.Where(r => !r.Fake && !r.isInternal && r.matchedRecordId == null && !r.isRefundMatched && r.date <= now)
            .ToLookup(r => new DateTime(r.date.Year, r.date.Month, 1));
        // Internal/matched conversion principal stays excluded; only its cached loss is an expense.
        var conversionsByMonth = records.Where(r => !r.Fake && !r.isRefundMatched && r.date <= now
            && r.exchangeLossCache.HasValue && IsCurrencyConversion(r))
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
            var ordinaryRecords = monthlyRecords.Where(r => !r.exchangeLossCache.HasValue
                || (!IsForeignFundedRmbPurchase(r) && !IsCurrencyConversion(r))).ToList();
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
            var conversions = conversionsByMonth[month].ToList();
            AddLossItem("消费汇损", lossTotal, debitTotal);
            AddLossItem("换汇汇损", conversions.Sum(r => r.exchangeLossCache!.Value),
                conversions.Aggregate((decimal?)0, (total, record) => total + GetConversionRmbDebit(record, exchangeLossRates)));
            series.Items = series.Items.OrderBy(item => item.IsIncome).ThenByDescending(item => item.Total)
                .ThenBy(item => item.Reason).ToList();
            series.TotalExpense = Currency.RoundMoney(series.Items.Where(item => !item.IsIncome).Sum(item => item.Total));
            series.TotalIncome = Currency.RoundMoney(series.Items.Where(item => item.IsIncome).Sum(item => item.Total));
            series.RateDescription = selectedRates.Count == 0 ? "未使用月末历史报价；汇损直接使用缓存。"
                : "Google 基准（RMB / 单位外币，报价日期 UTC）：" + String.Join("；", selectedRates.OrderBy(p => p.Key)
                    .Select(p => $"{p.Key} {p.Value.RmbPerUnit:0.########}（{p.Value.SourceDateUtc:yyyy-MM-dd}）"));
            result.Add(new(series, selectedRates));

            void AddLossItem(string reason, decimal loss, decimal? debit)
            {
                AddItem(reason, false, loss);
                var item = series.Items.FirstOrDefault(item => item.Reason == reason && !item.IsIncome);
                if (item is not null)
                    item.CurrencyDetails = debit is null or 0 ? "" : $"汇损{loss / debit.Value * 100:0.00}%";
            }

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

    // Paired conversions keep the other currency on the debit record; never count the credit again.
    private static bool IsCurrencyConversion(Record record) => record.Reason is "换汇" or "还款"
        && record.v < 0 && record.HoldingQuantity == 0
        && record.DescCurrency is { v: not 0 } description
        && record.t != description.t;

    private static bool IsForeignFundedRmbPurchase(Record record) => record.v < 0 && record.t != CurrencyType.RMB
        && record.DescCurrency is { t: CurrencyType.RMB, v: < 0 }
        && record.HoldingQuantity == 0
        && record.Reason is not ("换汇" or "还款" or "转账" or "内部转账" or "手续费" or "交易");
}

internal sealed record MonthlyReferenceRate(decimal RmbPerUnit, DateTime SourceDateUtc);
internal sealed record MonthlyRmbExpenseCalculation(ReasonFlowSeries Series,
    IReadOnlyDictionary<CurrencyType, MonthlyReferenceRate> Rates);
