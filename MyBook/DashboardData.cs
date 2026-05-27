namespace MyBook
{
    public class DashboardData
    {
        public List<AssetSummaryPoint> AssetSummaryPoints { get; set; } = [];
        public List<CurrencyBalanceSummary> CurrencySummaries { get; set; } = [];
        public List<MonthlyFlowSeries> MonthlyFlowSeries { get; set; } = [];
        public MonthlyFlowSeries RmbMonthlyFlowSeries { get; set; } = new();
        public List<MonthlyFlowAccountStatistics> MonthlyAccounts { get; set; } = [];
        public List<AccountNetFlowStatistics> AccountNetFlows { get; set; } = [];
        public List<ReasonFlowSeries> RmbReasonFlowSeriesByMonth { get; set; } = [];
        public int DefaultReasonMonthIndex { get; set; }
        public InvestmentStatistics InvestmentByReason { get; set; } = new();
        public InvestmentStatistics InvestmentByHolding { get; set; } = new();
        public List<InvestmentAccountStatistics> InvestmentAccounts { get; set; } = [];
        public decimal TotalAssetsRmb { get; set; }
        public List<CurrencyType> MissingExchangeRateCurrencies { get; set; } = [];
        public DateTime LastMonthStart { get; set; }
        public DateTime LastMonthEnd { get; set; }
    }

    public class AssetSummaryPoint
    {
        public DateTime Date { get; set; }
        public DateTime? SnapshotTime { get; set; }
        public bool IsToday { get; set; }
        public bool HasData { get; set; }
        public decimal TotalAssetsRmb { get; set; }
        public List<CurrencyBalanceSummary> CurrencySummaries { get; set; } = [];
    }

    public class CurrencyBalanceSummary
    {
        public CurrencyType Currency { get; set; }
        public decimal Assets { get; set; }
        public decimal Liabilities { get; set; }
        public decimal Net { get; set; }
        public decimal TotalIncome { get; set; }
        public decimal TotalExpense { get; set; }
        public int AccountCount { get; set; }
    }

    public class MonthlyFlowSeries
    {
        public string DisplayName { get; set; } = "";
        public CurrencyType Currency { get; set; }
        public List<MonthlyFlowPoint> Points { get; set; } = [];
        public decimal TotalIncome { get; set; }
        public decimal TotalExpense { get; set; }
        public decimal NetChange { get; set; }
    }

    public class MonthlyFlowAccountStatistics
    {
        public string DisplayName { get; set; } = "";
        public List<MonthlyFlowSeries> MonthlyFlowSeries { get; set; } = [];
        public MonthlyFlowSeries RmbMonthlyFlowSeries { get; set; } = new();
    }

    public class AccountNetFlowStatistics
    {
        public string AccountPrefix { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public decimal NetRmb { get; set; }
        public bool HasMissingExchangeRate { get; set; }
        public int RecordCount { get; set; }
        public List<AccountNetFlowCurrencyTotal> CurrencyTotals { get; set; } = [];
    }

    public class AccountNetFlowCurrencyTotal
    {
        public CurrencyType Currency { get; set; }
        public decimal Amount { get; set; }
        public decimal? RmbAmount { get; set; }
    }

    public class MonthlyFlowPoint
    {
        public DateTime Month { get; set; }
        public string MonthLabel { get; set; } = "";
        public decimal Income { get; set; }
        public decimal Expense { get; set; }
        public decimal NetChange { get; set; }
        public List<MonthlyFlowSegment> IncomeSegments { get; set; } = [];
        public List<MonthlyFlowSegment> ExpenseSegments { get; set; } = [];
    }

    public class MonthlyFlowSegment
    {
        public CurrencyType Currency { get; set; }
        public decimal Value { get; set; }
        public string Label { get; set; } = "";
    }

    public class ReasonFlowSeries
    {
        public string DisplayName { get; set; } = "";
        public CurrencyType Currency { get; set; }
        public DateTime Month { get; set; }
        public string MonthLabel { get; set; } = "";
        public List<ReasonFlowItem> Items { get; set; } = [];
        public decimal TotalIncome { get; set; }
        public decimal TotalExpense { get; set; }
    }

    public class ReasonFlowItem
    {
        public string Reason { get; set; } = "";
        public bool IsIncome { get; set; }
        public decimal Total { get; set; }
        public string CurrencyDetails { get; set; } = "";
    }

    public class InvestmentStatistics
    {
        public List<InvestmentStatisticsPeriod> Periods { get; set; } = [];
    }

    public class InvestmentAccountStatistics
    {
        public string DisplayName { get; set; } = "";
        public InvestmentStatistics ByReason { get; set; } = new();
        public InvestmentStatistics ByHolding { get; set; } = new();
    }

    public class InvestmentStatisticsPeriod
    {
        public string Title { get; set; } = "";
        public List<InvestmentStatisticsItem> Items { get; set; } = [];
        public decimal Total { get; set; }
    }

    public class InvestmentStatisticsItem
    {
        public string Name { get; set; } = "";
        public decimal Total { get; set; }
    }

    public class RecordDetailData
    {
        public int Id { get; set; }
        public string AccountName { get; set; } = "";
        public decimal Amount { get; set; }
        public CurrencyType Currency { get; set; }
        public DateTime Date { get; set; }
        public string DestAccount { get; set; } = "";
        public bool IsInternal { get; set; }
        public int? MatchedRecordId { get; set; }
        public string MatchedRecordReason { get; set; } = "";
        public bool IsRefundMatched { get; set; }
        public int HoldingQuantity { get; set; }
        public string Source { get; set; } = "";
        public string Reason { get; set; } = "";
        public StatementImportProvider StatementProvider { get; set; }
        public string StatementKey { get; set; } = "";
        public bool HasBackup { get; set; }
    }

    public class RecordDetailEdit
    {
        public int Id { get; set; }
        public string AccountName { get; set; } = "";
        public decimal Amount { get; set; }
        public CurrencyType Currency { get; set; }
        public DateTime Date { get; set; }
        public string DestAccount { get; set; } = "";
        public bool IsInternal { get; set; }
        public bool IsRefundMatched { get; set; }
        public int HoldingQuantity { get; set; }
        public string Source { get; set; } = "";
        public string Reason { get; set; } = "";
    }

    public class AccountBalanceDetailData
    {
        public string AccountName { get; set; } = "";
        public CurrencyType Currency { get; set; }
        public decimal Amount { get; set; }
    }
}
