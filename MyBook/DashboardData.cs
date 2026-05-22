namespace MyBook
{
    public class DashboardData
    {
        public List<CurrencyBalanceSummary> CurrencySummaries { get; set; } = [];
        public List<MonthlyFlowSeries> MonthlyFlowSeries { get; set; } = [];
        public MonthlyFlowSeries RmbMonthlyFlowSeries { get; set; } = new();
        public ReasonFlowSeries RmbReasonFlowSeries { get; set; } = new();
        public decimal TotalAssetsRmb { get; set; }
        public List<CurrencyType> MissingExchangeRateCurrencies { get; set; } = [];
        public DateTime LastMonthStart { get; set; }
        public DateTime LastMonthEnd { get; set; }
    }

    public class CurrencyBalanceSummary
    {
        public CurrencyType Currency { get; set; }
        public decimal Assets { get; set; }
        public decimal Liabilities { get; set; }
        public decimal Net { get; set; }
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
}
