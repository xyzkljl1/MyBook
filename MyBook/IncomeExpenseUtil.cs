namespace MyBook;

public static class IncomeExpenseUtil
{
    // Contributions retain their currency; callers choose the period, grouping and conversion.
    public static (Currency Income, Currency Expense) GetContribution(Record record)
    {
        var income = new Currency(0, record.t);
        var expense = new Currency(0, record.t);
        if (record.Fake || record.isInternal || record.matchedRecordId != null || record.isRefundMatched)
            return (income, expense);

        // Reversals reduce the original category instead of becoming the opposite kind of flow.
        var isIncome = record.Reason switch
        {
            "持仓价格变动" or "利息" or "股息" or "债息"
                or "应计利息" or "应计债息" or "应计债券利息" or "应计股息"
                or "应计利息汇率变动" or "其它外汇换算"
                or "DP" or "视频收益" or "返现" => true,
            "消费" or "吃喝" or "日用品" or "水电网" or "虚拟产品" or "游戏"
                or "手续费" or "税费" => false,
            _ => record.v > 0
        };
        if (isIncome)
            income.v = record.v;
        else
            expense.v = -record.v;
        return (income, expense);
    }
}
