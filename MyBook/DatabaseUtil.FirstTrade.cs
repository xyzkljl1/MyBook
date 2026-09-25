namespace MyBook;

partial class DatabaseUtil
{
    internal WebUtil.FirstTradeTransferEvidence GetFirstTradeTransferEvidence(Account account, DateTime date, string symbol, decimal quantity)
    {
        var since = date.AddDays(-7);
        var until = date.AddDays(8);
        var holdingType = GetKnownEquityHoldingType(symbol);
        var accounts = GetAllAccounts().ToDictionary(a => a.name, StringComparer.OrdinalIgnoreCase);
        var rows = db.Queryable<Record>().Includes(r => r.Holding).Includes(r => r.Account)
            .Where(r => r._account_Id != account.Id && r.date >= since && r.date < until
                && r.HoldingQuantity == -quantity && r.isInternal && r.t == CurrencyType.USD && !r.isRefundMatched)
            .ToList().Where(r => !IsAcatsCancellation(r)
                && r.Account is not null
                && r.Holding?.code == symbol && r.Holding.holdingType == holdingType && r.v < 0
                && r.matchedRecordId == null
                && ResolveInternalTransferTargetAccount(r, accounts)?.Id == account.Id).ToList();
        if (rows.Count != 1) throw new WebUtil.FirstTradeException(
            $"ACATS transfer on {date:yyyy-MM-dd} requires exactly one matching outgoing record; found {rows.Count}");
        return new(date, symbol, quantity, -rows[0].v, rows[0].Account!.name, rows[0].Id);
    }
}
