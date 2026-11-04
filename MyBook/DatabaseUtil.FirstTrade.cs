namespace MyBook;

partial class DatabaseUtil
{
    internal void OffsetFirstTradeFeeRefunds(int statementImportId)
    {
        var refunds = GetRecordsByStatementImport(statementImportId)
            .Where(r => r.Reason == "退款" && r.Source.Contains("|cash|FEE_REFUND ", StringComparison.Ordinal)).ToList();
        if (refunds.Count == 0) return;
        var previous = GetStatementRecords(StatementImportProvider.FirstTradeApi,
            refunds.Min(r => r.date.Date).AddDays(-7), refunds.Max(r => r.date.Date).AddDays(1));
        MarkRecordsAsRefundMatched(FindFirstTradeFeeRefundMatches(refunds, previous), withinImportTransaction: true);
    }

    internal static List<Record> FindFirstTradeFeeRefundMatches(IEnumerable<Record> refunds, IEnumerable<Record> previous)
    {
        var fees = previous.Where(r => r.Reason == "手续费" && r.v < 0 && !r.Fake && !r.isInternal
            && !r.isRefundMatched && r.matchedRecordId is null
            && r.Source.StartsWith("FirstTrade transaction|", StringComparison.Ordinal)
            && r.Source.Contains("|cash|FEE ", StringComparison.Ordinal)
            && r.Source.Contains("FOREIGN CARD TRANSACTION FEE ", StringComparison.Ordinal)).ToList();
        var matched = new List<Record>();
        foreach (var refund in refunds.Where(r => !r.isRefundMatched))
        {
            var candidates = fees.Where(r => r._account_Id == refund._account_Id && r.t == refund.t
                && r.v == -refund.v && r.date >= refund.date.Date.AddDays(-7) && r.date <= refund.date).ToList();
            if (refund.v <= 0 || refund.isInternal || refund.Fake || refund.matchedRecordId is not null || candidates.Count != 1)
                throw new WebUtil.FirstTradeException($"fee reimbursement on {refund.date:yyyy-MM-dd} requires exactly one unmatched foreign card fee in the previous 7 days");
            fees.Remove(candidates[0]);
            matched.Add(candidates[0]);
            matched.Add(refund);
        }
        return matched;
    }

    internal Dictionary<string, HoldingType> GetKnownEquityMarkets() => db.Queryable<Holding>()
        .Where(holding => holding.holdingType == HoldingType.NASDAQ || holding.holdingType == HoldingType.ARCA)
        .Select(holding => new { holding.code, holding.holdingType }).ToList()
        .GroupBy(holding => holding.code, StringComparer.Ordinal)
        .ToDictionary(group => group.Key, group => group.First().holdingType, StringComparer.Ordinal);

    internal WebUtil.FirstTradeTransferEvidence GetFirstTradeTransferEvidence(Account account, DateTime date, string symbol,
        decimal quantity, HoldingType holdingType)
    {
        var since = date.AddDays(-7);
        var until = date.AddDays(8);
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
