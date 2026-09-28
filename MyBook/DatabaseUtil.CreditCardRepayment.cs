using System.Globalization;

namespace MyBook;

partial class DatabaseUtil
{
    private void MatchFullStatementRepayments(StatementImportProvider billProvider, StatementImportProvider debitProvider,
        int statementImportId, List<Record> importedRecords, string repaymentSource)
    {
        var imported = importedRecords.Where(record => !record.Fake && !record.isRefundMatched
            && record.matchedRecordId is null && record.Reason == "还款"
            && record.Source.Contains(repaymentSource, StringComparison.Ordinal)
            && (record.v < 0 && record.t == CurrencyType.RMB || record.v > 0 && record.t != CurrencyType.RMB)).ToList();
        if (imported.Count == 0) return;
        var accounts = GetAllAccounts().ToDictionary(account => account.name, StringComparer.OrdinalIgnoreCase);
        foreach (var date in imported.Select(record => record.date.Date).Distinct())
        {
            var end = date.AddDays(1);
            // Keep same-day competitors for the two-way uniqueness check, but only pair newly imported entries.
            var candidates = db.Queryable<Record>()
                .InnerJoin<StatementImport>((record, statement) => record._statementImport_Id == statement.Id)
                .Where((record, statement) => record.date >= date && record.date < end
                    && !record.Fake && !record.isRefundMatched && record.matchedRecordId == null
                    && record.Reason == "还款" && record.Source.Contains(repaymentSource)
                    && (statement.provider == billProvider && record.v > 0 && record.t != CurrencyType.RMB
                        || statement.provider == debitProvider && record.v < 0 && record.t == CurrencyType.RMB))
                .Select((record, statement) => record).ToList();
            var credits = candidates.Where(record => record.v > 0).ToList();
            var debits = candidates.Where(record => record.v < 0).ToList();
            var targets = candidates.ToDictionary(record => record.Id,
                record => ResolveInternalTransferTargetAccount(record, accounts)?.Id);
            foreach (var (debit, credit) in FindFullStatementRepaymentPairs(credits, debits, targets, statementImportId))
            {
                var previousKey = db.Ado.GetString("""
                    SELECT MAX(statementKey) FROM StatementImports
                    WHERE provider = @provider AND statementKey <> '' AND statementKey < @repaymentDate
                    """, new { provider = billProvider.ToString(), repaymentDate = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) });
                if (String.IsNullOrEmpty(previousKey))
                    throw new InvalidOperationException($"Full statement repayment on {date:yyyy-MM-dd}: previous statement is missing.");

                // Include initialization, internal payments and refunds; statement membership defines the closing balance.
                var previousBalance = db.Ado.GetDecimal("""
                    SELECT COALESCE(SUM(r._Currency_v), 0) FROM Records r
                    JOIN StatementImports s ON s.Id = r._statementImport_Id
                    WHERE r._account_Id = @accountId AND r._Currency_t = @currency
                        AND s.provider = @provider AND s.statementKey <> '' AND s.statementKey <= @previousKey
                    """, new { accountId = credit._account_Id, currency = credit.t.ToString(), provider = billProvider.ToString(), previousKey });
                if (credit.v != -previousBalance)
                    throw new InvalidOperationException($"Full statement repayment mismatch: record={credit.Id}; "
                        + $"statement={previousKey}; debt={-previousBalance} {credit.t}; repayment={credit.v} {credit.t}.");

                debit.isInternal = credit.isInternal = true;
                MatchInternalTransferPair(debit, credit, "FullPreviousStatementRepayment");
            }
        }
    }

    private static List<(Record Debit, Record Credit)> FindFullStatementRepaymentPairs(List<Record> credits,
        List<Record> debits, Dictionary<int, int?> targets, int statementImportId)
    {
        var pairs = new List<(Record Debit, Record Credit)>();
        foreach (var day in credits.GroupBy(record => record.date.Date))
        {
            var outgoing = debits.Where(record => record.date.Date == day.Key).ToList();
            // SMS arrives before the following statement; either side can be imported first.
            if (outgoing.Count == 0) continue;
            // The user confirms these automatic repayments settle the entire previous statement.
            // Without a bank conversion ID, both directions must have exactly one eligible counterpart.
            var candidates = day.ToDictionary(credit => credit.Id, credit => outgoing.Where(debit =>
                (targets[credit.Id] is not { } debitAccount || debitAccount == debit._account_Id)
                && (targets[debit.Id] is not { } creditAccount || creditAccount == credit._account_Id)).ToList());
            foreach (var credit in day)
            {
                var matches = candidates[credit.Id];
                if (credit._statementImport_Id != statementImportId
                    && !matches.Any(record => record._statementImport_Id == statementImportId))
                    continue;
                if (matches.Count != 1 || candidates.Values.Count(list => list.Any(record => record.Id == matches[0].Id)) != 1)
                    throw new InvalidOperationException($"Ambiguous full statement repayment on {day.Key:yyyy-MM-dd}: "
                        + $"credit={credit.Id}; credits={day.Count()}; debits={outgoing.Count}; candidates={matches.Count}.");
                pairs.Add((matches[0], credit));
            }
        }
        return pairs;
    }
}
