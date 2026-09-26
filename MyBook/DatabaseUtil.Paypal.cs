namespace MyBook;

partial class DatabaseUtil
{
    internal void SavePayPalCombined(IReadOnlyList<CombinedUtil.PayPalImport> imports, CombinedUtil.PayPalPlan plan)
    {
        ExecuteLockedTransaction(() =>
        {
            if (imports.All(i => GetLatestStatementImport(i.Provider)?.statementKey == i.Key)) return;
            foreach (var import in imports)
                if (GetLatestStatementImport(import.Provider)?.statementKey != import.PreviousKey)
                    throw new InvalidOperationException("PayPal import state changed during retrieval; retry required.");
            foreach (var expected in plan.ExpectedItemAccounts)
                if (!db.Queryable<PlaidItem>().Any(i => i.Id == expected.Key && i._account_Id == expected.Value
                    && i.institutionId == PlaidUtil.PayPalInstitutionId))
                    throw new InvalidOperationException("PayPal Item/account binding changed during retrieval; retry required.");
            foreach (var expected in plan.ExpectedBankRecords)
            {
                var record = db.Queryable<Record>().First(r => r.Id == expected.Key);
                if (record is null || CombinedUtil.PayPalBankRecordFingerprint(record) != expected.Value)
                    throw new InvalidOperationException("PayPal bank counterpart changed during retrieval; retry required.");
            }
            foreach (var expected in plan.ExpectedBalances)
            {
                var account = db.Queryable<Account>().First(a => a.Id == expected.Key.AccountId);
                if (account is null || GetAccountBalance(account, expected.Key.Currency).v != expected.Value)
                    throw new InvalidOperationException("PayPal balance changed during retrieval; retry required.");
            }
            foreach (var record in plan.Records)
                if (db.Queryable<Record>().Any(r => r.Source == record.Source || r.Source.StartsWith(record.Source + ";")))
                    throw new InvalidOperationException("PayPal record already exists without matching source state.");
            var statementIds = new List<int>();
            foreach (var import in imports.Where(i => i.Key != i.PreviousKey))
            {
                // The originating receipt owns all its records, including a Nexus or refund counterpart.
                var records = plan.Records.Where(r => r.Source.StartsWith($"{CombinedUtil.PayPalPrefix}{import.AccountId}/", StringComparison.Ordinal)).ToList();
                var statementId = SaveStatementImportCore(import.Provider, DateTime.Today, import.Key,
                    records, [], [], false);
                if (!statementId.HasValue) throw new InvalidOperationException("PayPal statement already exists; transaction rolled back.");
                statementIds.Add(statementId.Value);
            }
            AppendRecordSourceSupplements(plan.Supplements);
            foreach (var pair in plan.Pairs)
                {
                    var left = db.Queryable<Record>().Single(r => r.Source == pair.LeftSource || r.Source.StartsWith(pair.LeftSource + ";"));
                    var right = pair.BankRecordId.HasValue
                        ? db.Queryable<Record>().Single(r => r.Id == pair.BankRecordId.Value)
                        : db.Queryable<Record>().Single(r => r.Source == pair.RightSource || r.Source.StartsWith(pair.RightSource + ";"));
                    if (left is null || right is null || left._account_Id == right._account_Id || left.t != right.t || left.v != -right.v
                        || !left.isInternal || !right.isInternal || left.backup is not null || right.backup is not null
                        || left.matchedRecordId.HasValue != right.matchedRecordId.HasValue
                        || (left.matchedRecordId.HasValue && left.matchedRecordId != right.Id)
                        || (right.matchedRecordId.HasValue && right.matchedRecordId != left.Id))
                        throw new InvalidOperationException("PayPal transfer pairing changed; transaction rolled back.");
                    MatchInternalTransferPair(left, right, "PayPalCombinedSourceEvidence");
                }
            ApplyAutomaticExpenseAllocationForStatements(statementIds);
            ProcessAllocatedExpenseDirtyRecordsCore();
        });
    }
}
