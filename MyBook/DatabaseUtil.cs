using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Text.Json;

namespace MyBook
{
    class DatabaseUtil
    {
        private const string DefaultConnectionString = "server=localhost;port=3306;database=mybook;uid=root;pwd=;charset=utf8mb4;";
        private const string DatabaseWriteLockName = "MyBook.DatabaseWrite";
        private const int DatabaseWriteLockTimeoutSeconds = 300;
        private const int CurrentSnapshotSchemaVersion = 1;
        private const int InternalTransferMatchWindowDays = 14;
        private readonly SqlSugarClient db;
        private static readonly JsonSerializerOptions SnapshotJsonOptions = new(JsonSerializerDefaults.Web);
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(AccountInternalId), typeof(AccountBalance), typeof(Record), typeof(Holding), typeof(Finance), typeof(Snapshot), typeof(SnapshotItem), typeof(StatementImport)];
        private static readonly ForeignKeyDefinition[] ForeignKeys =
        [
            new("fk_Accounts_primaryAccount", "Accounts", "_primaryAccount_Id", "Accounts", "Id"),
            new("fk_AccountInternalIds_account", "AccountInternalIds", "_account_Id", "Accounts", "Id"),
            new("fk_AccountBalances_account", "AccountBalances", "_account_Id", "Accounts", "Id"),
            new("fk_Holdings_account", "Holdings", "_account_Id", "Accounts", "Id"),
            new("fk_Records_account", "Records", "_account_Id", "Accounts", "Id"),
            new("fk_Records_statementImport", "Records", "_statementImport_Id", "StatementImports", "Id"),
            new("fk_Records_matchedRecord", "Records", "matchedRecordId", "Records", "Id"),
            new("fk_SnapshotItems_snapshot", "SnapshotItems", "_snapshot_Id", "Snapshots", "Id"),
            new("fk_SnapshotItems_account", "SnapshotItems", "_account_Id", "Accounts", "Id")
        ];

        public DatabaseUtil(IConfigurationRoot config)
        {
            var connectionString = config["database_connection"]
                ?? config.GetConnectionString("Default")
                ?? DefaultConnectionString;

            db = new SqlSugarClient(new ConnectionConfig
            {
                ConnectionString = connectionString,
                DbType = DbType.MySql,
                InitKeyType = InitKeyType.Attribute,
                IsAutoCloseConnection = true
            });
            ValidateSchema();
            ValidateForeignKeys();
            ValidateAccountPrimaryRelations();
        }

        private T ExecuteLockedTransaction<T>(Func<T> action)
        {
            db.Ado.BeginTran();
            var lockTaken = false;
            try
            {
                AcquireDatabaseWriteLock();
                lockTaken = true;
                var result = action();
                db.Ado.CommitTran();
                return result;
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
            finally
            {
                if (lockTaken)
                    TryReleaseDatabaseWriteLock();
            }
        }

        private void ExecuteLockedTransaction(Action action)
        {
            ExecuteLockedTransaction(() =>
            {
                action();
                return true;
            });
        }

        public System.Data.DataTable QueryDebugSql(string sql)
        {
            if (String.IsNullOrWhiteSpace(sql))
                throw new ArgumentException("Debug SQL is empty.", nameof(sql));

            return db.Ado.GetDataTable(sql);
        }

        public int ExecuteDebugSql(string sql)
        {
            if (String.IsNullOrWhiteSpace(sql))
                throw new ArgumentException("Debug SQL is empty.", nameof(sql));

            return ExecuteLockedTransaction(() => db.Ado.ExecuteCommand(sql));
        }

        private void AcquireDatabaseWriteLock()
        {
            var result = db.Ado.GetInt(
                "select get_lock(@lockName, @timeoutSeconds)",
                new SugarParameter("@lockName", DatabaseWriteLockName),
                new SugarParameter("@timeoutSeconds", DatabaseWriteLockTimeoutSeconds));
            if (result != 1)
                throw new TimeoutException($"Timed out waiting for database write lock: {DatabaseWriteLockName}");
        }

        private void TryReleaseDatabaseWriteLock()
        {
            try
            {
                db.Ado.GetInt(
                    "select release_lock(@lockName)",
                    new SugarParameter("@lockName", DatabaseWriteLockName));
            }
            catch (Exception e)
            {
                Console.WriteLine($"release database write lock fail: {e.Message}");
            }
        }

        public bool IsStatementImported(StatementImportProvider provider, DateTime time)
        {
            return IsStatementImported(provider, time, "");
        }

        public bool IsStatementImported(StatementImportProvider provider, DateTime time, string statementKey)
        {
            var importTime = NormalizeStatementImportTime(time);
            return db.Queryable<StatementImport>()
                .Any(it => it.provider == provider && it.time == importTime && it.statementKey == statementKey);
        }

        public bool IsStatementKeyImported(StatementImportProvider provider, string statementKey)
        {
            return db.Queryable<StatementImport>()
                .Any(it => it.provider == provider && it.statementKey == statementKey);
        }

        public DateTime? GetLatestStatementImportTime(StatementImportProvider provider)
        {
            var latestImport = db.Queryable<StatementImport>()
                .Where(it => it.provider == provider)
                .OrderByDescending(it => it.time)
                .First();
            return latestImport?.time;
        }

        public string? GetLatestStatementImportKey(StatementImportProvider provider)
        {
            var latestImport = db.Queryable<StatementImport>()
                .Where(it => it.provider == provider && it.statementKey != "")
                .OrderByDescending(it => it.statementKey)
                .First();
            return latestImport?.statementKey;
        }

        public bool SaveStatementRecordsOnce(
            StatementImportProvider provider,
            DateTime time,
            IEnumerable<Record> records,
            IEnumerable<AccountBalance>? accountBalances = null,
            string statementKey = "",
            IEnumerable<AccountBalance>? beginningAccountBalances = null,
            Action<int>? afterSaveInTransaction = null)
        {
            var recordList = records.ToList();
            var accountBalanceList = accountBalances?.ToList() ?? [];
            var beginningAccountBalanceList = beginningAccountBalances?.ToList() ?? [];
            try
            {
                return ExecuteLockedTransaction(() =>
                {
                    var statementImportId = SaveStatementImportCore(
                        provider,
                        time,
                        statementKey,
                        recordList,
                        accountBalanceList,
                        beginningAccountBalanceList,
                        ShouldValidateBeginningAccountBalances(provider),
                        afterSaveInTransaction: afterSaveInTransaction);
                    if (!statementImportId.HasValue)
                        return false;

                    MatchKnownInternalTransfersForStatements([statementImportId.Value]);
                    MatchInternalTransfersAroundStatement(statementImportId.Value);
                    return true;
                });
            }
            catch (Exception e)
            {
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time, statementKey))
                    return false;
                throw;
            }
        }

        public bool SaveOutOfOrderStatementRecordsOnce(
            StatementImportProvider provider,
            DateTime time,
            IEnumerable<Record> records,
            IEnumerable<AccountBalance>? accountBalances = null,
            string statementKey = "",
            IEnumerable<AccountInternalId>? internalCardNos = null,
            Action<int>? afterSaveInTransaction = null)
        {
            var recordList = records.ToList();
            var accountBalanceList = accountBalances?.ToList() ?? [];
            var internalCardNoList = internalCardNos?.ToList() ?? [];
            try
            {
                return ExecuteLockedTransaction(() =>
                {
                    if (IsStatementImported(provider, time, statementKey))
                        return false;

                    var statementImportId = InsertStatementImport(provider, time, statementKey);
                    if (accountBalanceList.Count > 0)
                        SaveAccountBalancesCore(accountBalanceList);
                    SaveRecordsCore(recordList, statementImportId);
                    if (internalCardNoList.Count > 0)
                        EnsureAccountInternalCardNos(internalCardNoList);

                    afterSaveInTransaction?.Invoke(statementImportId);
                    MatchKnownInternalTransfersForStatements([statementImportId]);
                    MatchInternalTransfersAroundStatement(statementImportId);
                    return true;
                });
            }
            catch (Exception e)
            {
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time, statementKey))
                    return false;
                throw;
            }
        }

        public List<bool> SaveStatementRecordsAndHoldingsOnce(IEnumerable<StatementRecordHoldingImport> imports)
        {
            var importList = imports.ToList();
            var saved = new List<bool>();
            return ExecuteLockedTransaction(() =>
            {
                var savedStatementImportIds = new List<int>();
                var shouldValidateBeginningBalances = importList
                    .Select(import => import.Provider)
                    .Distinct()
                    .ToDictionary(provider => provider, ShouldValidateBeginningAccountBalances);

                foreach (var import in importList)
                {
                    var statementImportId = SaveStatementImportCore(
                        import.Provider,
                        import.Time,
                        import.StatementKey,
                        import.Records,
                        import.AccountBalances,
                        import.BeginningAccountBalances,
                        shouldValidateBeginningBalances[import.Provider],
                        import.HoldingAccount,
                        import.Holdings,
                        import.InternalCardNos);
                    if (!statementImportId.HasValue)
                    {
                        saved.Add(false);
                        continue;
                    }

                    savedStatementImportIds.Add(statementImportId.Value);
                    saved.Add(true);
                }

                MatchKnownInternalTransfersForStatements(savedStatementImportIds);
                foreach (var statementImportId in savedStatementImportIds)
                    MatchInternalTransfersAroundStatement(statementImportId);

                return saved;
            });
        }

        private int? SaveStatementImportCore(
            StatementImportProvider provider,
            DateTime time,
            string statementKey,
            List<Record> records,
            List<AccountBalance> accountBalances,
            List<AccountBalance> beginningAccountBalances,
            bool shouldValidateBeginningBalances,
            Account? holdingAccount = null,
            List<Holding>? holdings = null,
            List<AccountInternalId>? internalCardNos = null,
            Action<int>? afterSaveInTransaction = null)
        {
            if (IsStatementImported(provider, time, statementKey))
                return null;

            var hasExternalBalances = accountBalances.Count > 0 || beginningAccountBalances.Count > 0;
            if (hasExternalBalances)
            {
                ValidateBeginningAccountBalances(
                    provider,
                    beginningAccountBalances,
                    shouldValidateBeginningBalances);
                ValidateRecordBalanceChanges(provider, records, beginningAccountBalances, accountBalances);
            }
            else
            {
                ValidateRelativeBalanceRecords(provider, records);
            }

            var statementImportId = InsertStatementImport(provider, time, statementKey);
            if (hasExternalBalances)
                SaveAccountBalancesCore(accountBalances);

            if (holdingAccount is not null && holdings is not null)
                SaveAccountHoldingsCore(holdingAccount, holdings);

            SaveRecordsCore(records, statementImportId);
            if (!hasExternalBalances)
                AddAccountBalanceDeltas(records);

            if (internalCardNos is not null)
                EnsureAccountInternalCardNos(internalCardNos);

            afterSaveInTransaction?.Invoke(statementImportId);
            return statementImportId;
        }

        public List<RecordDetailData> GetRecordDetails(DateTime start, DateTime end, string? accountName)
        {
            var startDate = start.Date;
            var endDate = end.Date.AddDays(1);
            var query = db.Queryable<Record>()
                .Where(record => record.date >= startDate && record.date < endDate);
            if (!String.IsNullOrWhiteSpace(accountName))
            {
                var account = GetPostingAccount(GetAccountByName(accountName));
                query = query.Where(record => record._account_Id == account.Id);
            }

            var records = query.ToList()
                .OrderBy(record => record.date)
                .ThenBy(record => record.Id)
                .ToList();
            var accounts = db.Queryable<Account>()
                .ToList()
                .ToDictionary(account => account.Id);
            var statementImports = db.Queryable<StatementImport>()
                .ToList()
                .ToDictionary(statementImport => statementImport.Id);
            return records
                .Select(record =>
                {
                    accounts.TryGetValue(record._account_Id, out var account);
                    statementImports.TryGetValue(record._statementImport_Id, out var statementImport);
                    return new RecordDetailData
                    {
                        Id = record.Id,
                        AccountName = account?.name ?? "",
                        Amount = record.v,
                        Currency = record.t,
                        Date = record.date,
                        DestAccount = record.DestAccount,
                        IsInternal = record.isInternal,
                        MatchedRecordId = record.matchedRecordId,
                        MatchedRecordReason = record.matchedRecordReason,
                        IsRefundMatched = record.isRefundMatched,
                        HoldingQuantity = record.HoldingQuantity,
                        Source = record.Source,
                        Reason = record.Reason,
                        StatementProvider = statementImport?.provider ?? StatementImportProvider.Manual,
                        StatementKey = statementImport?.statementKey ?? "",
                        HasBackup = !String.IsNullOrWhiteSpace(record.backup)
                    };
                })
                .ToList();
        }

        public List<string> GetAccountNames()
        {
            return db.Queryable<Account>()
                .ToList()
                .Select(account => account.name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public List<AccountBalanceDetailData> GetAccountBalanceDetails(string accountName)
        {
            if (String.IsNullOrWhiteSpace(accountName))
                return [];

            var account = GetPostingAccount(GetAccountByName(accountName));
            var balances = db.Queryable<AccountBalance>()
                .Where(balance => balance._account_Id == account.Id)
                .ToList()
                .ToDictionary(balance => balance.t, balance => balance.v);

            return Enum.GetValues<CurrencyType>()
                .Select(currency => new AccountBalanceDetailData
                {
                    AccountName = account.name,
                    Currency = currency,
                    Amount = balances.TryGetValue(currency, out var amount) ? amount : 0
                })
                .ToList();
        }

        public void SaveRecordDetails(IEnumerable<RecordDetailEdit> edits)
        {
            var editList = edits.ToList();
            ExecuteLockedTransaction(() =>
            {
                var accountsByName = db.Queryable<Account>()
                    .ToList()
                    .ToDictionary(account => account.name, StringComparer.OrdinalIgnoreCase);
                var manualStatementImportId = 0;
                var now = DateTime.Now;
                foreach (var edit in editList)
                {
                    var account = ResolveRecordDetailAccount(edit, accountsByName);
                    if (edit.Id <= 0)
                    {
                        manualStatementImportId = EnsureManualStatementImport(manualStatementImportId, now);
                        InsertManualRecordDetail(edit, account, manualStatementImportId, now);
                        continue;
                    }

                    UpdateRecordDetail(edit, account, now, ref manualStatementImportId);
                }
            });
        }

        private Account ResolveRecordDetailAccount(
            RecordDetailEdit edit,
            Dictionary<string, Account> accountsByName)
        {
            if (String.IsNullOrWhiteSpace(edit.AccountName))
                throw new InvalidOperationException("Record account is required.");
            if (!accountsByName.TryGetValue(CleanRecordText(edit.AccountName), out var account))
                throw new InvalidOperationException($"Account not found: {edit.AccountName}");

            return GetPostingAccount(account);
        }

        private int EnsureManualStatementImport(int currentStatementImportId, DateTime now)
        {
            return currentStatementImportId > 0
                ? currentStatementImportId
                : InsertStatementImport(StatementImportProvider.Manual, now, $"manual-edit-{now:yyyyMMddHHmmssffffff}");
        }

        private void InsertManualRecordDetail(
            RecordDetailEdit edit,
            Account account,
            int statementImportId,
            DateTime now)
        {
            var record = new Record
            {
                v = edit.Amount,
                t = edit.Currency,
                date = NormalizeRecordDetailDate(edit.Date),
                updateTime = now,
                DestAccount = CleanRecordText(edit.DestAccount),
                isInternal = edit.IsInternal,
                isRefundMatched = edit.IsRefundMatched,
                HoldingQuantity = edit.HoldingQuantity,
                Source = CleanRecordText(edit.Source),
                Reason = CleanRecordText(edit.Reason),
                _account_Id = account.Id,
                _statementImport_Id = statementImportId,
                backup = null
            };
            db.Insertable(record).ExecuteCommand();
            AddAccountBalanceDelta(account.Id, record.t, record.v);
        }

        private void UpdateRecordDetail(
            RecordDetailEdit edit,
            Account account,
            DateTime now,
            ref int manualStatementImportId)
        {
            var existing = db.Queryable<Record>()
                .Where(record => record.Id == edit.Id)
                .First();
            if (existing is null)
                throw new InvalidOperationException($"Record not found: {edit.Id}");

            var statementImport = db.Queryable<StatementImport>()
                .Where(import => import.Id == existing._statementImport_Id)
                .First();
            if (statementImport is null)
                throw new InvalidOperationException($"Record statement import not found: {existing._statementImport_Id}");

            var normalizedDate = NormalizeRecordDetailDate(edit.Date);
            var normalizedDestAccount = CleanRecordText(edit.DestAccount);
            var normalizedSource = CleanRecordText(edit.Source);
            var normalizedReason = CleanRecordText(edit.Reason);
            if (existing._account_Id != account.Id)
                throw new InvalidOperationException($"Record account is fixed and cannot be edited: {existing.Id}");
            if (statementImport.provider != StatementImportProvider.Manual && existing.date != normalizedDate)
                throw new InvalidOperationException($"Imported record date is fixed and cannot be edited: {existing.Id}");
            if (statementImport.provider != StatementImportProvider.Manual && existing.t != edit.Currency)
                throw new InvalidOperationException($"Imported record currency is fixed and cannot be edited: {existing.Id}");
            if (existing.HoldingQuantity == 0 && edit.HoldingQuantity != 0)
                throw new InvalidOperationException($"Record holding quantity is not applicable and cannot be edited: {existing.Id}");

            var changed = existing._account_Id != account.Id
                || existing.v != edit.Amount
                || existing.t != edit.Currency
                || existing.date != normalizedDate
                || existing.DestAccount != normalizedDestAccount
                || existing.isInternal != edit.IsInternal
                || existing.isRefundMatched != edit.IsRefundMatched
                || existing.HoldingQuantity != edit.HoldingQuantity
                || existing.Source != normalizedSource
                || existing.Reason != normalizedReason;
            if (!changed)
                return;

            manualStatementImportId = EnsureManualStatementImport(manualStatementImportId, now);
            if (statementImport.provider != StatementImportProvider.Manual && String.IsNullOrWhiteSpace(existing.backup))
                existing.backup = SerializeRecordBackup(existing, statementImport);

            AddAccountBalanceDelta(existing._account_Id, existing.t, -existing.v);
            existing._account_Id = account.Id;
            existing._statementImport_Id = manualStatementImportId;
            existing.v = edit.Amount;
            existing.t = edit.Currency;
            existing.date = normalizedDate;
            existing.updateTime = now;
            existing.DestAccount = normalizedDestAccount;
            existing.isInternal = edit.IsInternal;
            existing.isRefundMatched = edit.IsRefundMatched;
            existing.HoldingQuantity = edit.HoldingQuantity;
            existing.Source = normalizedSource;
            existing.Reason = normalizedReason;
            db.Updateable(existing).ExecuteCommand();
            AddAccountBalanceDelta(account.Id, existing.t, existing.v);
        }

        private void AddAccountBalanceDelta(int accountId, CurrencyType currency, decimal delta)
        {
            if (delta == 0)
                return;

            var existing = db.Queryable<AccountBalance>()
                .Where(balance => balance._account_Id == accountId && balance.t == currency)
                .First();
            if (existing is null)
            {
                db.Insertable(new AccountBalance
                {
                    _account_Id = accountId,
                    t = currency,
                    v = delta
                }).ExecuteCommand();
                return;
            }

            existing.v += delta;
            if (existing.v == 0)
            {
                db.Deleteable<AccountBalance>()
                    .Where(balance => balance.Id == existing.Id)
                    .ExecuteCommand();
                return;
            }

            db.Updateable(existing).ExecuteCommand();
        }

        private void AddAccountBalanceDeltas(IEnumerable<Record> records)
        {
            foreach (var group in records
                         .GroupBy(record => (record._account_Id, record.t))
                         .Select(group => new
                         {
                             AccountId = group.Key._account_Id,
                             Currency = group.Key.t,
                             Amount = group.Sum(record => record.v)
                         }))
            {
                AddAccountBalanceDelta(group.AccountId, group.Currency, group.Amount);
            }
        }

        private void MatchInternalTransfersAroundStatement(int statementImportId)
        {
            var importedRecords = GetRecordsByStatementImport(statementImportId);
            if (importedRecords.Count == 0)
                return;

            var start = importedRecords.Min(record => record.date).AddDays(-InternalTransferMatchWindowDays);
            var end = importedRecords.Max(record => record.date).AddDays(InternalTransferMatchWindowDays);
            var records = db.Queryable<Record>()
                .Where(record => record.matchedRecordId == null
                    && !record.isRefundMatched
                    && record.date >= start
                    && record.date <= end)
                .ToList();
            if (records.Count < 2)
                return;

            var accountsByName = db.Queryable<Account>()
                .ToList()
                .ToDictionary(account => account.name, StringComparer.OrdinalIgnoreCase);

            MatchInternalTransfers(records, accountsByName, requireKnownCounterparty: false);
        }

        private void MatchKnownInternalTransfersForStatements(List<int> statementImportIds)
        {
            if (statementImportIds.Count == 0)
                return;

            var records = db.Queryable<Record>()
                .Where(record => statementImportIds.Contains(record._statementImport_Id)
                    && record.matchedRecordId == null
                    && !record.isRefundMatched)
                .ToList();
            if (records.Count < 2)
                return;

            var accountsByName = db.Queryable<Account>()
                .ToList()
                .ToDictionary(account => account.name, StringComparer.OrdinalIgnoreCase);

            MatchInternalTransfers(records, accountsByName, requireKnownCounterparty: true);
        }

        private void MatchInternalTransfers(
            List<Record> records,
            Dictionary<string, Account> accountsByName,
            bool requireKnownCounterparty)
        {
            foreach (var anchor in records
                         .Where(record => record.isInternal && record.matchedRecordId is null)
                         .OrderBy(record => record.date)
                         .ThenBy(record => record.Id)
                         .ToList())
            {
                if (anchor.matchedRecordId is not null)
                    continue;

                if (TryMatchCurrencyExchange(anchor, records))
                    continue;

                TryMatchInternalTransfer(anchor, records, accountsByName, requireKnownCounterparty);
            }
        }

        private bool TryMatchCurrencyExchange(Record anchor, List<Record> records)
        {
            var conversionCode = ExtractCurrencyExchangeCode(anchor.Source);
            if (String.IsNullOrWhiteSpace(conversionCode))
                return false;

            var candidates = records
                .Where(record => record.Id != anchor.Id
                    && record.matchedRecordId is null
                    && record.isInternal
                    && !record.isRefundMatched
                    && HasOppositeSigns(anchor.v, record.v)
                    && String.Equals(ExtractCurrencyExchangeCode(record.Source), conversionCode, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (candidates.Count != 1)
                return false;

            MatchInternalTransferPair(anchor, candidates[0], $"SameConversionCode:{conversionCode}");
            return true;
        }

        private bool TryMatchInternalTransfer(
            Record anchor,
            List<Record> records,
            Dictionary<string, Account> accountsByName,
            bool requireKnownCounterparty)
        {
            if (anchor.v == 0)
                return false;

            var targetAccount = ResolveInternalTransferTargetAccount(anchor, accountsByName);
            if (requireKnownCounterparty && targetAccount is null)
                return false;

            var start = anchor.date.AddDays(-InternalTransferMatchWindowDays);
            var end = anchor.date.AddDays(InternalTransferMatchWindowDays);
            var candidates = records
                .Where(record => record.Id != anchor.Id
                    && record.matchedRecordId is null
                    && !record.isRefundMatched
                    && record.isInternal
                    && (targetAccount is null || record._account_Id == targetAccount.Id)
                    && record.t == anchor.t
                    && record.v == -anchor.v
                    && record.date >= start
                    && record.date <= end)
                .ToList();
            if (requireKnownCounterparty)
            {
                candidates = candidates
                    .Where(candidate =>
                    {
                        var candidateTarget = ResolveInternalTransferTargetAccount(candidate, accountsByName);
                        return candidateTarget is not null && candidateTarget.Id == anchor._account_Id;
                    })
                    .ToList();
            }

            if (candidates.Count != 1)
                return false;

            MatchInternalTransferPair(
                anchor,
                candidates[0],
                requireKnownCounterparty ? "KnownCounterpartyAccountAmountDate" : "CounterpartyAccountAmountDate");
            return true;
        }

        private Account? ResolveInternalTransferTargetAccount(
            Record record,
            Dictionary<string, Account> accountsByName)
        {
            if (TryResolveCounterpartyName(record, out var counterpartyName)
                && accountsByName.TryGetValue(counterpartyName, out var namedAccount)
                && !IsUnknownAccount(namedAccount))
            {
                return GetPostingAccount(namedAccount);
            }

            var matchedAccount = FindAccountByInternalCardNoText(
                null,
                $"record {record.Id} internal transfer match",
                record.DestAccount);
            if (matchedAccount is null || IsUnknownAccount(matchedAccount))
                return null;

            return GetPostingAccount(matchedAccount);
        }

        private static bool TryResolveCounterpartyName(Record record, out string counterpartyName)
        {
            counterpartyName = "";
            var destAccount = record.DestAccount.Trim();
            if (String.IsNullOrWhiteSpace(destAccount))
                return false;

            var parts = destAccount.Split("->", 2, StringSplitOptions.TrimEntries);
            if (parts.Length == 2)
            {
                counterpartyName = record.v < 0 ? parts[1] : parts[0];
                return !String.IsNullOrWhiteSpace(counterpartyName);
            }

            counterpartyName = destAccount;
            return true;
        }

        private void MatchInternalTransferPair(Record left, Record right, string reason)
        {
            if (left.Id == right.Id)
                throw new InvalidOperationException($"Record cannot match itself: {left.Id}");
            if (left.matchedRecordId is not null || right.matchedRecordId is not null)
                return;

            var updateTime = DateTime.Now;
            left.matchedRecordId = right.Id;
            left.matchedRecordReason = reason;
            left.updateTime = updateTime;
            right.matchedRecordId = left.Id;
            right.matchedRecordReason = reason;
            right.updateTime = updateTime;

            db.Updateable(left)
                .UpdateColumns(record => new { record.matchedRecordId, record.matchedRecordReason, record.updateTime })
                .ExecuteCommand();
            db.Updateable(right)
                .UpdateColumns(record => new { record.matchedRecordId, record.matchedRecordReason, record.updateTime })
                .ExecuteCommand();
        }

        private static bool HasOppositeSigns(decimal left, decimal right)
        {
            return left != 0 && right != 0 && Math.Sign(left) == -Math.Sign(right);
        }

        private static string ExtractCurrencyExchangeCode(string source)
        {
            var recordSourceCode = ExtractRecordSourceCode(source);
            if (recordSourceCode.StartsWith("BALANCE-", StringComparison.OrdinalIgnoreCase))
                return recordSourceCode;

            var ocbcFxMatch = Regex.Match(
                source ?? "",
                @"\bFX\s+Transaction\s*/\s*(?<code>[A-Za-z0-9_-]+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (ocbcFxMatch.Success)
                return $"OCBC-FX-{ocbcFxMatch.Groups["code"].Value.Trim()}";

            return "";
        }

        private static string ExtractRecordSourceCode(string source)
        {
            var match = Regex.Match(
                source ?? "",
                @"(?:^|;\s*)code=(?<code>[^;]+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["code"].Value.Trim() : "";
        }

        private static DateTime NormalizeRecordDetailDate(DateTime date)
        {
            if (date.Year < 1900)
                throw new InvalidOperationException($"Invalid record date: {date:O}");
            return date;
        }

        private static string CleanRecordText(string? value)
        {
            return value?.Trim() ?? "";
        }

        private static string LimitRecordText(string text)
        {
            const int maxLength = 1024;
            var normalized = CleanRecordText(text);
            return normalized.Length <= maxLength ? normalized : normalized[..maxLength];
        }

        private static string AppendLimitedRecordText(string current, string append)
        {
            const int maxLength = 1024;
            var normalizedAppend = CleanRecordText(append);
            if (String.IsNullOrWhiteSpace(normalizedAppend))
                return LimitRecordText(current);
            if (normalizedAppend.Length >= maxLength)
                return normalizedAppend[..maxLength];

            var normalizedCurrent = CleanRecordText(current);
            var separator = String.IsNullOrWhiteSpace(normalizedCurrent) ? "" : "; ";
            var availableCurrentLength = maxLength - normalizedAppend.Length - separator.Length;
            if (availableCurrentLength <= 0)
                return normalizedAppend;
            if (normalizedCurrent.Length > availableCurrentLength)
                normalizedCurrent = normalizedCurrent[..availableCurrentLength];

            return $"{normalizedCurrent}{separator}{normalizedAppend}";
        }

        private static string SerializeRecordBackup(Record record, StatementImport statementImport)
        {
            return JsonSerializer.Serialize(
                new RecordBackupPayloadV1(
                    1,
                    record.Id,
                    record._account_Id,
                    record.v,
                    record.t,
                    record.date,
                    record.postingDate,
                    record.updateTime,
                    record.DestAccount,
                    record.isInternal,
                    record.matchedRecordId,
                    record.matchedRecordReason,
                    record.isRefundMatched,
                    record.HoldingQuantity,
                    record.Source,
                    record.Reason,
                    record._descCurrency_v,
                    record._descCurrency_t,
                    record._statementImport_Id,
                    statementImport.provider,
                    statementImport.statementKey),
                SnapshotJsonOptions);
        }

        public Account GetAccountByTypeAndId(string? accountType, string id)
        {
            return GetAccountByName(BuildAccountName(accountType, id));
        }

        public Account GetAccountByName(string accountName)
        {
            var account = FindAccountByName(accountName);
            if (account is null)
                throw new InvalidOperationException($"Account not found: {accountName}");

            return account;
        }

        public List<Account> GetAllAccounts()
        {
            return db.Queryable<Account>().ToList();
        }

        public void EnsureAccountInternalCardNos(IEnumerable<AccountInternalId> internalIds)
        {
            var uniqueInternalIds = new Dictionary<string, AccountInternalId>(StringComparer.OrdinalIgnoreCase);
            foreach (var internalId in internalIds)
            {
                if (internalId.Account is null)
                    throw new InvalidOperationException("Account internal id account is required.");

                var cardNo = NormalizeInternalCardNoForStorage(internalId.cardNo);
                if (String.IsNullOrWhiteSpace(cardNo))
                    continue;

                var key = $"{internalId.Account.name}|{cardNo}|{internalId.currencyType?.ToString() ?? ""}";
                if (!uniqueInternalIds.ContainsKey(key))
                    uniqueInternalIds.Add(key, internalId);
            }

            foreach (var internalId in uniqueInternalIds.Values)
                EnsureAccountInternalCardNo(internalId);
        }

        public void EnsureAccountInternalCardNo(AccountInternalId internalId)
        {
            if (internalId.Account is null)
                throw new InvalidOperationException("Account internal id account is required.");

            var cardNo = NormalizeInternalCardNoForStorage(internalId.cardNo);
            if (String.IsNullOrWhiteSpace(cardNo))
                return;

            var account = GetExistingAccountByName(internalId.Account);
            var existing = db.Queryable<AccountInternalId>()
                .Where(it => it._account_Id == account.Id && it.cardNo == cardNo)
                .First();
            if (existing is not null)
            {
                LogAccountInternalCardNoKnown("already exists", internalId, account, cardNo);
                return;
            }

            internalId.Account = account;
            internalId._account_Id = account.Id;
            internalId.cardNo = cardNo;
            db.Insertable(internalId).ExecuteCommand();
            LogAccountInternalCardNoKnown("stored", internalId, account, cardNo);
        }

        public Account? FindAccountByInternalCardNoText(string? preferredAccountType, string? matchContext, params string?[] texts)
        {
            var usefulTexts = texts
                .Where(text => !String.IsNullOrWhiteSpace(text))
                .Select(text => text!)
                .ToList();
            if (usefulTexts.Count == 0)
                return null;

            var accounts = db.Queryable<Account>()
                .ToList()
                .ToDictionary(account => account.Id);
            var accountTypes = accounts.Values
                .Select(account => GetAccountType(account.name))
                .Where(type => !String.IsNullOrWhiteSpace(type))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var explicitTypes = ExtractAccountTypesFromTexts(usefulTexts, accountTypes);
            var matches = db.Queryable<AccountInternalId>()
                .ToList()
                .Where(internalId => accounts.ContainsKey(internalId._account_Id))
                .Select(internalId => new
                {
                    InternalId = internalId,
                    Account = accounts[internalId._account_Id],
                    AccountType = GetAccountType(accounts[internalId._account_Id].name),
                    Match = FindInternalCardNoMatch(internalId.cardNo, usefulTexts)
                })
                .Where(item => item.Match is not null)
                .Select(item => new InternalCardNoCandidate(item.InternalId, item.Account, item.AccountType, item.Match!))
                .ToList();
            if (matches.Count == 0)
                return null;

            var typedMatches = matches;
            if (explicitTypes.Count > 0)
                typedMatches = matches
                    .Where(item => explicitTypes.Contains(item.AccountType, StringComparer.OrdinalIgnoreCase))
                    .ToList();
            else if (!String.IsNullOrWhiteSpace(preferredAccountType))
            {
                var preferredMatches = matches
                    .Where(item => String.Equals(item.AccountType, preferredAccountType, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (preferredMatches.Count > 0)
                    typedMatches = preferredMatches;
            }

            var prioritizedMatches = PreferKnownAccountMatches(typedMatches);
            var matchedAccounts = prioritizedMatches
                .Select(item => item.Account)
                .GroupBy(account => account.Id)
                .Select(group => group.First())
                .ToList();
            if (matchedAccounts.Count > 1)
            {
                throw new InvalidOperationException(
                    $"Ambiguous internal account id match: {String.Join(", ", matchedAccounts.Select(account => account.name))}; text={String.Join(" / ", usefulTexts)}");
            }

            if (matchedAccounts.Count == 1)
                LogInternalCardNoMatch(matchContext, prioritizedMatches.First(), usefulTexts);

            return matchedAccounts.FirstOrDefault();
        }

        public Currency GetAccountBalance(Account account, CurrencyType currencyType)
        {
            account = GetPostingAccount(account);
            var balance = db.Queryable<AccountBalance>()
                .Where(it => it._account_Id == account.Id && it.t == currencyType)
                .First();
            return new Currency(balance?.v ?? 0, currencyType);
        }

        private Account? FindAccountByName(string accountName)
        {
            return db.Queryable<Account>()
                .Where(it => it.name == accountName)
                .First();
        }

        public static string BuildAccountName(string? accountType, string id)
        {
            var parts = new[] { accountType, id }
                .Where(part => !String.IsNullOrWhiteSpace(part))
                .Select(part => part!.Trim());
            return String.Join("_", parts);
        }

        private Account GetExistingAccountByName(Account account)
        {
            var existing = GetAccountByName(account.name);
            account.Id = existing.Id;
            return existing;
        }

        public Account GetPostingAccount(Account account)
        {
            var current = GetExistingAccountByName(account);
            if (current._primaryAccount_Id is null)
                return current;

            if (current._primaryAccount_Id.Value == current.Id)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name} points to itself");

            var primary = db.Queryable<Account>()
                .Where(it => it.Id == current._primaryAccount_Id.Value)
                .First();
            if (primary is null)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name} points to missing account {current._primaryAccount_Id.Value}");
            if (primary._primaryAccount_Id is not null)
                throw new InvalidOperationException($"Invalid account primary relation: primary account {primary.name} is also a supplementary account");

            return primary;
        }

        private int InsertStatementImport(StatementImportProvider provider, DateTime time, string statementKey)
        {
            return db.Insertable(new StatementImport
            {
                provider = provider,
                time = NormalizeStatementImportTime(time),
                statementKey = statementKey
            }).ExecuteReturnIdentity();
        }

        private void SaveRecordsCore(List<Record> recordList, int statementImportId)
        {
            if (statementImportId <= 0)
                throw new InvalidOperationException("Record statement import is required.");

            foreach (var record in recordList)
            {
                if (record.Account is null)
                    throw new InvalidOperationException("Record account is required.");

                var account = GetPostingAccount(record.Account);
                record.Account = account;
                record._account_Id = account.Id;
                record._statementImport_Id = statementImportId;
            }

            if (recordList.Count > 0)
                db.Insertable(recordList).ExecuteCommand();
        }

        private void ValidateRelativeBalanceRecords(StatementImportProvider provider, List<Record> recordList)
        {
            foreach (var record in recordList)
            {
                if (record.Account is null)
                    throw new InvalidOperationException("Record account is required.");

                var account = GetPostingAccount(record.Account);
                if (!account.relativeBalance)
                {
                    throw new InvalidOperationException(
                        $"Statement import without external balances is only allowed for relative-balance accounts: provider={provider}, account={account.name}");
                }
            }
        }

        private bool ShouldValidateBeginningAccountBalances(StatementImportProvider provider)
        {
            return db.Queryable<StatementImport>()
                .Where(it => it.provider == provider)
                .Count() > 1;
        }

        private void ValidateBeginningAccountBalances(
            StatementImportProvider provider,
            List<AccountBalance> beginningAccountBalances,
            bool shouldValidate)
        {
            if (!shouldValidate || beginningAccountBalances.Count == 0)
                return;

            foreach (var beginningAccountBalance in beginningAccountBalances)
            {
                if (beginningAccountBalance.Account is null)
                    throw new InvalidOperationException("Beginning account balance account is required.");

                var account = GetPostingAccount(beginningAccountBalance.Account);
                var existing = db.Queryable<AccountBalance>()
                    .Where(it => it._account_Id == account.Id && it.t == beginningAccountBalance.t)
                    .First();
                if (existing is null)
                {
                    if (beginningAccountBalance.v == 0 || !HasAccountCurrencyHistory(account.Id, beginningAccountBalance.t))
                        continue;

                    throw new InvalidOperationException(
                        $"Missing current account balance for {provider}: {account.name}/{beginningAccountBalance.t}");
                }

                if (existing.v != beginningAccountBalance.v)
                {
                    throw new InvalidOperationException(
                        $"Beginning account balance mismatch for {provider}: {account.name}/{beginningAccountBalance.t}, current={existing.v}, beginning={beginningAccountBalance.v}");
                }
            }
        }

        private bool HasAccountCurrencyHistory(int accountId, CurrencyType currencyType)
        {
            return db.Queryable<AccountBalance>()
                    .Where(it => it._account_Id == accountId && it.t == currencyType)
                    .Any()
                || db.Queryable<Record>()
                    .Where(it => it._account_Id == accountId && it.t == currencyType)
                    .Any();
        }

        private void ValidateRecordBalanceChanges(
            StatementImportProvider provider,
            List<Record> records,
            List<AccountBalance> beginningAccountBalances,
            List<AccountBalance> endingAccountBalances)
        {
            if (beginningAccountBalances.Count == 0 || endingAccountBalances.Count == 0)
                throw new InvalidOperationException($"Record balance validation requires beginning and ending account balances for {provider}.");

            var beginningBalances = BuildAccountBalanceMap(beginningAccountBalances, "Beginning account balance");
            var endingBalances = BuildAccountBalanceMap(endingAccountBalances, "Ending account balance");
            var recordChanges = SumRecordAmountsByAccountAndCurrency(records);
            var keys = beginningBalances.Keys
                .Union(endingBalances.Keys)
                .Union(recordChanges.Keys)
                .ToList();

            foreach (var key in keys)
            {
                var beginning = beginningBalances.TryGetValue(key, out var beginningValue) ? beginningValue : 0;
                var ending = endingBalances.TryGetValue(key, out var endingValue) ? endingValue : 0;
                var change = recordChanges.TryGetValue(key, out var changeValue) ? changeValue : 0;
                var expectedEnding = beginning + change;
                if (expectedEnding != ending)
                {
                    throw new InvalidOperationException(
                        $"Record balance mismatch for {provider}: accountId={key.AccountId}, currency={key.Currency}, beginning={beginning}, records={change}, ending={ending}");
                }
            }
        }

        public Dictionary<(int AccountId, CurrencyType Currency), decimal> SumRecordAmountsByAccountAndCurrency(IEnumerable<Record> records)
        {
            var sums = new Dictionary<(int AccountId, CurrencyType Currency), decimal>();
            foreach (var record in records)
            {
                if (record.Account is null)
                    throw new InvalidOperationException("Record account is required.");

                var account = GetPostingAccount(record.Account);
                var key = (account.Id, record.t);
                sums[key] = sums.TryGetValue(key, out var current)
                    ? current + record.v
                    : record.v;
            }

            return sums.ToDictionary(item => item.Key, item => item.Value);
        }

        private Dictionary<(int AccountId, CurrencyType Currency), decimal> BuildAccountBalanceMap(
            List<AccountBalance> accountBalances,
            string context)
        {
            var balances = new Dictionary<(int AccountId, CurrencyType Currency), decimal>();
            foreach (var accountBalance in accountBalances)
            {
                if (accountBalance.Account is null)
                    throw new InvalidOperationException($"{context} account is required.");

                var account = GetPostingAccount(accountBalance.Account);
                balances[(account.Id, accountBalance.t)] = accountBalance.v;
            }

            return balances;
        }

        private void SaveAccountBalancesCore(List<AccountBalance> accountBalances)
        {
            foreach (var accountBalance in accountBalances)
            {
                if (accountBalance.Account is null)
                    throw new InvalidOperationException("Account balance account is required.");

                var account = GetPostingAccount(accountBalance.Account);
                accountBalance.Account = account;
                accountBalance._account_Id = account.Id;

                var existing = db.Queryable<AccountBalance>()
                    .Where(it => it._account_Id == account.Id && it.t == accountBalance.t)
                    .First();
                if (existing is null)
                {
                    db.Insertable(accountBalance).ExecuteCommand();
                    continue;
                }

                existing.v = accountBalance.v;
                db.Updateable(existing).ExecuteCommand();
            }
        }

        public void SaveAccountHoldings(Account account, IEnumerable<Holding> holdings)
        {
            var holdingList = holdings.ToList();
            ExecuteLockedTransaction(() =>
            {
                SaveAccountHoldingsCore(account, holdingList);
            });
        }

        public List<Holding> GetHoldings()
        {
            return db.Queryable<Holding>()
                .Includes(it => it.Account)
                .ToList();
        }

        public List<Finance> GetFinances()
        {
            return db.Queryable<Finance>().ToList();
        }

        public Snapshot CreateDailySnapshot()
        {
            return CreateDailySnapshot(DateTime.Today);
        }

        public Snapshot CreateDailySnapshot(DateTime snapshotDate)
        {
            return CreateSnapshot(
                DateTime.Now,
                SnapshotSource.AutoDaily,
                null,
                revision => BuildDailySnapshotKey(snapshotDate, revision));
        }

        public Snapshot CreateSnapshot(DateTime time, SnapshotSource source = SnapshotSource.Manual, string? snapshotKey = null)
        {
            return CreateSnapshot(
                time,
                source,
                snapshotKey,
                revision => BuildSnapshotKey(source, time, revision));
        }

        private Snapshot CreateSnapshot(
            DateTime time,
            SnapshotSource source,
            string? snapshotKey,
            Func<int, string> defaultKeyBuilder)
        {
            return ExecuteLockedTransaction(() =>
            {
                var revision = GetCurrentStatementImportRevision();
                var key = String.IsNullOrWhiteSpace(snapshotKey)
                    ? defaultKeyBuilder(revision)
                    : snapshotKey.Trim();
                var existing = db.Queryable<Snapshot>()
                    .Where(snapshot => snapshot.source == source && snapshot.snapshotKey == key)
                    .First();
                if (existing is not null)
                    return existing;

                var snapshot = new Snapshot
                {
                    source = source,
                    time = time,
                    schemaVersion = CurrentSnapshotSchemaVersion,
                    maxStatementImportId = revision,
                    effectiveDate = time.Date,
                    snapshotKey = key,
                    createdAt = DateTime.Now
                };
                snapshot.Id = db.Insertable(snapshot).ExecuteReturnIdentity();

                var accounts = db.Queryable<Account>()
                    .ToList()
                    .ToDictionary(account => account.Id);
                var items = new List<SnapshotItem>();

                foreach (var balance in db.Queryable<AccountBalance>().ToList())
                {
                    if (!accounts.TryGetValue(balance._account_Id, out var account))
                        throw new InvalidOperationException($"Snapshot account balance points to missing account: {balance._account_Id}");

                    var payload = new SnapshotAccountBalancePayloadV1(
                        account.Id,
                        account.name,
                        balance.t.ToString(),
                        balance.v);
                    items.Add(new SnapshotItem
                    {
                        Snapshot = snapshot,
                        Account = account,
                        itemType = SnapshotItemType.AccountBalance,
                        stableKey = BuildSnapshotBalanceStableKey(account.name, balance.t),
                        accountName = account.name,
                        currencyType = balance.t,
                        amount = balance.v,
                        payloadJson = JsonSerializer.Serialize(payload, SnapshotJsonOptions),
                        _snapshot_Id = snapshot.Id,
                        _account_Id = account.Id
                    });
                }

                foreach (var holding in db.Queryable<Holding>().ToList())
                {
                    if (!accounts.TryGetValue(holding._account_Id, out var account))
                        throw new InvalidOperationException($"Snapshot holding points to missing account: {holding._account_Id}");

                    var totalPrice = holding.totalPrice;
                    var payload = new SnapshotHoldingPayloadV1(
                        account.Id,
                        account.name,
                        holding.code,
                        holding.holdingType.ToString(),
                        holding.quantity,
                        holding.currentPrice.v,
                        holding.currentPrice.t.ToString(),
                        totalPrice.v,
                        totalPrice.t.ToString(),
                        holding.displayText,
                        holding.desc);
                    items.Add(new SnapshotItem
                    {
                        Snapshot = snapshot,
                        Account = account,
                        itemType = SnapshotItemType.Holding,
                        stableKey = BuildSnapshotHoldingStableKey(account.name, holding.code, holding.holdingType),
                        accountName = account.name,
                        currencyType = totalPrice.t,
                        amount = totalPrice.v,
                        payloadJson = JsonSerializer.Serialize(payload, SnapshotJsonOptions),
                        _snapshot_Id = snapshot.Id,
                        _account_Id = account.Id
                    });
                }

                if (items.Count > 0)
                    db.Insertable(items).ExecuteCommand();

                return snapshot;
            });
        }

        private int GetCurrentStatementImportRevision()
        {
            return db.Queryable<StatementImport>()
                .OrderByDescending(import => import.Id)
                .First()
                ?.Id ?? 0;
        }

        public SnapshotData? GetDailySnapshot(DateTime date)
        {
            var keyPrefix = BuildDailySnapshotKey(date);
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.source == SnapshotSource.AutoDaily)
                .OrderByDescending(it => it.maxStatementImportId)
                .ToList()
                .FirstOrDefault(it => it.snapshotKey == keyPrefix
                    || it.snapshotKey.StartsWith($"{keyPrefix}-r", StringComparison.Ordinal));
            return snapshot is null ? null : ReadSnapshotData(snapshot);
        }

        public SnapshotData? GetLatestSnapshotAtOrBefore(DateTime time)
        {
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.time <= time)
                .OrderByDescending(it => it.time)
                .First();
            return snapshot is null ? null : ReadSnapshotData(snapshot);
        }

        public SnapshotData GetSnapshot(int snapshotId)
        {
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.Id == snapshotId)
                .First();
            if (snapshot is null)
                throw new InvalidOperationException($"Snapshot not found: {snapshotId}");

            return ReadSnapshotData(snapshot);
        }

        private SnapshotData ReadSnapshotData(Snapshot snapshot)
        {
            if (snapshot.schemaVersion != CurrentSnapshotSchemaVersion)
                throw new NotSupportedException($"Unsupported snapshot schema version: {snapshot.schemaVersion}");

            var items = db.Queryable<SnapshotItem>()
                .Where(item => item._snapshot_Id == snapshot.Id)
                .ToList();
            var balances = items
                .Where(item => item.itemType == SnapshotItemType.AccountBalance)
                .Select(ParseSnapshotAccountBalanceV1)
                .ToList();
            var holdings = items
                .Where(item => item.itemType == SnapshotItemType.Holding)
                .Select(ParseSnapshotHoldingV1)
                .ToList();
            return new SnapshotData(snapshot, balances, holdings);
        }

        private static SnapshotAccountBalanceData ParseSnapshotAccountBalanceV1(SnapshotItem item)
        {
            var payload = DeserializeSnapshotPayload<SnapshotAccountBalancePayloadV1>(item);
            return new SnapshotAccountBalanceData(
                payload.AccountId,
                payload.AccountName,
                ParseEnum<CurrencyType>(payload.CurrencyType),
                payload.Amount);
        }

        private static SnapshotHoldingData ParseSnapshotHoldingV1(SnapshotItem item)
        {
            var payload = DeserializeSnapshotPayload<SnapshotHoldingPayloadV1>(item);
            return new SnapshotHoldingData(
                payload.AccountId,
                payload.AccountName,
                payload.Code,
                ParseEnum<HoldingType>(payload.HoldingType),
                payload.Quantity,
                new Currency(payload.PriceAmount, ParseEnum<CurrencyType>(payload.PriceCurrencyType)),
                new Currency(payload.TotalAmount, ParseEnum<CurrencyType>(payload.TotalCurrencyType)),
                payload.DisplayText,
                payload.Description);
        }

        private static T DeserializeSnapshotPayload<T>(SnapshotItem item)
        {
            var payload = JsonSerializer.Deserialize<T>(item.payloadJson, SnapshotJsonOptions);
            if (payload is null)
                throw new InvalidOperationException($"Invalid snapshot payload: {item.Id}");

            return payload;
        }

        private static T ParseEnum<T>(string value) where T : struct, Enum
        {
            if (Enum.TryParse<T>(value, out var parsed))
                return parsed;

            throw new InvalidOperationException($"Invalid snapshot enum value: {typeof(T).Name}.{value}");
        }

        private static string BuildSnapshotKey(SnapshotSource source, DateTime time, int revision)
        {
            return source == SnapshotSource.AutoDaily
                ? BuildDailySnapshotKey(time, revision)
                : time.ToString("O", CultureInfo.InvariantCulture);
        }

        private static string BuildDailySnapshotKey(DateTime date)
        {
            return date.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static string BuildDailySnapshotKey(DateTime date, int revision)
        {
            return $"{BuildDailySnapshotKey(date)}-r{revision}";
        }

        private static string BuildSnapshotBalanceStableKey(string accountName, CurrencyType currencyType)
        {
            return $"{accountName}\t{currencyType}";
        }

        private static string BuildSnapshotHoldingStableKey(string accountName, string code, HoldingType holdingType)
        {
            return $"{accountName}\t{holdingType}\t{code}";
        }

        public DashboardData GetDashboardData(DateTime today)
        {
            var currentMonth = new DateTime(today.Year, today.Month, 1);
            var firstMonth = currentMonth.AddMonths(-11);
            var nextMonth = currentMonth.AddMonths(1);
            var lastMonthStart = currentMonth.AddMonths(-1);
            var lastMonthEnd = currentMonth.AddDays(-1);
            var reasonFirstMonth = lastMonthStart.AddMonths(-11);
            var reasonMonths = Enumerable.Range(0, 12)
                .Select(reasonFirstMonth.AddMonths)
                .ToList();
            var records = db.Queryable<Record>()
                .Where(record => !record.isInternal && record.matchedRecordId == null && !record.isRefundMatched && record.date >= reasonFirstMonth && record.date < nextMonth)
                .ToList();
            var accountList = db.Queryable<Account>().ToList();
            var lifeAccountIds = accountList
                .Where(account => account.usage == AccountUsage.Life)
                .Select(account => account.Id)
                .ToHashSet();
            var lifeRecords = records
                .Where(record => lifeAccountIds.Contains(record._account_Id))
                .ToList();
            var monthlyAccountIds = accountList
                .Where(account => account.usage != AccountUsage.Unknown)
                .Select(account => account.Id)
                .ToHashSet();
            var monthlyRecords = records
                .Where(record => monthlyAccountIds.Contains(record._account_Id))
                .ToList();
            var balances = db.Queryable<AccountBalance>()
                .Where(balance => balance.v != 0)
                .ToList();
            var holdings = db.Queryable<Holding>().ToList();
            var investmentAccountIds = accountList
                .Where(account => account.usage == AccountUsage.Investment)
                .Select(account => account.Id)
                .ToHashSet();
            var investmentRecords = db.Queryable<Record>()
                .Where(record => !record.isInternal && record.matchedRecordId == null && !record.isRefundMatched && record.v > 0 && record.date < today.Date.AddDays(1))
                .ToList()
                .Where(record => investmentAccountIds.Contains(record._account_Id))
                .ToList();
            var accounts = accountList.ToDictionary(account => account.Id, account => account.name);
            var holdingNames = BuildHoldingNames(holdings);
            var exchangeRates = GetCurrencyToRmbRates();
            var assetSummaryDates = BuildAssetSummaryDates(today.Date);
            var assetSummaryBalances = BuildAssetSummaryBalanceSets(assetSummaryDates, balances, today.Date);
            var usedCurrencies = balances
                .Select(balance => balance.t)
                .Union(records.Select(record => record.t))
                .Union(investmentRecords.Select(record => record.t))
                .Union(assetSummaryBalances.SelectMany(point => point.Balances.Select(balance => balance.t)))
                .Distinct()
                .ToList();
            var assetSummaryPoints = BuildAssetSummaryPoints(assetSummaryBalances, records, exchangeRates, today.Date);
            var todayAssetSummary = assetSummaryPoints.LastOrDefault(point => point.IsToday);

            return new DashboardData
            {
                AssetSummaryPoints = assetSummaryPoints,
                CurrencySummaries = BuildCurrencySummaries(
                    balances,
                    records.Where(record => record.date >= currentMonth && record.date < nextMonth).ToList(),
                    exchangeRates),
                MonthlyFlowSeries = BuildMonthlyFlowSeries(monthlyRecords, firstMonth),
                RmbMonthlyFlowSeries = BuildRmbMonthlyFlowSeries(monthlyRecords, firstMonth, exchangeRates),
                MonthlyAccounts = BuildMonthlyFlowAccountStatistics(
                    accountList,
                    records,
                    firstMonth,
                    exchangeRates),
                RmbReasonFlowSeriesByMonth = reasonMonths
                    .Select(month => BuildRmbReasonFlowSeries(lifeRecords, month, month.AddMonths(1), exchangeRates))
                    .ToList(),
                DefaultReasonMonthIndex = Math.Max(0, reasonMonths.FindIndex(month => month == lastMonthStart)),
                InvestmentByReason = BuildInvestmentStatistics(
                    investmentRecords,
                    today,
                    exchangeRates,
                    record => String.IsNullOrWhiteSpace(record.Reason) ? "未分类" : record.Reason),
                InvestmentByHolding = BuildInvestmentStatistics(
                    investmentRecords,
                    today,
                    exchangeRates,
                    record => BuildHoldingInvestmentKey(record, holdingNames)),
                InvestmentAccounts = BuildInvestmentAccountStatistics(
                    accounts,
                    investmentAccountIds,
                    investmentRecords,
                    today,
                    exchangeRates,
                    holdingNames),
                TotalAssetsRmb = todayAssetSummary?.TotalAssetsRmb ?? BuildTotalAssetsRmb(balances, exchangeRates),
                MissingExchangeRateCurrencies = usedCurrencies
                    .Where(currency => currency != CurrencyType.RMB && !exchangeRates.ContainsKey(currency))
                    .OrderBy(currency => currency)
                    .ToList(),
                LastMonthStart = lastMonthStart,
                LastMonthEnd = lastMonthEnd
            };
        }

        private Dictionary<CurrencyType, decimal> GetCurrencyToRmbRates()
        {
            var rates = new Dictionary<CurrencyType, decimal>
            {
                [CurrencyType.RMB] = 1
            };
            var finances = db.Queryable<Finance>()
                .Where(finance => finance.holdingType == HoldingType.Cash && finance._currentPrice_t == CurrencyType.RMB)
                .ToList();
            foreach (var finance in finances)
            {
                if (finance._currentPrice_v <= 0 || !Enum.TryParse<CurrencyType>(finance.code, out var currency))
                    continue;

                rates[currency] = finance._currentPrice_v;
            }

            return rates;
        }

        private static List<DateTime> BuildAssetSummaryDates(DateTime today)
        {
            var start = new DateTime(today.Year, today.Month, 1).AddMonths(-3);
            var dates = new List<DateTime>();
            for (var month = start; month <= today; month = month.AddMonths(1))
            {
                var daysInMonth = DateTime.DaysInMonth(month.Year, month.Month);
                for (var day = 1; day <= daysInMonth; day += 5)
                {
                    var date = new DateTime(month.Year, month.Month, day);
                    if (date >= start && date <= today)
                        dates.Add(date);
                }
            }

            if (!dates.Contains(today))
                dates.Add(today);

            return dates
                .Distinct()
                .OrderBy(date => date)
                .ToList();
        }

        private List<AssetSummaryBalanceSet> BuildAssetSummaryBalanceSets(
            List<DateTime> dates,
            List<AccountBalance> currentBalances,
            DateTime today)
        {
            var targetRevision = GetCurrentStatementImportRevision();
            var rollForwardStartDate = GetBalanceRollForwardStartDate();
            return dates
                .Select(date =>
                {
                    if (date == today)
                    {
                        return new AssetSummaryBalanceSet(
                            date,
                            null,
                            true,
                            currentBalances.Select(CloneDashboardBalance).ToList());
                    }

                    var baseSnapshot = GetLatestBalanceSnapshotAtOrBefore(date, targetRevision);
                    if (baseSnapshot is null)
                        return new AssetSummaryBalanceSet(date, null, false, []);

                    return new AssetSummaryBalanceSet(
                        date,
                        baseSnapshot.Snapshot.time,
                        true,
                        BuildRolledForwardBalances(baseSnapshot, date, targetRevision, rollForwardStartDate));
                })
                .ToList();
        }

        private DateTime GetBalanceRollForwardStartDate()
        {
            var snapshot = GetStartSnapshot();
            return snapshot?.effectiveDate.Date ?? DateTime.MinValue.Date;
        }

        private SnapshotData? GetLatestBalanceSnapshotAtOrBefore(DateTime date, int targetRevision)
        {
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.maxStatementImportId >= 0
                    && it.maxStatementImportId <= targetRevision
                    && it.effectiveDate <= date.Date)
                .OrderByDescending(it => it.effectiveDate)
                .OrderByDescending(it => it.maxStatementImportId)
                .OrderByDescending(it => it.time)
                .First();
            return snapshot is null ? null : ReadSnapshotData(snapshot);
        }

        private static AssetSummaryBalanceSet BuildAssetSummaryBalanceSetFromSnapshot(
            DateTime date,
            SnapshotData snapshot)
        {
            var balances = snapshot.AccountBalances
                .Where(balance => balance.Amount != 0)
                .Select(balance => new AccountBalance
                {
                    _account_Id = balance.AccountId,
                    t = balance.CurrencyType,
                    v = balance.Amount
                })
                .ToList();
            return new AssetSummaryBalanceSet(date, snapshot.Snapshot.time, true, balances);
        }

        private List<AccountBalance> BuildRolledForwardBalances(
            SnapshotData baseSnapshot,
            DateTime targetDate,
            int targetRevision,
            DateTime rollForwardStartDate)
        {
            var balances = baseSnapshot.AccountBalances
                .GroupBy(balance => (balance.AccountId, balance.CurrencyType))
                .ToDictionary(group => group.Key, group => group.Sum(balance => balance.Amount));
            var targetEnd = targetDate.Date.AddDays(1);
            var baseRevision = baseSnapshot.Snapshot.maxStatementImportId;
            // 快照表示某个StatementImport导入进度下的最新状态，而不是某个自然日的余额。
            // 因此只能向前滚动：从快照revision之后新增的导入里，取所有发生在目标日结束前的record。
            // 如果这些新导入的record发生在快照effectiveDate之前，也仍然要计入，因为快照创建时尚未包含它们。
            var records = db.Queryable<Record>()
                .Where(record => record._statementImport_Id > baseRevision
                    && record._statementImport_Id <= targetRevision
                    && record.date > rollForwardStartDate
                    && record.date < targetEnd)
                .ToList();
            foreach (var record in records)
            {
                var key = (record._account_Id, record.t);
                balances[key] = balances.TryGetValue(key, out var current)
                    ? current + record.v
                    : record.v;
            }

            return balances
                .Where(item => item.Value != 0)
                .Select(item => new AccountBalance
                {
                    _account_Id = item.Key.AccountId,
                    t = item.Key.CurrencyType,
                    v = item.Value
                })
                .ToList();
        }

        private static AccountBalance CloneDashboardBalance(AccountBalance balance)
        {
            return new AccountBalance
            {
                _account_Id = balance._account_Id,
                t = balance.t,
                v = balance.v
            };
        }

        private static List<AssetSummaryPoint> BuildAssetSummaryPoints(
            List<AssetSummaryBalanceSet> balanceSets,
            List<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates,
            DateTime today)
        {
            return balanceSets
                .Select(point =>
                {
                    if (!point.HasData)
                    {
                        return new AssetSummaryPoint
                        {
                            Date = point.Date,
                            SnapshotTime = point.SnapshotTime,
                            IsToday = point.Date == today,
                            HasData = false
                        };
                    }

                    var monthStart = new DateTime(point.Date.Year, point.Date.Month, 1);
                    var nextDay = point.Date.AddDays(1);
                    var monthToDateRecords = records
                        .Where(record => record.date >= monthStart && record.date < nextDay)
                        .ToList();
                    return new AssetSummaryPoint
                    {
                        Date = point.Date,
                        SnapshotTime = point.SnapshotTime,
                        IsToday = point.Date == today,
                        HasData = true,
                        TotalAssetsRmb = BuildTotalAssetsRmb(point.Balances, exchangeRates),
                        CurrencySummaries = BuildCurrencySummaries(point.Balances, monthToDateRecords, exchangeRates)
                    };
                })
                .ToList();
        }

        private static decimal BuildTotalAssetsRmb(
            List<AccountBalance> balances,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            return Currency.RoundMoney(balances
                .Select(balance => TryConvertToRmb(balance.v, balance.t, exchangeRates))
                .Where(value => value.HasValue)
                .Sum(value => value!.Value));
        }

        private static List<CurrencyBalanceSummary> BuildCurrencySummaries(
            List<AccountBalance> balances,
            List<Record> currentMonthRecords,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            return balances
                .GroupBy(balance => balance.t)
                .Select(group =>
                {
                    var currencyRecords = currentMonthRecords
                        .Where(record => record.t == group.Key)
                        .ToList();
                    var assets = Currency.RoundMoney(group.Where(balance => balance.v > 0).Sum(balance => balance.v));
                    var liabilities = Currency.RoundMoney(group.Where(balance => balance.v < 0).Sum(balance => balance.v));
                    return new CurrencyBalanceSummary
                    {
                        Currency = group.Key,
                        Assets = assets,
                        Liabilities = liabilities,
                        Net = Currency.RoundMoney(assets + liabilities),
                        TotalIncome = Currency.RoundMoney(currencyRecords.Where(record => record.v > 0).Sum(record => record.v)),
                        TotalExpense = Currency.RoundMoney(-currencyRecords.Where(record => record.v < 0).Sum(record => record.v)),
                        AccountCount = group.Select(balance => balance._account_Id).Distinct().Count()
                    };
                })
                .OrderByDescending(summary => Math.Abs(TryConvertToRmb(summary.Net, summary.Currency, exchangeRates) ?? 0))
                .ThenBy(summary => summary.Currency)
                .ToList();
        }

        private static List<MonthlyFlowSeries> BuildMonthlyFlowSeries(List<Record> records, DateTime firstMonth)
        {
            var months = Enumerable.Range(0, 12)
                .Select(firstMonth.AddMonths)
                .ToList();
            return records
                .Select(record => record.t)
                .Distinct()
                .OrderBy(currency => currency)
                .Select(currency =>
                {
                    var points = months
                        .Select(month =>
                        {
                            var monthRecords = records
                                .Where(record => record.t == currency
                                    && record.date >= month
                                    && record.date < month.AddMonths(1))
                                .ToList();
                            return new MonthlyFlowPoint
                            {
                                Month = month,
                                MonthLabel = month.ToString("MM月"),
                                Income = Currency.RoundMoney(monthRecords.Where(record => record.v > 0).Sum(record => record.v)),
                                Expense = Currency.RoundMoney(-monthRecords.Where(record => record.v < 0).Sum(record => record.v)),
                                NetChange = Currency.RoundMoney(monthRecords.Sum(record => record.v)),
                                IncomeSegments = BuildSingleCurrencySegments(currency, monthRecords.Where(record => record.v > 0).Sum(record => record.v)),
                                ExpenseSegments = BuildSingleCurrencySegments(currency, -monthRecords.Where(record => record.v < 0).Sum(record => record.v))
                            };
                        })
                        .ToList();
                    return new MonthlyFlowSeries
                    {
                        DisplayName = currency.ToString(),
                        Currency = currency,
                        Points = points,
                        TotalIncome = Currency.RoundMoney(points.Sum(point => point.Income)),
                        TotalExpense = Currency.RoundMoney(points.Sum(point => point.Expense)),
                        NetChange = Currency.RoundMoney(points.Sum(point => point.NetChange))
                    };
                })
                .ToList();
        }

        private static MonthlyFlowSeries BuildRmbMonthlyFlowSeries(
            List<Record> records,
            DateTime firstMonth,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var months = Enumerable.Range(0, 12)
                .Select(firstMonth.AddMonths)
                .ToList();
            var points = months
                .Select(month =>
                {
                    var convertedRecords = records
                        .Where(record => record.date >= month && record.date < month.AddMonths(1))
                        .Select(record => TryConvertToRmb(record.v, record.t, exchangeRates))
                        .Where(value => value.HasValue)
                        .Select(value => value!.Value)
                        .ToList();
                    return new MonthlyFlowPoint
                    {
                        Month = month,
                        MonthLabel = month.ToString("MM月"),
                        Income = Currency.RoundMoney(convertedRecords.Where(value => value > 0).Sum()),
                        Expense = Currency.RoundMoney(-convertedRecords.Where(value => value < 0).Sum()),
                        NetChange = Currency.RoundMoney(convertedRecords.Sum()),
                        IncomeSegments = BuildRmbMonthlySegments(
                            records.Where(record => record.date >= month && record.date < month.AddMonths(1) && record.v > 0),
                            exchangeRates),
                        ExpenseSegments = BuildRmbMonthlySegments(
                            records.Where(record => record.date >= month && record.date < month.AddMonths(1) && record.v < 0),
                            exchangeRates,
                            invertSign: true)
                    };
                })
                .ToList();
            return new MonthlyFlowSeries
            {
                DisplayName = "",
                Currency = CurrencyType.RMB,
                Points = points,
                TotalIncome = Currency.RoundMoney(points.Sum(point => point.Income)),
                TotalExpense = Currency.RoundMoney(points.Sum(point => point.Expense)),
                NetChange = Currency.RoundMoney(points.Sum(point => point.NetChange))
            };
        }

        private static List<MonthlyFlowAccountStatistics> BuildMonthlyFlowAccountStatistics(
            List<Account> accounts,
            List<Record> records,
            DateTime firstMonth,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var filterAccounts = accounts
                .Where(account => account.usage != AccountUsage.Unknown)
                .Select(account => new
                {
                    AccountId = account.Id,
                    AccountName = account.name,
                    AccountType = GetAccountType(account.name)
                })
                .OrderBy(account => account.AccountType)
                .ThenBy(account => account.AccountName)
                .ToList();
            var allAccountIds = filterAccounts.Select(account => account.AccountId).ToHashSet();
            var statistics = new List<MonthlyFlowAccountStatistics>
            {
                BuildMonthlyFlowAccountStatistic(
                    "所有账户",
                    records.Where(record => allAccountIds.Contains(record._account_Id)).ToList(),
                    firstMonth,
                    exchangeRates)
            };

            foreach (var group in filterAccounts
                .GroupBy(account => account.AccountType, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key))
            {
                var groupAccounts = group
                    .OrderBy(account => account.AccountName)
                    .ToList();
                if (groupAccounts.Count > 1)
                {
                    var accountIds = groupAccounts.Select(account => account.AccountId).ToHashSet();
                    statistics.Add(BuildMonthlyFlowAccountStatistic(
                        BuildAccountTypeDisplayName(group.Key),
                        records.Where(record => accountIds.Contains(record._account_Id)).ToList(),
                        firstMonth,
                        exchangeRates));
                }

                foreach (var account in groupAccounts)
                {
                    statistics.Add(BuildMonthlyFlowAccountStatistic(
                        account.AccountName,
                        records.Where(record => record._account_Id == account.AccountId).ToList(),
                        firstMonth,
                        exchangeRates));
                }
            }

            return statistics;
        }

        private static MonthlyFlowAccountStatistics BuildMonthlyFlowAccountStatistic(
            string displayName,
            List<Record> records,
            DateTime firstMonth,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            return new MonthlyFlowAccountStatistics
            {
                DisplayName = displayName,
                MonthlyFlowSeries = BuildMonthlyFlowSeries(records, firstMonth),
                RmbMonthlyFlowSeries = BuildRmbMonthlyFlowSeries(records, firstMonth, exchangeRates)
            };
        }

        private static ReasonFlowSeries BuildRmbReasonFlowSeries(
            List<Record> records,
            DateTime start,
            DateTime end,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var lastMonthRecords = records
                .Where(record => record.date >= start && record.date < end)
                .ToList();
            var items = lastMonthRecords
                .GroupBy(record => String.IsNullOrWhiteSpace(record.Reason) ? "未分类" : record.Reason)
                .SelectMany(group =>
                {
                    var result = new List<ReasonFlowItem>();
                    var expenses = BuildReasonFlowItem(group.Key, false, group.Where(record => record.v < 0), exchangeRates);
                    var incomes = BuildReasonFlowItem(group.Key, true, group.Where(record => record.v > 0), exchangeRates);
                    if (expenses is not null)
                        result.Add(expenses);
                    if (incomes is not null)
                        result.Add(incomes);
                    return result;
                })
                .Where(item => item.Total != 0)
                .OrderBy(item => item.IsIncome)
                .ThenByDescending(item => item.Total)
                .ThenBy(item => item.Reason)
                .ToList();
            return new ReasonFlowSeries
            {
                DisplayName = "",
                Currency = CurrencyType.RMB,
                Month = start,
                MonthLabel = start.ToString("yyyy年MM月"),
                Items = items,
                TotalIncome = Currency.RoundMoney(items.Where(item => item.IsIncome).Sum(item => item.Total)),
                TotalExpense = Currency.RoundMoney(items.Where(item => !item.IsIncome).Sum(item => item.Total))
            };
        }

        private static List<MonthlyFlowSegment> BuildSingleCurrencySegments(CurrencyType currency, decimal value)
        {
            var rounded = Currency.RoundMoney(value);
            return rounded == 0
                ? []
                : [new MonthlyFlowSegment { Currency = currency, Value = rounded, Label = currency.ToString() }];
        }

        private static List<MonthlyFlowSegment> BuildRmbMonthlySegments(
            IEnumerable<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates,
            bool invertSign = false)
        {
            return records
                .GroupBy(record => record.t)
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    var original = group.Sum(record => record.v);
                    if (invertSign)
                        original = -original;
                    var converted = TryConvertToRmb(original, group.Key, exchangeRates) ?? 0;
                    return new MonthlyFlowSegment
                    {
                        Currency = group.Key,
                        Value = converted,
                        Label = group.Key.ToString()
                    };
                })
                .Where(segment => segment.Value != 0)
                .ToList();
        }

        private static ReasonFlowItem? BuildReasonFlowItem(
            string reason,
            bool isIncome,
            IEnumerable<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var groups = records
                .GroupBy(record => record.t)
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    var original = group.Sum(record => record.v);
                    if (!isIncome)
                        original = -original;
                    var converted = TryConvertToRmb(original, group.Key, exchangeRates);
                    return new
                    {
                        Currency = group.Key,
                        Original = Currency.RoundMoney(original),
                        Converted = converted
                    };
                })
                .Where(item => item.Converted.HasValue && item.Converted.Value != 0)
                .ToList();
            if (groups.Count == 0)
                return null;

            var total = Currency.RoundMoney(groups.Sum(item => item.Converted!.Value));
            var detailSum = Currency.RoundMoney(groups.Sum(item => item.Converted!.Value));
            if (total != detailSum)
                throw new InvalidOperationException($"Reason currency conversion mismatch: {reason}/{(isIncome ? "income" : "expense")}");

            return new ReasonFlowItem
            {
                Reason = reason,
                IsIncome = isIncome,
                Total = total,
                CurrencyDetails = BuildCurrencyDetailText(groups.Select(item => (item.Currency, item.Original, item.Converted!.Value)))
            };
        }

        private static decimal? TryConvertToRmb(decimal value, CurrencyType currency, Dictionary<CurrencyType, decimal> exchangeRates)
        {
            return exchangeRates.TryGetValue(currency, out var rate)
                ? Currency.RoundMoney(value * rate)
                : null;
        }

        private static string BuildCurrencyDetailText(IEnumerable<(CurrencyType Currency, decimal Original, decimal Converted)> details)
        {
            var currencies = details
                .Where(item => item.Currency != CurrencyType.RMB)
                .Select(item => $"{item.Currency} ¥{item.Converted:N0}")
                .ToList();
            return currencies.Count == 0 ? "" : String.Join("、", currencies);
        }

        private static Dictionary<(int AccountId, string Code), string> BuildHoldingNames(List<Holding> holdings)
        {
            return holdings
                .GroupBy(holding => (holding._account_Id, holding.code))
                .ToDictionary(
                    group => group.Key,
                    group =>
                    {
                        var holding = group.First();
                        var display = String.IsNullOrWhiteSpace(holding.displayText)
                            ? holding.code
                            : holding.displayText;
                        if (String.IsNullOrWhiteSpace(display))
                            display = holding.holdingType.ToString();
                        return display;
                    });
        }

        private static string BuildHoldingInvestmentKey(Record record, Dictionary<(int AccountId, string Code), string> holdingNames)
        {
            if (String.IsNullOrWhiteSpace(record.DestAccount))
                return "未关联持仓";

            return holdingNames.TryGetValue((record._account_Id, record.DestAccount), out var display)
                ? display
                : record.DestAccount;
        }

        private static List<InvestmentAccountStatistics> BuildInvestmentAccountStatistics(
            Dictionary<int, string> accounts,
            IEnumerable<int> investmentAccountIds,
            List<Record> investmentRecords,
            DateTime today,
            Dictionary<CurrencyType, decimal> exchangeRates,
            Dictionary<(int AccountId, string Code), string> holdingNames)
        {
            var investmentAccounts = investmentAccountIds
                .Select(accountId =>
                {
                    var accountName = accounts.TryGetValue(accountId, out var name) ? name : $"Account {accountId}";
                    return new
                    {
                        AccountId = accountId,
                        AccountName = accountName,
                        AccountType = GetAccountType(accountName)
                    };
                })
                .OrderBy(account => account.AccountType)
                .ThenBy(account => account.AccountName)
                .ToList();

            var statistics = new List<InvestmentAccountStatistics>
            {
                BuildInvestmentAccountStatistic("所有账户", investmentRecords, today, exchangeRates, holdingNames)
            };

            foreach (var group in investmentAccounts
                .GroupBy(account => account.AccountType, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key))
            {
                var groupAccounts = group
                    .OrderBy(account => account.AccountName)
                    .ToList();
                if (groupAccounts.Count > 1)
                {
                    var accountIds = groupAccounts.Select(account => account.AccountId).ToHashSet();
                    var accountRecords = investmentRecords
                        .Where(record => accountIds.Contains(record._account_Id))
                        .ToList();
                    statistics.Add(BuildInvestmentAccountStatistic(
                        BuildAccountTypeDisplayName(group.Key),
                        accountRecords,
                        today,
                        exchangeRates,
                        holdingNames));
                }

                foreach (var account in groupAccounts)
                {
                    var accountRecords = investmentRecords
                        .Where(record => record._account_Id == account.AccountId)
                        .ToList();
                    var statistic = BuildInvestmentAccountStatistic(
                        account.AccountName,
                        accountRecords,
                        today,
                        exchangeRates,
                        holdingNames);
                    statistics.Add(statistic);
                }
            }

            return statistics;
        }

        private static InvestmentAccountStatistics BuildInvestmentAccountStatistic(
            string displayName,
            List<Record> records,
            DateTime today,
            Dictionary<CurrencyType, decimal> exchangeRates,
            Dictionary<(int AccountId, string Code), string> holdingNames)
        {
            return new InvestmentAccountStatistics
            {
                DisplayName = displayName,
                ByReason = BuildInvestmentStatistics(
                    records,
                    today,
                    exchangeRates,
                    record => String.IsNullOrWhiteSpace(record.Reason) ? "未分类" : record.Reason),
                ByHolding = BuildInvestmentStatistics(
                    records,
                    today,
                    exchangeRates,
                    record => BuildHoldingInvestmentKey(record, holdingNames))
            };
        }

        private static string GetAccountType(string accountName)
        {
            var normalized = accountName.Trim();
            var separatorIndex = normalized.IndexOf('_');
            return separatorIndex > 0 ? normalized[..separatorIndex] : normalized;
        }

        private static string BuildAccountTypeDisplayName(string accountType)
        {
            return String.IsNullOrWhiteSpace(accountType)
                ? "未分类账户"
                : $"{accountType} 账户";
        }

        private static string NormalizeInternalCardNoForStorage(string value)
        {
            return Regex.Replace(value.Trim(), @"\s+", "");
        }

        private static HashSet<string> ExtractAccountTypesFromTexts(List<string> texts, List<string> accountTypes)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var accountType in accountTypes)
            {
                var escapedType = Regex.Escape(accountType);
                if (texts.Any(text => Regex.IsMatch(
                        text,
                        $@"(^|[^A-Za-z0-9]){escapedType}\s*[-_:]",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)))
                    result.Add(accountType);
            }

            return result;
        }

        private static InternalCardNoMatchDetail? FindInternalCardNoMatch(string cardNo, List<string> texts)
        {
            var cardToken = NormalizeInternalCardToken(cardNo);
            var cardDigits = NormalizeInternalCardDigits(cardNo);
            if (cardToken.Length < 3 && cardDigits.Length < 3)
                return null;

            foreach (var text in texts)
            {
                var textToken = NormalizeInternalCardToken(text);
                if (cardToken.Length >= 3 && textToken.Contains(cardToken, StringComparison.Ordinal))
                    return new InternalCardNoMatchDetail(text, "normalized token contains full card number");

                var textDigits = NormalizeInternalCardDigits(text);
                if (cardDigits.Length >= 3 && textDigits.Contains(cardDigits, StringComparison.Ordinal))
                    return new InternalCardNoMatchDetail(text, "digits contain full card number");

                foreach (Match match in Regex.Matches(text, @"\d{3,}", RegexOptions.CultureInvariant))
                {
                    var digits = match.Value;
                    if (cardDigits.EndsWith(digits, StringComparison.Ordinal))
                        return new InternalCardNoMatchDetail(text, $"text digits {digits} match card suffix");
                }
            }

            return null;
        }

        private static List<InternalCardNoCandidate> PreferKnownAccountMatches(List<InternalCardNoCandidate> matches)
        {
            var knownMatches = matches
                .Where(match => !IsUnknownAccount(match.Account))
                .ToList();
            return knownMatches.Count > 0 ? knownMatches : matches;
        }

        private static bool IsUnknownAccount(Account account)
        {
            return account.usage == AccountUsage.Unknown;
        }

        private static string NormalizeInternalCardToken(string value)
        {
            return Regex.Replace(value, @"[^A-Za-z0-9]", "").ToUpperInvariant();
        }

        private static string NormalizeInternalCardDigits(string value)
        {
            return Regex.Replace(value, @"\D", "");
        }

        private static void LogAccountInternalCardNoKnown(string action, AccountInternalId internalId, Account account, string cardNo)
        {
            if (String.IsNullOrWhiteSpace(internalId.sourceText))
                return;

            Console.WriteLine(
                $"Internal card {action}: source=\"{NormalizeLogText(internalId.sourceText)}\"; account={account.name}; cardNo={cardNo}; currency={internalId.currencyType?.ToString() ?? ""}");
        }

        private static void LogInternalCardNoMatch(string? matchContext, InternalCardNoCandidate candidate, List<string> allTexts)
        {
            var source = String.IsNullOrWhiteSpace(matchContext)
                ? "unspecified"
                : NormalizeLogText(matchContext);
            Console.WriteLine(
                $"Internal card match: source=\"{source}\"; account={candidate.Account.name}; cardNo={candidate.InternalId.cardNo}; matchedText=\"{NormalizeLogText(candidate.Match.Text)}\"; match={candidate.Match.MatchReason}; allText=\"{NormalizeLogText(String.Join(" | ", allTexts))}\"");
        }

        private static string NormalizeLogText(string text)
        {
            return Regex.Replace(text, @"\s+", " ").Trim().Replace("\"", "'");
        }

        private sealed record InternalCardNoCandidate(
            AccountInternalId InternalId,
            Account Account,
            string AccountType,
            InternalCardNoMatchDetail Match);

        private sealed record InternalCardNoMatchDetail(string Text, string MatchReason);

        private static InvestmentStatistics BuildInvestmentStatistics(
            List<Record> records,
            DateTime today,
            Dictionary<CurrencyType, decimal> exchangeRates,
            Func<Record, string> keySelector)
        {
            var end = today.Date.AddDays(1);
            return new InvestmentStatistics
            {
                Periods =
                [
                    BuildInvestmentStatisticsPeriod("过去一个月", records, today.Date.AddMonths(-1), end, exchangeRates, keySelector),
                    BuildInvestmentStatisticsPeriod("年初至今", records, new DateTime(today.Year, 1, 1), end, exchangeRates, keySelector)
                ]
            };
        }

        private static InvestmentStatisticsPeriod BuildInvestmentStatisticsPeriod(
            string title,
            List<Record> records,
            DateTime start,
            DateTime end,
            Dictionary<CurrencyType, decimal> exchangeRates,
            Func<Record, string> keySelector)
        {
            var items = records
                .Where(record => record.date >= start && record.date < end && record.v > 0)
                .GroupBy(keySelector)
                .Select(group => new InvestmentStatisticsItem
                {
                    Name = group.Key,
                    Total = Currency.RoundMoney(group
                        .Select(record => TryConvertToRmb(record.v, record.t, exchangeRates))
                        .Where(value => value.HasValue)
                        .Sum(value => value!.Value))
                })
                .Where(item => item.Total > 0)
                .OrderByDescending(item => item.Total)
                .ThenBy(item => item.Name)
                .ToList();
            return new InvestmentStatisticsPeriod
            {
                Title = title,
                Items = items,
                Total = Currency.RoundMoney(items.Sum(item => item.Total))
            };
        }

        public DatabaseCleanupResult CleanVolatileData(int? cleanToSnapshotId = null)
        {
            var beforeCounts = ReadCleanupCounts();
            ExecuteLockedTransaction(() =>
            {
                if (cleanToSnapshotId.HasValue)
                    CleanToSnapshotCore(cleanToSnapshotId.Value);
                else
                    CleanToStartSnapshotCore();
            });

            return new DatabaseCleanupResult(
                beforeCounts,
                ReadCleanupCounts(),
                db.Queryable<StatementImport>()
                    .OrderBy(it => it.provider)
                    .OrderBy(it => it.time)
                    .ToList());
        }

        private void CleanToStartSnapshotCore()
        {
            var startSnapshot = GetStartSnapshot();
            CleanSnapshotsAfterStartSnapshot(startSnapshot);
            ClearAllRecordMatches();
            db.Deleteable<Record>().ExecuteCommand();
            db.Ado.ExecuteCommand("""
                delete statementImport
                from `StatementImports` statementImport
                left join (
                    select candidate.`provider`, min(candidate.`Id`) as `Id`
                    from `StatementImports` candidate
                    join (
                        select `provider`, min(`time`) as `time`
                        from `StatementImports`
                        group by `provider`
                    ) earliest
                        on candidate.`provider` = earliest.`provider`
                        and candidate.`time` = earliest.`time`
                    group by candidate.`provider`
                ) fixedImport
                    on statementImport.`Id` = fixedImport.`Id`
                where fixedImport.`Id` is null
                """);
            if (startSnapshot is null)
            {
                db.Deleteable<Holding>().ExecuteCommand();
                db.Deleteable<AccountBalance>().ExecuteCommand();
                return;
            }

            RestoreCurrentStateFromSnapshot(ReadSnapshotData(startSnapshot));
        }

        private Snapshot? GetStartSnapshot()
        {
            return db.Queryable<Snapshot>()
                .OrderBy(snapshot => snapshot.Id)
                .First();
        }

        private void CleanSnapshotsAfterStartSnapshot(Snapshot? startSnapshot)
        {
            if (startSnapshot is null)
            {
                db.Deleteable<SnapshotItem>().ExecuteCommand();
                db.Deleteable<Snapshot>().ExecuteCommand();
                return;
            }

            db.Ado.ExecuteCommand("""
                delete from `SnapshotItems`
                where `_snapshot_Id` <> @snapshotId
                """,
                new SugarParameter("@snapshotId", startSnapshot.Id));
            db.Ado.ExecuteCommand("""
                delete from `Snapshots`
                where `Id` <> @snapshotId
                """,
                new SugarParameter("@snapshotId", startSnapshot.Id));
        }

        private void CleanToSnapshotCore(int snapshotId)
        {
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.Id == snapshotId)
                .First();
            if (snapshot is null)
                throw new InvalidOperationException($"Snapshot not found: {snapshotId}");
            if (snapshot.maxStatementImportId < 0)
                throw new InvalidOperationException($"Snapshot does not have a valid import revision: {snapshotId}");

            var snapshotData = ReadSnapshotData(snapshot);
            DeleteSnapshotsAfter(snapshot);
            ClearAllRecordMatches();
            db.Ado.ExecuteCommand("""
                delete from `Records`
                where `_statementImport_Id` > @maxStatementImportId
                """,
                new SugarParameter("@maxStatementImportId", snapshot.maxStatementImportId));
            db.Ado.ExecuteCommand("""
                delete from `StatementImports`
                where `Id` > @maxStatementImportId
                """,
                new SugarParameter("@maxStatementImportId", snapshot.maxStatementImportId));
            RestoreCurrentStateFromSnapshot(snapshotData);
        }

        private void DeleteSnapshotsAfter(Snapshot snapshot)
        {
            db.Ado.ExecuteCommand("""
                delete snapshotItem
                from `SnapshotItems` snapshotItem
                join `Snapshots` snapshot
                    on snapshotItem.`_snapshot_Id` = snapshot.`Id`
                where snapshot.`maxStatementImportId` > @maxStatementImportId
                    or (
                        snapshot.`maxStatementImportId` = @maxStatementImportId
                        and snapshot.`Id` > @snapshotId
                    )
                """,
                new SugarParameter("@maxStatementImportId", snapshot.maxStatementImportId),
                new SugarParameter("@snapshotId", snapshot.Id));
            db.Ado.ExecuteCommand("""
                delete from `Snapshots`
                where `maxStatementImportId` > @maxStatementImportId
                    or (
                        `maxStatementImportId` = @maxStatementImportId
                        and `Id` > @snapshotId
                    )
                """,
                new SugarParameter("@maxStatementImportId", snapshot.maxStatementImportId),
                new SugarParameter("@snapshotId", snapshot.Id));
        }

        private void RestoreCurrentStateFromSnapshot(SnapshotData snapshot)
        {
            db.Deleteable<Holding>().ExecuteCommand();
            db.Deleteable<AccountBalance>().ExecuteCommand();

            var accountBalances = snapshot.AccountBalances
                .Select(balance => new AccountBalance
                {
                    _account_Id = balance.AccountId,
                    t = balance.CurrencyType,
                    v = balance.Amount
                })
                .ToList();
            if (accountBalances.Count > 0)
                db.Insertable(accountBalances).ExecuteCommand();

            var holdings = snapshot.Holdings
                .Select(holding => new Holding
                {
                    _account_Id = holding.AccountId,
                    code = holding.Code,
                    holdingType = holding.HoldingType,
                    quantity = holding.Quantity,
                    _currentPrice_v = holding.CurrentPrice.v,
                    _currentPrice_t = holding.CurrentPrice.t,
                    displayText = holding.DisplayText,
                    desc = holding.Description
                })
                .ToList();
            if (holdings.Count > 0)
                db.Insertable(holdings).ExecuteCommand();
        }

        public void CleanWiseImportedData()
        {
            ExecuteLockedTransaction(() =>
            {
                var wiseAccount = GetAccountByName("WISE");
                ClearRecordMatchesForStatementProvider(StatementImportProvider.WiseMail);
                db.Ado.ExecuteCommand("""
                    delete record
                    from `Records` record
                    join `StatementImports` statementImport
                        on record.`_statementImport_Id` = statementImport.`Id`
                    where statementImport.`provider` = @provider
                    """,
                    new SugarParameter("@provider", StatementImportProvider.WiseMail.ToString()));
                db.Deleteable<AccountBalance>()
                    .Where(balance => balance._account_Id == wiseAccount.Id)
                    .ExecuteCommand();
                db.Deleteable<AccountInternalId>()
                    .Where(internalId => internalId._account_Id == wiseAccount.Id
                        && (internalId.desc == "XML statement balance id" || internalId.desc == "XML file balance id"))
                    .ExecuteCommand();
                db.Ado.ExecuteCommand("""
                    delete statementImport
                    from `StatementImports` statementImport
                    left join (
                        select `Id`
                        from `StatementImports`
                        where `provider` = @provider
                        order by `time`, `Id`
                        limit 1
                    ) fixedImport
                        on statementImport.`Id` = fixedImport.`Id`
                    where statementImport.`provider` = @provider
                        and fixedImport.`Id` is null
                    """,
                    new SugarParameter("@provider", StatementImportProvider.WiseMail.ToString()));

                if (wiseAccount.relativeBalance)
                {
                    wiseAccount.relativeBalance = false;
                    db.Updateable(wiseAccount)
                        .UpdateColumns(account => new { account.relativeBalance })
                        .ExecuteCommand();
                }
            });
        }

        private void ClearAllRecordMatches()
        {
            db.Ado.ExecuteCommand("""
                update `Records`
                set `matchedRecordId` = null,
                    `matchedRecordReason` = ''
                where `matchedRecordId` is not null
                    or `matchedRecordReason` <> ''
                """);
        }

        private void ClearRecordMatchesForStatementProvider(StatementImportProvider provider)
        {
            db.Ado.ExecuteCommand("""
                update `Records` record
                left join `Records` matchedRecord
                    on record.`matchedRecordId` = matchedRecord.`Id`
                left join `StatementImports` recordImport
                    on record.`_statementImport_Id` = recordImport.`Id`
                left join `StatementImports` matchedImport
                    on matchedRecord.`_statementImport_Id` = matchedImport.`Id`
                set record.`matchedRecordId` = null,
                    record.`matchedRecordReason` = ''
                where recordImport.`provider` = @provider
                    or matchedImport.`provider` = @provider
                """,
                new SugarParameter("@provider", provider.ToString()));
        }

        private Dictionary<string, int> ReadCleanupCounts()
        {
            return new Dictionary<string, int>
            {
                ["Accounts"] = db.Queryable<Account>().Count(),
                ["AccountBalances"] = db.Queryable<AccountBalance>().Count(),
                ["StatementImports"] = db.Queryable<StatementImport>().Count(),
                ["Records"] = db.Queryable<Record>().Count(),
                ["Holdings"] = db.Queryable<Holding>().Count(),
                ["Finance"] = db.Queryable<Finance>().Count(),
                ["Snapshots"] = db.Queryable<Snapshot>().Count(),
                ["SnapshotItems"] = db.Queryable<SnapshotItem>().Count()
            };
        }

        public void SaveFinance(Finance finance)
        {
            ExecuteLockedTransaction(() =>
            {
                if (finance.currentPriceTime <= 0)
                    finance.currentPriceTime = DateTimeOffset.Now.ToUnixTimeSeconds();

                var existing = db.Queryable<Finance>()
                    .Where(it => it.code == finance.code && it.holdingType == finance.holdingType)
                    .First();
                if (existing is null)
                {
                    finance.Id = db.Insertable(finance).ExecuteReturnIdentity();
                    return;
                }

                existing._currentPrice_v = finance._currentPrice_v;
                existing._currentPrice_t = finance._currentPrice_t;
                existing.currentPriceTime = finance.currentPriceTime;
                db.Updateable(existing).ExecuteCommand();
            });
        }

        public List<Record> GetRecordsByStatementImport(int statementImportId)
        {
            return db.Queryable<Record>()
                .Where(record => record._statementImport_Id == statementImportId)
                .ToList();
        }

        public List<Record> GetStatementRecords(StatementImportProvider provider, DateTime start, DateTime end)
        {
            var importIds = db.Queryable<StatementImport>()
                .Where(statementImport => statementImport.provider == provider)
                .Select(statementImport => statementImport.Id)
                .ToList();
            if (importIds.Count == 0)
                return [];

            return db.Queryable<Record>()
                .Where(record => importIds.Contains(record._statementImport_Id)
                    && record.date >= start
                    && record.date < end)
                .ToList();
        }

        public List<Record> GetStatementRecords(StatementImportProvider provider, Account account, DateTime start, DateTime end)
        {
            var postingAccount = GetPostingAccount(account);
            return GetStatementRecords(provider, start, end)
                .Where(record => record._account_Id == postingAccount.Id)
                .ToList();
        }

        public List<StatementImport> GetStatementImports(StatementImportProvider provider)
        {
            return db.Queryable<StatementImport>()
                .Where(statementImport => statementImport.provider == provider)
                .ToList();
        }

        public List<string> GetStatementImportKeys(StatementImportProvider provider)
        {
            return db.Queryable<StatementImport>()
                .Where(statementImport => statementImport.provider == provider && statementImport.statementKey != "")
                .Select(statementImport => statementImport.statementKey)
                .ToList();
        }

        public HashSet<string> GetRecordSourceCodes(string prefix)
        {
            return db.Queryable<Record>()
                .Where(record => record.Source.Contains($"code={prefix}"))
                .Select(record => record.Source)
                .ToList()
                .Select(ExtractRecordSourceCode)
                .Where(code => code.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        public bool SaveRecordSourceSupplementsOnce(
            StatementImportProvider provider,
            DateTime time,
            string statementKey,
            IEnumerable<RecordSourceSupplement> supplements,
            IEnumerable<AccountInternalId>? internalCardNos = null)
        {
            var supplementList = supplements.ToList();
            var internalCardNoList = internalCardNos?.ToList() ?? [];
            if (supplementList.Count == 0)
                return false;

            try
            {
                return ExecuteLockedTransaction(() =>
                {
                    if (IsStatementImported(provider, time, statementKey))
                        return false;

                    InsertStatementImport(provider, time, statementKey);
                    var now = DateTime.Now;
                    foreach (var supplement in supplementList)
                    {
                        var record = db.Queryable<Record>()
                            .Where(it => it.Id == supplement.RecordId)
                            .First();
                        if (record is null)
                            throw new InvalidOperationException($"Supplement target record not found: {supplement.RecordId}");

                        if (record.Source.Contains($"code={supplement.SourceCode}", StringComparison.OrdinalIgnoreCase))
                            continue;

                        record.Source = AppendLimitedRecordText(record.Source, supplement.SourceAppend);
                        record.updateTime = now;
                        db.Updateable(record)
                            .UpdateColumns(it => new { it.Source, it.updateTime })
                            .ExecuteCommand();
                    }

                    if (internalCardNoList.Count > 0)
                        EnsureAccountInternalCardNos(internalCardNoList);

                    return true;
                });
            }
            catch (Exception e)
            {
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time, statementKey))
                    return false;
                throw;
            }
        }

        public void MarkRecordsAsRefundMatched(IEnumerable<Record> records)
        {
            var updates = records.ToList();
            if (updates.Count == 0)
                return;

            var now = DateTime.Now;
            foreach (var record in updates)
            {
                record.isRefundMatched = true;
                record.updateTime = now;
            }

            db.Updateable(updates)
                .UpdateColumns(record => new { record.isRefundMatched, record.updateTime })
                .ExecuteCommand();
        }

        public static DateTime NormalizeStatementImportTime(DateTime time)
        {
            return time.Date;
        }

        private static bool IsDuplicateKeyException(Exception exception)
        {
            for (var current = exception; current is not null; current = current.InnerException)
            {
                if (current.Message.Contains("Duplicate entry", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private void ValidateSchema()
        {
            foreach (var type in SchemaTypes)
            {
                var tableName = GetTableName(type);
                if (!db.DbMaintenance.IsAnyTable(tableName, false))
                    throw new InvalidOperationException($"Database schema mismatch: missing table {tableName}");

                var dbColumns = db.DbMaintenance.GetColumnInfosByTableName(tableName, false)
                    .Select(column => column.DbColumnName)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                var codeColumns = GetColumnNames(type);

                var missingColumns = codeColumns.Except(dbColumns, StringComparer.OrdinalIgnoreCase).ToList();
                var extraColumns = dbColumns.Except(codeColumns, StringComparer.OrdinalIgnoreCase).ToList();
                if (missingColumns.Count > 0 || extraColumns.Count > 0)
                {
                    var missingText = missingColumns.Count > 0 ? String.Join(",", missingColumns) : "none";
                    var extraText = extraColumns.Count > 0 ? String.Join(",", extraColumns) : "none";
                    throw new InvalidOperationException(
                        $"Database schema mismatch: table {tableName}, missing columns [{missingText}], extra columns [{extraText}]");
                }
            }
        }

        private void ValidateForeignKeys()
        {
            foreach (var foreignKey in ForeignKeys)
            {
                var exists = db.Ado.GetInt("""
                    select count(*)
                    from information_schema.key_column_usage
                    where table_schema = database()
                        and constraint_name = @constraintName
                        and table_name = @tableName
                        and column_name = @columnName
                        and referenced_table_name = @referencedTableName
                        and referenced_column_name = @referencedColumnName
                    """,
                    new SugarParameter("@constraintName", foreignKey.ConstraintName),
                    new SugarParameter("@tableName", foreignKey.TableName),
                    new SugarParameter("@columnName", foreignKey.ColumnName),
                    new SugarParameter("@referencedTableName", foreignKey.ReferencedTableName),
                    new SugarParameter("@referencedColumnName", foreignKey.ReferencedColumnName)) > 0;
                if (!exists)
                {
                    throw new InvalidOperationException(
                        $"Database schema mismatch: missing foreign key {foreignKey.ConstraintName} on {foreignKey.TableName}.{foreignKey.ColumnName}");
                }
            }
        }

        private void ValidateAccountPrimaryRelations()
        {
            var accounts = db.Queryable<Account>().ToList();
            var accountsById = accounts.ToDictionary(account => account.Id);
            foreach (var account in accounts.Where(account => account._primaryAccount_Id.HasValue))
            {
                var primaryAccountId = account._primaryAccount_Id!.Value;
                if (primaryAccountId == account.Id)
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name} points to itself");

                if (!accountsById.TryGetValue(primaryAccountId, out var primaryAccount))
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name} points to missing account {primaryAccountId}");

                if (primaryAccount._primaryAccount_Id.HasValue)
                    throw new InvalidOperationException($"Invalid account primary relation: primary account {primaryAccount.name} is also a supplementary account");
            }
        }

        private static string GetTableName(Type type)
        {
            if (type == typeof(Account))
                return "Accounts";
            if (type == typeof(AccountInternalId))
                return "AccountInternalIds";
            if (type == typeof(AccountBalance))
                return "AccountBalances";
            if (type == typeof(Record))
                return "Records";
            if (type == typeof(Holding))
                return "Holdings";
            if (type == typeof(Finance))
                return "Finance";
            if (type == typeof(Snapshot))
                return "Snapshots";
            if (type == typeof(SnapshotItem))
                return "SnapshotItems";
            if (type == typeof(StatementImport))
                return "StatementImports";

            return type.Name;
        }

        private static HashSet<string> GetColumnNames(Type type)
        {
            return type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.GetIndexParameters().Length == 0)
                .Where(property => property.GetCustomAttribute<SugarColumn>()?.IsIgnore != true)
                .Where(property => !HasNavigateAttribute(property))
                .Select(property =>
                {
                    var column = property.GetCustomAttribute<SugarColumn>();
                    return String.IsNullOrWhiteSpace(column?.ColumnName) ? property.Name : column.ColumnName;
                })
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private static bool HasNavigateAttribute(PropertyInfo property)
        {
            return property.GetCustomAttributes()
                .Any(attribute => attribute.GetType().Name is "Navigate" or "NavigateAttribute");
        }

        private static string GetHoldingKey(Holding holding)
        {
            return $"{holding.code}\t{holding.holdingType}";
        }

        private void SaveAccountHoldingsCore(Account account, List<Holding> holdingList)
        {
            account = GetAccountByName(account.name);

            foreach (var holding in holdingList)
            {
                NormalizeHolding(holding);
                holding.Account = account;
                holding._account_Id = account.Id;
            }

            ValidateAccountHoldingsBalance(account, holdingList);

            var existingHoldings = db.Queryable<Holding>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            var currentKeys = holdingList.Select(GetHoldingKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var deletedIds = existingHoldings
                .Where(holding => !currentKeys.Contains(GetHoldingKey(holding)))
                .Select(holding => holding.Id)
                .ToList();
            if (deletedIds.Count > 0)
                db.Deleteable<Holding>().In(deletedIds).ExecuteCommand();

            foreach (var holding in holdingList)
            {
                var existing = existingHoldings.FirstOrDefault(it =>
                    it.code == holding.code && it.holdingType == holding.holdingType);
                if (existing is null)
                {
                    db.Insertable(holding).ExecuteCommand();
                    continue;
                }

                existing.quantity = holding.quantity;
                existing.desc = holding.desc;
                existing.displayText = holding.displayText;
                existing._currentPrice_v = holding._currentPrice_v;
                existing._currentPrice_t = holding._currentPrice_t;
                existing._account_Id = account.Id;
                db.Updateable(existing).ExecuteCommand();
            }
        }

        private static void NormalizeHolding(Holding holding)
        {
            if (Holding.IsSingleValueAsset(holding.holdingType))
                holding.quantity = 1;
        }

        private void ValidateAccountHoldingsBalance(Account account, List<Holding> holdings)
        {
            var accountBalances = db.Queryable<AccountBalance>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            if (accountBalances.Count == 0)
                throw new InvalidOperationException($"Missing account balance for holdings: {account.name}");

            var balanceSums = accountBalances
                .GroupBy(balance => balance.t)
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(balance => balance.v)));
            var holdingSums = holdings
                .GroupBy(holding => holding.currentPrice.t)
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(holding => holding.totalPrice.v)));
            var currencies = balanceSums.Keys
                .Union(holdingSums.Keys)
                .ToList();

            foreach (var currency in currencies)
            {
                var accountBalance = balanceSums.TryGetValue(currency, out var balance) ? balance : 0;
                var holdingTotal = holdingSums.TryGetValue(currency, out var total) ? total : 0;
                if (accountBalance != holdingTotal)
                {
                    throw new InvalidOperationException(
                        $"Holding balance mismatch for {account.name}/{currency}: accountBalance={accountBalance}, holdings={holdingTotal}");
                }
            }
        }

        private sealed record SnapshotAccountBalancePayloadV1(
            int AccountId,
            string AccountName,
            string CurrencyType,
            decimal Amount);

        private sealed record SnapshotHoldingPayloadV1(
            int AccountId,
            string AccountName,
            string Code,
            string HoldingType,
            int Quantity,
            decimal PriceAmount,
            string PriceCurrencyType,
            decimal TotalAmount,
            string TotalCurrencyType,
            string DisplayText,
            string Description);

        private sealed record RecordBackupPayloadV1(
            int SchemaVersion,
            int Id,
            int AccountId,
            decimal Amount,
            CurrencyType Currency,
            DateTime Date,
            DateTime? PostingDate,
            DateTime UpdateTime,
            string DestAccount,
            bool IsInternal,
            int? MatchedRecordId,
            string MatchedRecordReason,
            bool IsRefundMatched,
            int HoldingQuantity,
            string Source,
            string Reason,
            decimal? DescCurrencyAmount,
            CurrencyType? DescCurrency,
            int StatementImportId,
            StatementImportProvider StatementProvider,
            string StatementKey);

        private sealed record AssetSummaryBalanceSet(
            DateTime Date,
            DateTime? SnapshotTime,
            bool HasData,
            List<AccountBalance> Balances);

        private sealed record ForeignKeyDefinition(
            string ConstraintName,
            string TableName,
            string ColumnName,
            string ReferencedTableName,
            string ReferencedColumnName);
    }

    class SnapshotData
    {
        public SnapshotData(
            Snapshot snapshot,
            List<SnapshotAccountBalanceData> accountBalances,
            List<SnapshotHoldingData> holdings)
        {
            Snapshot = snapshot;
            AccountBalances = accountBalances;
            Holdings = holdings;
        }

        public Snapshot Snapshot { get; }
        public List<SnapshotAccountBalanceData> AccountBalances { get; }
        public List<SnapshotHoldingData> Holdings { get; }
    }

    record SnapshotAccountBalanceData(
        int AccountId,
        string AccountName,
        CurrencyType CurrencyType,
        decimal Amount);

    record SnapshotHoldingData(
        int AccountId,
        string AccountName,
        string Code,
        HoldingType HoldingType,
        int Quantity,
        Currency CurrentPrice,
        Currency TotalPrice,
        string DisplayText,
        string Description);

    record RecordSourceSupplement(
        int RecordId,
        string SourceCode,
        string SourceAppend);

    class DatabaseCleanupResult
    {
        public DatabaseCleanupResult(
            Dictionary<string, int> beforeCounts,
            Dictionary<string, int> afterCounts,
            List<StatementImport> fixedStatementImports)
        {
            BeforeCounts = beforeCounts;
            AfterCounts = afterCounts;
            FixedStatementImports = fixedStatementImports;
        }

        public Dictionary<string, int> BeforeCounts { get; }
        public Dictionary<string, int> AfterCounts { get; }
        public List<StatementImport> FixedStatementImports { get; }
    }

    class StatementRecordHoldingImport
    {
        public StatementRecordHoldingImport(
            StatementImportProvider provider,
            DateTime time,
            string statementKey,
            Account holdingAccount,
            List<Record> records,
            List<Holding> holdings,
            List<AccountBalance> accountBalances,
            List<AccountBalance> beginningAccountBalances,
            List<AccountInternalId>? internalCardNos = null)
        {
            Provider = provider;
            Time = time;
            StatementKey = statementKey;
            HoldingAccount = holdingAccount;
            Records = records;
            Holdings = holdings;
            AccountBalances = accountBalances;
            BeginningAccountBalances = beginningAccountBalances;
            InternalCardNos = internalCardNos ?? [];
        }

        public StatementImportProvider Provider { get; }
        public DateTime Time { get; }
        public string StatementKey { get; }
        public Account HoldingAccount { get; }
        public List<Record> Records { get; }
        public List<Holding> Holdings { get; }
        public List<AccountBalance> AccountBalances { get; }
        public List<AccountBalance> BeginningAccountBalances { get; }
        public List<AccountInternalId> InternalCardNos { get; }
    }
}
