using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Globalization;
using System.Reflection;
using System.Text.Json;

namespace MyBook
{
    class DatabaseUtil
    {
        private const string DefaultConnectionString = "server=localhost;port=3306;database=mybook;uid=root;pwd=;charset=utf8mb4;";
        private const string DatabaseWriteLockName = "MyBook.DatabaseWrite";
        private const int DatabaseWriteLockTimeoutSeconds = 300;
        private const int CurrentSnapshotSchemaVersion = 1;
        private readonly SqlSugarClient db;
        private static readonly JsonSerializerOptions SnapshotJsonOptions = new(JsonSerializerDefaults.Web);
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(AccountBalance), typeof(Record), typeof(Holding), typeof(Finance), typeof(Snapshot), typeof(SnapshotItem), typeof(StatementImport)];
        private static readonly ForeignKeyDefinition[] ForeignKeys =
        [
            new("fk_Accounts_primaryAccount", "Accounts", "_primaryAccount_Id", "Accounts", "Id"),
            new("fk_AccountBalances_account", "AccountBalances", "_account_Id", "Accounts", "Id"),
            new("fk_Holdings_account", "Holdings", "_account_Id", "Accounts", "Id"),
            new("fk_Records_account", "Records", "_account_Id", "Accounts", "Id"),
            new("fk_Records_statementImport", "Records", "_statementImport_Id", "StatementImports", "Id"),
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
            var hasExternalBalances = accountBalanceList.Count > 0 || beginningAccountBalanceList.Count > 0;
            try
            {
                return ExecuteLockedTransaction(() =>
                {
                    if (IsStatementImported(provider, time, statementKey))
                        return false;

                    var statementImportId = InsertStatementImport(provider, time, statementKey);
                    if (hasExternalBalances)
                    {
                        ValidateBeginningAccountBalances(
                            provider,
                            beginningAccountBalanceList,
                            ShouldValidateBeginningAccountBalances(provider));
                        ValidateRecordBalanceChanges(provider, recordList, beginningAccountBalanceList, accountBalanceList);
                        SaveAccountBalancesCore(accountBalanceList);
                    }
                    else
                    {
                        ValidateRelativeBalanceRecords(provider, recordList);
                    }

                    SaveRecordsCore(recordList, statementImportId);
                    if (!hasExternalBalances)
                        AddAccountBalanceDeltas(recordList);

                    afterSaveInTransaction?.Invoke(statementImportId);
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
                var shouldValidateBeginningBalances = importList
                    .Select(import => import.Provider)
                    .Distinct()
                    .ToDictionary(provider => provider, ShouldValidateBeginningAccountBalances);

                foreach (var import in importList)
                {
                    if (IsStatementImported(import.Provider, import.Time, import.StatementKey))
                    {
                        saved.Add(false);
                        continue;
                    }

                    ValidateBeginningAccountBalances(
                        import.Provider,
                        import.BeginningAccountBalances,
                        shouldValidateBeginningBalances[import.Provider]);
                    ValidateRecordBalanceChanges(import.Provider, import.Records, import.BeginningAccountBalances, import.AccountBalances);
                    var statementImportId = InsertStatementImport(import.Provider, import.Time, import.StatementKey);
                    SaveAccountBalancesCore(import.AccountBalances);
                    SaveAccountHoldingsCore(import.HoldingAccount, import.Holdings);
                    SaveRecordsCore(import.Records, statementImportId);
                    saved.Add(true);
                }

                return saved;
            });
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
                    record.updateTime,
                    record.DestAccount,
                    record.isInternal,
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
            return CreateSnapshot(DateTime.Now, SnapshotSource.AutoDaily, BuildDailySnapshotKey(snapshotDate));
        }

        public Snapshot CreateSnapshot(DateTime time, SnapshotSource source = SnapshotSource.Manual, string? snapshotKey = null)
        {
            var key = String.IsNullOrWhiteSpace(snapshotKey)
                ? BuildSnapshotKey(source, time)
                : snapshotKey.Trim();

            return ExecuteLockedTransaction(() =>
            {
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

        public SnapshotData? GetDailySnapshot(DateTime date)
        {
            var key = BuildDailySnapshotKey(date);
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.source == SnapshotSource.AutoDaily && it.snapshotKey == key)
                .First();
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

        private static string BuildSnapshotKey(SnapshotSource source, DateTime time)
        {
            return source == SnapshotSource.AutoDaily
                ? BuildDailySnapshotKey(time)
                : time.ToString("O", CultureInfo.InvariantCulture);
        }

        private static string BuildDailySnapshotKey(DateTime date)
        {
            return date.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
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
                .Where(record => !record.isInternal && !record.isRefundMatched && record.date >= reasonFirstMonth && record.date < nextMonth)
                .ToList();
            var accountList = db.Queryable<Account>().ToList();
            var lifeAccountIds = accountList
                .Where(account => account.usage == AccountUsage.Life)
                .Select(account => account.Id)
                .ToHashSet();
            var lifeRecords = records
                .Where(record => lifeAccountIds.Contains(record._account_Id))
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
                .Where(record => !record.isInternal && !record.isRefundMatched && record.v > 0 && record.date < today.Date.AddDays(1))
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
                MonthlyFlowSeries = BuildMonthlyFlowSeries(records, firstMonth),
                RmbMonthlyFlowSeries = BuildRmbMonthlyFlowSeries(records, firstMonth, exchangeRates),
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

                    var snapshot = GetDailySnapshot(date);
                    if (snapshot is not null)
                        return BuildAssetSummaryBalanceSetFromSnapshot(date, snapshot);

                    var baseSnapshot = GetLatestDailySnapshotBefore(date);
                    if (baseSnapshot is null)
                        return new AssetSummaryBalanceSet(date, null, false, []);

                    return new AssetSummaryBalanceSet(
                        date,
                        baseSnapshot.Snapshot.time,
                        true,
                        BuildRolledForwardBalances(baseSnapshot, date));
                })
                .ToList();
        }

        private SnapshotData? GetLatestDailySnapshotBefore(DateTime date)
        {
            var key = BuildDailySnapshotKey(date);
            var snapshot = db.Queryable<Snapshot>()
                .Where(it => it.source == SnapshotSource.AutoDaily)
                .OrderByDescending(it => it.snapshotKey)
                .ToList()
                .FirstOrDefault(it => String.CompareOrdinal(it.snapshotKey, key) < 0);
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

        private List<AccountBalance> BuildRolledForwardBalances(SnapshotData baseSnapshot, DateTime targetDate)
        {
            var balances = baseSnapshot.AccountBalances
                .GroupBy(balance => (balance.AccountId, balance.CurrencyType))
                .ToDictionary(group => group.Key, group => group.Sum(balance => balance.Amount));
            var baseDate = baseSnapshot.Snapshot.time.Date;
            var targetEnd = targetDate.Date.AddDays(1);
            var records = db.Queryable<Record>()
                .Where(record => !record.isInternal
                    && !record.isRefundMatched
                    && record.date > baseDate
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

        public DatabaseCleanupResult CleanVolatileData()
        {
            var beforeCounts = ReadCleanupCounts();
            ExecuteLockedTransaction(() =>
            {
                db.Deleteable<SnapshotItem>().ExecuteCommand();
                db.Deleteable<Snapshot>().ExecuteCommand();
                db.Deleteable<Record>().ExecuteCommand();
                db.Deleteable<Holding>().ExecuteCommand();
                db.Deleteable<Finance>().ExecuteCommand();
                db.Deleteable<AccountBalance>().ExecuteCommand();
                db.Ado.ExecuteCommand("""
                    delete statementImport
                    from `StatementImports` statementImport
                    left join (
                        select `provider`, min(`time`) as `time`
                        from `StatementImports`
                        where `provider` in ('IBKRReportMail', 'ICBCBillMail', 'OCBCStatementMail')
                        group by `provider`
                    ) fixedImport
                        on statementImport.`provider` = fixedImport.`provider`
                        and statementImport.`time` = fixedImport.`time`
                    where fixedImport.`provider` is null
                    """);
            });

            return new DatabaseCleanupResult(
                beforeCounts,
                ReadCleanupCounts(),
                db.Queryable<StatementImport>()
                    .OrderBy(it => it.provider)
                    .OrderBy(it => it.time)
                    .ToList());
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
                .ToDictionary(group => group.Key, group => group.Sum(balance => balance.v));
            var holdingSums = holdings
                .GroupBy(holding => holding.currentPrice.t)
                .ToDictionary(group => group.Key, group => group.Sum(holding => holding.totalPrice.v));
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
            DateTime UpdateTime,
            string DestAccount,
            bool IsInternal,
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
            List<AccountBalance> beginningAccountBalances)
        {
            Provider = provider;
            Time = time;
            StatementKey = statementKey;
            HoldingAccount = holdingAccount;
            Records = records;
            Holdings = holdings;
            AccountBalances = accountBalances;
            BeginningAccountBalances = beginningAccountBalances;
        }

        public StatementImportProvider Provider { get; }
        public DateTime Time { get; }
        public string StatementKey { get; }
        public Account HoldingAccount { get; }
        public List<Record> Records { get; }
        public List<Holding> Holdings { get; }
        public List<AccountBalance> AccountBalances { get; }
        public List<AccountBalance> BeginningAccountBalances { get; }
    }
}
