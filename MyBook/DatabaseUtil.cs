using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Reflection;

namespace MyBook
{
    class DatabaseUtil
    {
        private const string DefaultConnectionString = "server=localhost;port=3306;database=mybook;uid=root;pwd=;charset=utf8mb4;";
        private readonly SqlSugarClient db;
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(AccountBalance), typeof(Record), typeof(Holding), typeof(Finance), typeof(StatementImport)];

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
            ValidateAccountPrimaryRelations();
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
            IEnumerable<AccountBalance>? beginningAccountBalances = null)
        {
            var recordList = records.ToList();
            var accountBalanceList = accountBalances?.ToList() ?? [];
            var beginningAccountBalanceList = beginningAccountBalances?.ToList() ?? [];
            db.Ado.BeginTran();
            try
            {
                if (IsStatementImported(provider, time, statementKey))
                {
                    db.Ado.CommitTran();
                    return false;
                }

                ValidateBeginningAccountBalances(
                    provider,
                    beginningAccountBalanceList,
                    ShouldValidateBeginningAccountBalances(provider));
                ValidateRecordBalanceChanges(provider, recordList, beginningAccountBalanceList, accountBalanceList);
                var statementImportId = InsertStatementImport(provider, time, statementKey);
                SaveAccountBalancesCore(accountBalanceList);
                SaveRecordsCore(recordList, statementImportId);
                db.Ado.CommitTran();
                return true;
            }
            catch (Exception e)
            {
                db.Ado.RollbackTran();
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time, statementKey))
                    return false;
                throw;
            }
        }

        public List<bool> SaveStatementRecordsAndHoldingsOnce(IEnumerable<StatementRecordHoldingImport> imports)
        {
            var importList = imports.ToList();
            var saved = new List<bool>();
            var shouldValidateBeginningBalances = importList
                .Select(import => import.Provider)
                .Distinct()
                .ToDictionary(provider => provider, ShouldValidateBeginningAccountBalances);
            db.Ado.BeginTran();
            try
            {
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

                db.Ado.CommitTran();
                return saved;
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
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
                var expectedEnding = Currency.RoundMoney(beginning + change);
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

            return sums.ToDictionary(item => item.Key, item => Currency.RoundMoney(item.Value));
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
                balances[(account.Id, accountBalance.t)] = Currency.RoundMoney(accountBalance.v);
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
            db.Ado.BeginTran();
            try
            {
                SaveAccountHoldingsCore(account, holdingList);
                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
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

        public DatabaseCleanupResult CleanVolatileData()
        {
            var beforeCounts = ReadCleanupCounts();
            db.Ado.BeginTran();
            try
            {
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
                        where `provider` in ('IBKRReportMail', 'ICBCBillMail')
                        group by `provider`
                    ) fixedImport
                        on statementImport.`provider` = fixedImport.`provider`
                        and statementImport.`time` = fixedImport.`time`
                    where fixedImport.`provider` is null
                    """);
                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }

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
                ["Finance"] = db.Queryable<Finance>().Count()
            };
        }

        public void SaveFinance(Finance finance)
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
    }

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
