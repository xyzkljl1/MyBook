using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Reflection;

namespace MyBook
{
    class DatabaseUtil
    {
        private const string DefaultConnectionString = "server=localhost;port=3306;database=mybook;uid=root;pwd=;charset=utf8mb4;";
        private readonly SqlSugarClient db;
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(Record), typeof(Stock), typeof(StatementImport)];

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

        public void SaveRecords(IEnumerable<Record> records)
        {
            var recordList = records.ToList();
            if (recordList.Count == 0)
                return;

            db.Ado.BeginTran();
            try
            {
                var statementImportId = InsertStatementImport(StatementImportProvider.Manual, DateTime.Now);
                SaveRecordsCore(recordList, statementImportId);
                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
        }

        public bool IsStatementImported(StatementImportProvider provider, DateTime time)
        {
            var importTime = NormalizeStatementImportTime(time);
            return db.Queryable<StatementImport>()
                .Any(it => it.provider == provider && it.time == importTime);
        }

        public DateTime? GetLatestStatementImportTime(StatementImportProvider provider)
        {
            var latestImport = db.Queryable<StatementImport>()
                .Where(it => it.provider == provider)
                .OrderByDescending(it => it.time)
                .First();
            return latestImport?.time;
        }

        public bool SaveStatementImportOnce(StatementImportProvider provider, DateTime time)
        {
            try
            {
                if (IsStatementImported(provider, time))
                    return false;

                InsertStatementImport(provider, time);
                return true;
            }
            catch (Exception e)
            {
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time))
                    return false;
                throw;
            }
        }

        public bool SaveStatementRecordsOnce(
            StatementImportProvider provider,
            DateTime time,
            IEnumerable<Record> records,
            IEnumerable<Account>? accountBalances = null)
        {
            var recordList = records.ToList();
            var accountBalanceList = accountBalances?.ToList() ?? [];
            db.Ado.BeginTran();
            try
            {
                if (IsStatementImported(provider, time))
                {
                    db.Ado.CommitTran();
                    return false;
                }

                var statementImportId = InsertStatementImport(provider, time);
                SaveAccountBalancesCore(accountBalanceList);
                SaveRecordsCore(recordList, statementImportId);
                db.Ado.CommitTran();
                return true;
            }
            catch (Exception e)
            {
                db.Ado.RollbackTran();
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time))
                    return false;
                throw;
            }
        }

        public Account GetAccountByTypeAndId(string? accountType, string id, CurrencyType currencyType)
        {
            return GetAccountByName(BuildAccountName(accountType, id), currencyType);
        }

        public Account GetAccountByName(string accountName, CurrencyType? currencyType = null)
        {
            var account = FindAccountByName(accountName, currencyType);
            if (account is null)
            {
                var currencyText = currencyType.HasValue ? currencyType.Value.ToString() : "Any";
                throw new InvalidOperationException($"Account not found: {accountName}/{currencyText}");
            }

            return account;
        }

        public List<Account> GetAllAccounts()
        {
            return db.Queryable<Account>().ToList();
        }

        private Account? FindAccountByName(string accountName, CurrencyType? currencyType)
        {
            var accountQuery = db.Queryable<Account>()
                .Where(it => it.name == accountName);

            if (currencyType.HasValue)
                accountQuery = accountQuery.Where(it => it._v_t == currencyType.Value);

            return accountQuery.First();
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
            var existing = GetAccountByName(account.name, account._v_t);
            account.Id = existing.Id;
            return existing;
        }

        public Account GetPostingAccount(Account account)
        {
            var current = GetExistingAccountByName(account);
            if (current._primaryAccount_Id is null)
                return current;

            if (current._primaryAccount_Id.Value == current.Id)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name}/{current._v_t} points to itself");

            var primary = db.Queryable<Account>()
                .Where(it => it.Id == current._primaryAccount_Id.Value)
                .First();
            if (primary is null)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name}/{current._v_t} points to missing account {current._primaryAccount_Id.Value}");
            if (primary._primaryAccount_Id is not null)
                throw new InvalidOperationException($"Invalid account primary relation: primary account {primary.name}/{primary._v_t} is also a supplementary account");
            if (primary._v_t != current._v_t)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name}/{current._v_t} points to {primary.name}/{primary._v_t}");

            return primary;
        }

        private int InsertStatementImport(StatementImportProvider provider, DateTime time)
        {
            return db.Insertable(new StatementImport
            {
                provider = provider,
                time = NormalizeStatementImportTime(time)
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

        private void SaveAccountBalancesCore(List<Account> accounts)
        {
            foreach (var account in accounts)
            {
                var existing = GetExistingAccountByName(account);
                existing._v_v = account._v_v;
                db.Updateable(existing).ExecuteCommand();
            }
        }

        public void SaveAccountStocks(Account account, IEnumerable<Stock> stocks)
        {
            account = GetAccountByName(account.name, account._v_t);

            var stockList = stocks.ToList();
            foreach (var stock in stockList)
            {
                stock.Account = account;
                stock._account_Id = account.Id;
            }

            db.Ado.BeginTran();
            try
            {
                var existingStocks = db.Queryable<Stock>()
                    .Where(it => it._account_Id == account.Id &&
                        (it.stockType == StockType.NASDAQ || it.stockType == StockType.ARCA || it.stockType == StockType.UST))
                    .ToList();
                var currentKeys = stockList.Select(GetStockKey)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                var deletedIds = existingStocks
                    .Where(stock => !currentKeys.Contains(GetStockKey(stock)))
                    .Select(stock => stock.Id)
                    .ToList();
                if (deletedIds.Count > 0)
                    db.Deleteable<Stock>().In(deletedIds).ExecuteCommand();

                foreach (var stock in stockList)
                {
                    var existing = existingStocks.FirstOrDefault(it =>
                        it.code == stock.code && it.stockType == stock.stockType);
                    if (existing is null)
                    {
                        db.Insertable(stock).ExecuteCommand();
                        continue;
                    }

                    existing.quantity = stock.quantity;
                    existing.desc = stock.desc;
                    existing.displayText = stock.displayText;
                    existing._currentPrice_t = stock._currentPrice_t;
                    existing._account_Id = account.Id;
                    db.Updateable(existing).ExecuteCommand();
                }

                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
        }

        public List<Stock> GetStocks()
        {
            return db.Queryable<Stock>()
                .Includes(it => it.Account)
                .ToList();
        }

        public DatabaseCleanupResult CleanVolatileData()
        {
            var beforeCounts = ReadCleanupCounts();
            db.Ado.BeginTran();
            try
            {
                db.Deleteable<Record>().ExecuteCommand();
                db.Deleteable<Stock>().ExecuteCommand();
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
                ["StatementImports"] = db.Queryable<StatementImport>().Count(),
                ["Records"] = db.Queryable<Record>().Count(),
                ["Stocks"] = db.Queryable<Stock>().Count()
            };
        }

        public void SaveStock(Stock stock)
        {
            if (stock.Account is not null)
                stock._account_Id = GetExistingAccountByName(stock.Account).Id;

            if (stock.Id <= 0)
            {
                stock.Id = db.Insertable(stock).ExecuteReturnIdentity();
                return;
            }

            db.Updateable(stock).ExecuteCommand();
        }

        public static DateTime NormalizeStatementImportTime(DateTime time)
        {
            return new DateTime(time.Ticks - time.Ticks % 10, time.Kind);
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
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name}/{account._v_t} points to itself");

                if (!accountsById.TryGetValue(primaryAccountId, out var primaryAccount))
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name}/{account._v_t} points to missing account {primaryAccountId}");

                if (primaryAccount._primaryAccount_Id.HasValue)
                    throw new InvalidOperationException($"Invalid account primary relation: primary account {primaryAccount.name}/{primaryAccount._v_t} is also a supplementary account");

                if (primaryAccount._v_t != account._v_t)
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name}/{account._v_t} points to {primaryAccount.name}/{primaryAccount._v_t}");
            }
        }

        private static string GetTableName(Type type)
        {
            if (type == typeof(Account))
                return "Accounts";
            if (type == typeof(Record))
                return "Records";
            if (type == typeof(Stock))
                return "Stocks";
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

        private static string GetStockKey(Stock stock)
        {
            return $"{stock.code}\t{stock.stockType}";
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
}
