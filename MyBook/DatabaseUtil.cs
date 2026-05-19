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
#if DEBUG
            // This is schema sync, not migration; renamed columns may lose old data.
            PrepareSchemaSync();
            db.CodeFirst.InitTables(SchemaTypes);
            PrepareForeignKeys();
#endif
            ValidateSchema();
        }

        public void SaveRecords(IEnumerable<Record> records)
        {
            var recordList = records.ToList();
            db.Ado.BeginTran();
            try
            {
                SaveRecordsCore(recordList);
                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
        }

        public bool IsStatementImported(string provider, string month)
        {
            return db.Queryable<StatementImport>()
                .Any(it => it.provider == provider && it.month == month);
        }

        public bool SaveStatementRecordsOnce(string provider, string month, IEnumerable<Record> records)
        {
            var recordList = records.ToList();
            db.Ado.BeginTran();
            try
            {
                if (IsStatementImported(provider, month))
                {
                    db.Ado.CommitTran();
                    return false;
                }

                db.Insertable(new StatementImport { provider = provider, month = month })
                    .ExecuteCommand();
                SaveRecordsCore(recordList);
                db.Ado.CommitTran();
                return true;
            }
            catch (Exception e)
            {
                db.Ado.RollbackTran();
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, month))
                    return false;
                throw;
            }
        }

        public Account GetOrAddAccount(string? accountType, string id, CurrencyType currencyType)
        {
            var accountName = BuildAccountName(accountType, id);
            var account = FindAccount(accountName, currencyType);
            if (account is not null)
                return account;

            account = new Account { name = accountName, _v_t = currencyType };
            account.Id = db.Insertable(account).ExecuteReturnIdentity();
            return account;
        }

        public Account? FindAccount(string? accountType, string id, CurrencyType? currencyType = null)
        {
            return FindAccount(BuildAccountName(accountType, id), currencyType);
        }

        private Account? FindAccount(string accountName, CurrencyType? currencyType)
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

        private Account SaveAccount(Account account)
        {
            var existing = db.Queryable<Account>()
                .First(it => it.name == account.name && it._v_t == account._v_t);

            if (existing is null)
            {
                account.Id = db.Insertable(account).ExecuteReturnIdentity();
                return account;
            }

            account.Id = existing.Id;
            existing._v_v = account.v.v;
            existing._v_t = account.v.t;
            existing.desc = account.desc;
            db.Updateable(existing).ExecuteCommand();
            return existing;
        }

        private void SaveRecordsCore(List<Record> recordList)
        {
            foreach (var record in recordList)
            {
                if (record.Account is null)
                    continue;

                var account = SaveAccount(record.Account);
                record._account_Id = account.Id;
            }

            if (recordList.Count > 0)
                db.Insertable(recordList).ExecuteCommand();
        }

        public void SaveAccountStocks(Account account, IEnumerable<Stock> stocks)
        {
            if (account.Id <= 0)
            {
                var existingAccount = FindAccount(account.name, account._v_t);
                if (existingAccount is null)
                {
                    account.Id = db.Insertable(account).ExecuteReturnIdentity();
                }
                else
                {
                    account = existingAccount;
                }
            }

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

        public void SaveStock(Stock stock)
        {
            if (stock.Account is not null && stock._account_Id is null && stock.Account.Id > 0)
                stock._account_Id = stock.Account.Id;

            if (stock.Id <= 0)
            {
                stock.Id = db.Insertable(stock).ExecuteReturnIdentity();
                return;
            }

            db.Updateable(stock).ExecuteCommand();
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

        private void PrepareSchemaSync()
        {
            DropForeignKeyIfExists("Records", "fk_Records_Accounts_account_Id");
            DropForeignKeyIfExists("Stocks", "fk_Stocks_Accounts_account_Id");
            DropIndexIfExists("Accounts", "unique_Accounts_name");
            DropIndexIfExists("Stocks", "unique_Stocks_account_stockId");
            DropIndexIfExists("Stocks", "unique_Stocks_account_code_type");
            RenameColumnIfNeeded("Stocks", "stockId", "code", "varchar(255) not null default ''");
            RenameColumnIfNeeded("Stocks", "t", "stockType", "enum('US','NASDAQ','UST','SHANGHAI','CNFUND','Cash') not null default 'NASDAQ'");
            MigrateStockTypeUS();
            MigrateStockTypeEnum();
            RenameColumnIfNeeded("Stocks", "currentPrice", "_currentPrice_v", "decimal(18,4) not null default 0");
            MigrateStockCurrentPriceTime();
        }

        private void MigrateStockCurrentPriceTime()
        {
            if (!ColumnExists("Stocks", "currentPriceTime"))
                return;

            var dataType = GetColumnDataType("Stocks", "currentPriceTime");
            if (String.Equals(dataType, "bigint", StringComparison.OrdinalIgnoreCase))
                return;

            if (dataType is not null &&
                (dataType.Equals("int", StringComparison.OrdinalIgnoreCase) ||
                 dataType.Equals("integer", StringComparison.OrdinalIgnoreCase) ||
                 dataType.Equals("mediumint", StringComparison.OrdinalIgnoreCase) ||
                 dataType.Equals("smallint", StringComparison.OrdinalIgnoreCase) ||
                 dataType.Equals("tinyint", StringComparison.OrdinalIgnoreCase)))
            {
                db.Ado.ExecuteCommand("alter table `Stocks` modify column `currentPriceTime` bigint not null default 0");
                return;
            }

            if (ColumnExists("Stocks", "currentPriceTime_unix"))
                db.Ado.ExecuteCommand("alter table `Stocks` drop column `currentPriceTime_unix`");

            db.Ado.ExecuteCommand("alter table `Stocks` add column `currentPriceTime_unix` bigint not null default 0");
            db.Ado.ExecuteCommand("""
                update `Stocks`
                set `currentPriceTime_unix` = coalesce(unix_timestamp(`currentPriceTime`), 0)
                """);
            db.Ado.ExecuteCommand("alter table `Stocks` drop column `currentPriceTime`");
            db.Ado.ExecuteCommand("alter table `Stocks` change column `currentPriceTime_unix` `currentPriceTime` bigint not null default 0");
        }

        private void PrepareForeignKeys()
        {
            AddForeignKeyIfMissing("Records", "_account_Id", "Accounts", "Id", "fk_Records_Accounts_account_Id");
            AddForeignKeyIfMissing("Stocks", "_account_Id", "Accounts", "Id", "fk_Stocks_Accounts_account_Id");
        }

        private void AddForeignKeyIfMissing(
            string tableName,
            string columnName,
            string referencedTableName,
            string referencedColumnName,
            string foreignKeyName)
        {
            if (!db.DbMaintenance.IsAnyTable(tableName, false) ||
                !db.DbMaintenance.IsAnyTable(referencedTableName, false) ||
                !ColumnExists(tableName, columnName) ||
                !ColumnExists(referencedTableName, referencedColumnName) ||
                ForeignKeyExists(foreignKeyName))
            {
                return;
            }

            db.Ado.ExecuteCommand($"""
                update `{tableName}` child
                left join `{referencedTableName}` parent
                  on child.`{columnName}` = parent.`{referencedColumnName}`
                set child.`{columnName}` = null
                where child.`{columnName}` is not null
                  and parent.`{referencedColumnName}` is null
                """);

            db.Ado.ExecuteCommand($"""
                alter table `{tableName}`
                add constraint `{foreignKeyName}`
                foreign key (`{columnName}`)
                references `{referencedTableName}`(`{referencedColumnName}`)
                on delete set null
                on update cascade
                """);
        }

        private bool ForeignKeyExists(string foreignKeyName)
        {
            return db.Ado.SqlQuery<int>($"""
                select count(*)
                from information_schema.referential_constraints
                where constraint_schema = database()
                  and constraint_name = '{foreignKeyName}'
                """).FirstOrDefault() > 0;
        }

        private void DropIndexIfExists(string tableName, string indexName)
        {
            var exists = db.Ado.SqlQuery<int>($"""
                select count(*)
                from information_schema.statistics
                where table_schema = database()
                  and table_name = '{tableName}'
                  and index_name = '{indexName}'
                """).FirstOrDefault() > 0;

            if (exists)
                db.Ado.ExecuteCommand($"alter table `{tableName}` drop index `{indexName}`");
        }

        private void DropForeignKeyIfExists(string tableName, string foreignKeyName)
        {
            if (ForeignKeyExists(foreignKeyName))
                db.Ado.ExecuteCommand($"alter table `{tableName}` drop foreign key `{foreignKeyName}`");
        }

        private void RenameColumnIfNeeded(string tableName, string oldColumnName, string newColumnName, string columnDefinition)
        {
            var hasOldColumn = ColumnExists(tableName, oldColumnName);
            if (!hasOldColumn)
                return;

            if (ColumnExists(tableName, newColumnName))
            {
                db.Ado.ExecuteCommand($"alter table `{tableName}` drop column `{oldColumnName}`");
                return;
            }

            db.Ado.ExecuteCommand($"alter table `{tableName}` change column `{oldColumnName}` `{newColumnName}` {columnDefinition}");
        }

        private bool ColumnExists(string tableName, string columnName)
        {
            return db.Ado.SqlQuery<int>($"""
                select count(*)
                from information_schema.columns
                where table_schema = database()
                  and table_name = '{tableName}'
                  and column_name = '{columnName}'
                """).FirstOrDefault() > 0;
        }

        private string? GetColumnDataType(string tableName, string columnName)
        {
            return db.Ado.SqlQuery<string>($"""
                select data_type
                from information_schema.columns
                where table_schema = database()
                  and table_name = '{tableName}'
                  and column_name = '{columnName}'
                """).FirstOrDefault();
        }

        private void MigrateStockTypeUS()
        {
            if (!ColumnExists("Stocks", "stockType"))
                return;

            var columnType = db.Ado.SqlQuery<string>("""
                select column_type
                from information_schema.columns
                where table_schema = database()
                  and table_name = 'Stocks'
                  and column_name = 'stockType'
                """).FirstOrDefault();

            if (columnType is null || !columnType.Contains("'US'", StringComparison.OrdinalIgnoreCase))
                return;

            if (!columnType.Contains("'NASDAQ'", StringComparison.OrdinalIgnoreCase))
            {
                db.Ado.ExecuteCommand(
                    "alter table `Stocks` modify column `stockType` enum('US','NASDAQ','UST','SHANGHAI','CNFUND','Cash') not null default 'NASDAQ'");
            }

            db.Ado.ExecuteCommand("update `Stocks` set `stockType` = 'NASDAQ' where `stockType` = 'US'");
        }

        private void MigrateStockTypeEnum()
        {
            if (!ColumnExists("Stocks", "stockType"))
                return;

            var columnType = db.Ado.SqlQuery<string>("""
                select column_type
                from information_schema.columns
                where table_schema = database()
                  and table_name = 'Stocks'
                  and column_name = 'stockType'
                """).FirstOrDefault();

            if (columnType is not null && columnType.Contains("'ARCA'", StringComparison.OrdinalIgnoreCase))
                return;

            db.Ado.ExecuteCommand(
                "alter table `Stocks` modify column `stockType` enum('NASDAQ','ARCA','UST','SHANGHAI','CNFUND','Cash') not null default 'NASDAQ'");
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
}
