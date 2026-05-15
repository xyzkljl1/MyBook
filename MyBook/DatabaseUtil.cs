using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Reflection;

namespace MyBook
{
    class DatabaseUtil
    {
        private const string DefaultConnectionString = "server=localhost;port=3306;database=mybook;uid=root;pwd=;charset=utf8mb4;";
        private readonly SqlSugarClient db;
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(Record)];

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
            DropObsoleteIndexes();
            db.CodeFirst.InitTables<Account, Record>();
#endif
            ValidateSchema();
        }

        public void SaveRecords(IEnumerable<Record> records)
        {
            db.Ado.BeginTran();
            try
            {
                var recordList = records.ToList();
                foreach (var record in recordList)
                {
                    if (record.Account is null)
                        continue;

                    var account = SaveAccount(record.Account);
                    record._account_Id = account.Id;
                }

                if (recordList.Count > 0)
                    db.Insertable(recordList).ExecuteCommand();

                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
        }

        public Account GetOrAddAccount(string? accountType, string id, CurrencyType currencyType)
        {
            var accountName = BuildAccountName(accountType, id);
            var accountQuery = db.Queryable<Account>()
                .Where(it => it.name == accountName);

            if (currencyType != CurrencyType.Any)
                accountQuery = accountQuery.Where(it => it._v_t == currencyType);

            var account = accountQuery.First();
            if (account is not null)
                return account;

            account = new Account { name = accountName, _v_t = currencyType };
            account.Id = db.Insertable(account).ExecuteReturnIdentity();
            return account;
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

        private void DropObsoleteIndexes()
        {
            DropIndexIfExists("Accounts", "unique_Accounts_name");
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
    }
}
