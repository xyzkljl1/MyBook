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

        public Account GetOrAddAccount(string? accountType, string id, string? secondaryId)
        {
            var accountName = BuildAccountName(accountType, id, secondaryId);
            var account = db.Queryable<Account>().First(it => it.name == accountName);
            if (account is not null)
                return account;

            account = new Account { name = accountName };
            account.Id = db.Insertable(account).ExecuteReturnIdentity();
            return account;
        }

        public static string BuildAccountName(string? accountType, string id, string? secondaryId)
        {
            var parts = new[] { accountType, id, secondaryId }
                .Where(part => !String.IsNullOrWhiteSpace(part))
                .Select(part => part!.Trim());
            return String.Join("_", parts);
        }

        private Account SaveAccount(Account account)
        {
            var existing = db.Queryable<Account>()
                .First(it => it.name == account.name);

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
