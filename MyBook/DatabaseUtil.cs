using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Reflection;

namespace MyBook
{
    class DatabaseUtil
    {
        private readonly SqlSugarClient db;
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(Record)];

        public DatabaseUtil(IConfigurationRoot config)
        {
            var connectionString = config["database_connection"]
                ?? config.GetConnectionString("Default")
                ?? "Data Source=mybook.db";
            var dbType = ReadDbType(config);
            if (dbType == DbType.Sqlite)
                SQLitePCL.Batteries.Init();

            db = new SqlSugarClient(new ConnectionConfig
            {
                ConnectionString = connectionString,
                DbType = dbType,
                InitKeyType = InitKeyType.Attribute,
                IsAutoCloseConnection = true,
                MoreSettings = new ConnMoreSettings
                {
                    SqliteCodeFirstEnableDropColumn = dbType == DbType.Sqlite // 仅Sqlite需要，开启时InitTables时会删除多余的列
                }
            });
#if DEBUG
            // 注意这并不是迁移，列改名时会删除旧列从而丢失所有数据
            db.CodeFirst.InitTables<Account, Record>();
#endif
            ValidateSchema();
        }

        public void SaveAccountsAndRecords(IEnumerable<Account> accounts, IEnumerable<Record> records)
        {
            db.Ado.BeginTran();
            try
            {
                foreach (var account in accounts)
                    SaveAccount(account);

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

        private Account SaveAccount(Account account)
        {
            var existing = db.Queryable<Account>()
                .First(it => it.name == account.name && it._v_t == account.v.t);

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

        private static DbType ReadDbType(IConfigurationRoot config)
        {
            var dbTypeText = config["database_type"] ?? "Sqlite";
            return dbTypeText.Trim().ToLowerInvariant() switch
            {
                "mysql" => DbType.MySql,
                "sqlite" => DbType.Sqlite,
                "sqllite" => DbType.Sqlite,
                _ => throw new ArgumentException($"Unsupported database_type: {dbTypeText}")
            };
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
