using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
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
        private const string InitializationRecordSourcePrefix = "InitialRecord/";
        private const string InitialHoldingReason = "Initial holding";
        private const string InitialCashBalanceReason = "Initial cash balance";
        private const int BootstrapBackupRetention = 3;
        private const string BootstrapSqlRelativePath = "Database/bootstrap.sql";
        private const string BootstrapBackupDirectoryName = "DatabaseBackups";
        private readonly SqlSugarClient db;
        private static readonly JsonSerializerOptions SnapshotJsonOptions = new(JsonSerializerDefaults.Web);
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(AccountInternalId), typeof(AccountBalance), typeof(Record), typeof(Holding), typeof(Finance), typeof(Snapshot), typeof(SnapshotItem), typeof(StatementImport)];
        private static readonly Type[] SchemaTableTypes = [typeof(Account), typeof(AccountInternalId), typeof(Finance), typeof(StatementImport), typeof(Holding), typeof(Record), typeof(Snapshot), typeof(SnapshotItem)];
        private static readonly HashSet<string> SchemaViewNames = ["AccountBalances"];
        private static readonly ForeignKeyDefinition[] ForeignKeys =
        [
            new("fk_Accounts_primaryAccount", "Accounts", "_primaryAccount_Id", "Accounts", "Id"),
            new("fk_AccountInternalIds_account", "AccountInternalIds", "_account_Id", "Accounts", "Id"),
            new("fk_Holdings_account", "Holdings", "_account_Id", "Accounts", "Id"),
            new("fk_Records_account", "Records", "_account_Id", "Accounts", "Id"),
            new("fk_Records_holding", "Records", "_holding_Id", "Holdings", "Id"),
            new("fk_Records_statementImport", "Records", "_statementImport_Id", "StatementImports", "Id"),
            new("fk_Records_matchedRecord", "Records", "matchedRecordId", "Records", "Id"),
            new("fk_SnapshotItems_snapshot", "SnapshotItems", "_snapshot_Id", "Snapshots", "Id"),
            new("fk_SnapshotItems_account", "SnapshotItems", "_account_Id", "Accounts", "Id")
        ];

        public DatabaseUtil(IConfigurationRoot config)
        {
            db = CreateDatabaseClient(GetDatabaseConnectionString(config));
            DbValidateSchema();
            DbValidateForeignKeys();
            ValidateAccountPrimaryRelations();
            DbValidateAccountBalancesViewDefinition();
            ValidateAllAccountBalancesFromHoldings();
        }

        public static void DbRebuildFromBootstrapSql(IConfigurationRoot config)
        {
            var bootstrapPath = Path.Combine(FindWorkspaceRoot(), BootstrapSqlRelativePath);
            if (!File.Exists(bootstrapPath))
                throw new FileNotFoundException($"Missing bootstrap SQL: {bootstrapPath}", bootstrapPath);

            var db = CreateDatabaseClient(GetDatabaseConnectionString(config));
            db.DbMaintenance.CreateDatabase();
            var objectCount = db.Ado.GetInt("""
                select count(*)
                from information_schema.tables
                where table_schema = database()
                """);
            if (objectCount != 0)
                throw new InvalidOperationException("Refuse to rebuild database because the target database is not empty.");

            foreach (var statement in SplitSqlStatements(File.ReadAllText(bootstrapPath, Encoding.UTF8)))
                db.Ado.ExecuteCommand(statement);

            _ = new DatabaseUtil(config);
        }

        private static string GetDatabaseConnectionString(IConfigurationRoot config)
        {
            return config["database_connection"]
                ?? config.GetConnectionString("Default")
                ?? DefaultConnectionString;
        }

        private static SqlSugarClient CreateDatabaseClient(string connectionString)
        {
            return new SqlSugarClient(new ConnectionConfig
            {
                ConnectionString = connectionString,
                DbType = DbType.MySql,
                InitKeyType = InitKeyType.Attribute,
                IsAutoCloseConnection = true
            });
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

        public BootstrapSqlBackupResult EnsureBootstrapSqlBackupIfChanged(string reason)
        {
            var sql = ExecuteLockedTransaction(BuildBootstrapSql);
            var workspaceRoot = FindWorkspaceRoot();
            var bootstrapPath = Path.Combine(workspaceRoot, BootstrapSqlRelativePath);
            var backupDirectory = Path.Combine(workspaceRoot, BootstrapBackupDirectoryName);
            var hash = ComputeContentHash(sql);

            var bootstrapChanged = WriteTextIfChanged(bootstrapPath, sql);
            var backupWritten = WriteBootstrapBackupIfChanged(backupDirectory, sql, hash);
            PruneBootstrapBackups(backupDirectory);

            return new BootstrapSqlBackupResult(
                bootstrapPath,
                backupDirectory,
                hash,
                bootstrapChanged,
                backupWritten,
                reason);
        }

        private string BuildBootstrapSql()
        {
            DbValidateSchema();
            DbValidateForeignKeys();
            ValidateAccountPrimaryRelations();
            DbValidateAccountBalancesViewDefinition();
            ValidateAllAccountBalancesFromHoldings();

            var builder = new StringBuilder();
            builder.AppendLine("-- MyBook bootstrap SQL. Schema plus fixed data only.");
            builder.AppendLine("-- Generated by MyBook; keep this file in sync with database structure.");
            builder.AppendLine("SET NAMES utf8mb4;");
            builder.AppendLine("SET FOREIGN_KEY_CHECKS=0;");
            builder.AppendLine();

            foreach (var type in SchemaTableTypes)
            {
                builder.AppendLine(NormalizeCreateTableSql(GetTableCreateSql(GetTableName(type))));
                builder.AppendLine();
            }

            builder.AppendLine(NormalizeCreateViewSql(GetViewCreateSql("AccountBalances")));
            builder.AppendLine();

            AppendFixedDataSql(builder);
            builder.AppendLine("SET FOREIGN_KEY_CHECKS=1;");

            return builder.ToString().ReplaceLineEndings("\n");
        }

        private string GetTableCreateSql(string tableName)
        {
            var result = db.Ado.GetDataTable($"show create table `{tableName}`");
            if (result.Rows.Count == 0)
                throw new InvalidOperationException($"Database schema mismatch: missing table {tableName}");

            foreach (System.Data.DataColumn column in result.Columns)
            {
                if (String.Equals(column.ColumnName, "Create Table", StringComparison.OrdinalIgnoreCase))
                    return result.Rows[0][column]?.ToString() ?? "";
            }

            throw new InvalidOperationException($"Database schema mismatch: cannot read create SQL for table {tableName}");
        }

        private static string NormalizeCreateTableSql(string sql)
        {
            var normalized = sql.Trim().TrimEnd(';');
            normalized = Regex.Replace(normalized, @"\sAUTO_INCREMENT=\d+", "", RegexOptions.IgnoreCase);
            return normalized + ";";
        }

        private static string NormalizeCreateViewSql(string sql)
        {
            var normalized = sql.Trim().TrimEnd(';');
            normalized = Regex.Replace(
                normalized,
                @"\sDEFINER=`[^`]+`@`[^`]+`",
                "",
                RegexOptions.IgnoreCase);
            return normalized + ";";
        }

        private void AppendFixedDataSql(StringBuilder builder)
        {
            var accounts = db.Queryable<Account>()
                .OrderBy(account => account.Id)
                .ToList();
            AppendInsertSql(
                builder,
                "Accounts",
                ["Id", "name", "desc", "email", "relativeBalance", "isCredit", "usage", "_primaryAccount_Id"],
                accounts.Select(account => new[]
                {
                    SqlValue(account.Id),
                    SqlValue(account.name),
                    SqlValue(account.desc),
                    SqlValue(account.email),
                    SqlValue(account.relativeBalance),
                    SqlValue(account.isCredit),
                    SqlValue(account.usage),
                    SqlValue(account._primaryAccount_Id)
                }));

            var fixedImports = GetFixedStatementImports();
            AppendInsertSql(
                builder,
                "StatementImports",
                ["Id", "provider", "time", "statementKey"],
                fixedImports.Select(statementImport => new[]
                {
                    SqlValue(statementImport.Id),
                    SqlValue(statementImport.provider),
                    SqlValue(statementImport.time),
                    SqlValue(statementImport.statementKey)
                }));
        }

        private static void AppendInsertSql(
            StringBuilder builder,
            string tableName,
            IReadOnlyList<string> columns,
            IEnumerable<IReadOnlyList<string>> rows)
        {
            var rowList = rows.ToList();
            if (rowList.Count == 0)
            {
                builder.AppendLine($"-- No fixed data for `{tableName}`.");
                builder.AppendLine();
                return;
            }

            builder.Append("INSERT INTO `");
            builder.Append(tableName);
            builder.Append("` (");
            builder.Append(String.Join(", ", columns.Select(column => $"`{column}`")));
            builder.AppendLine(") VALUES");
            for (var i = 0; i < rowList.Count; i++)
            {
                builder.Append("  (");
                builder.Append(String.Join(", ", rowList[i]));
                builder.Append(i == rowList.Count - 1 ? ");" : "),");
                builder.AppendLine();
            }

            builder.AppendLine();
        }

        private List<StatementImport> GetFixedStatementImports()
        {
            return db.Ado.SqlQuery<StatementImport>("""
                select statementImport.`Id`, statementImport.`provider`, statementImport.`time`, statementImport.`statementKey`
                from `StatementImports` statementImport
                where statementImport.`statementKey` = ''
                order by statementImport.`Id`
                """);
        }

        private static string SqlValue(int value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static string SqlValue(int? value)
        {
            return value.HasValue ? SqlValue(value.Value) : "NULL";
        }

        private static string SqlValue(bool value)
        {
            return value ? "1" : "0";
        }

        private static string SqlValue(DateTime value)
        {
            return "'" + value.ToString("yyyy-MM-dd HH:mm:ss.ffffff", CultureInfo.InvariantCulture) + "'";
        }

        private static string SqlValue(Enum value)
        {
            return SqlValue(value.ToString());
        }

        private static string SqlValue(string? value)
        {
            if (value is null)
                return "NULL";

            return "'" + value
                .Replace("\\", "\\\\", StringComparison.Ordinal)
                .Replace("'", "''", StringComparison.Ordinal)
                + "'";
        }

        private static bool WriteTextIfChanged(string path, string content)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            if (File.Exists(path) && File.ReadAllText(path, Encoding.UTF8) == content)
                return false;

            WriteAllTextAtomic(path, content);
            return true;
        }

        private static bool WriteBootstrapBackupIfChanged(string backupDirectory, string sql, string hash)
        {
            Directory.CreateDirectory(backupDirectory);
            var latestBackup = Directory.GetFiles(backupDirectory, "bootstrap-*.sql")
                .Select(path => new FileInfo(path))
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .ThenByDescending(file => file.Name, StringComparer.Ordinal)
                .FirstOrDefault();
            if (latestBackup is not null && ComputeFileHash(latestBackup.FullName) == hash)
                return false;

            var backupPath = Path.Combine(
                backupDirectory,
                $"bootstrap-{DateTime.Now:yyyyMMdd-HHmmss-ffffff}-{hash[..12]}.sql");
            WriteAllTextAtomic(backupPath, sql);
            return true;
        }

        private static void PruneBootstrapBackups(string backupDirectory)
        {
            if (!Directory.Exists(backupDirectory))
                return;

            var keptHashes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var keepPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in Directory.GetFiles(backupDirectory, "bootstrap-*.sql")
                .Select(path => new FileInfo(path))
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .ThenByDescending(file => file.Name, StringComparer.Ordinal))
            {
                var hash = ComputeFileHash(file.FullName);
                if (keptHashes.Count < BootstrapBackupRetention && keptHashes.Add(hash))
                {
                    keepPaths.Add(file.FullName);
                    continue;
                }

                if (!keepPaths.Contains(file.FullName))
                    File.Delete(file.FullName);
            }
        }

        private static void WriteAllTextAtomic(string path, string content)
        {
            var tempPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(tempPath, content, new UTF8Encoding(false));
            File.Move(tempPath, path, true);
        }

        private static string ComputeFileHash(string path)
        {
            return ComputeContentHash(File.ReadAllText(path, Encoding.UTF8));
        }

        private static string ComputeContentHash(string content)
        {
            return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(content))).ToLowerInvariant();
        }

        private static List<string> SplitSqlStatements(string sql)
        {
            var statements = new List<string>();
            var current = new StringBuilder();
            var inSingleQuote = false;
            var inDoubleQuote = false;
            var inBacktick = false;
            var inLineComment = false;
            var inBlockComment = false;

            for (var i = 0; i < sql.Length; i++)
            {
                var c = sql[i];
                var next = i + 1 < sql.Length ? sql[i + 1] : '\0';

                if (inLineComment)
                {
                    current.Append(c);
                    if (c == '\n')
                        inLineComment = false;
                    continue;
                }

                if (inBlockComment)
                {
                    current.Append(c);
                    if (c == '*' && next == '/')
                    {
                        current.Append(next);
                        i++;
                        inBlockComment = false;
                    }
                    continue;
                }

                if (inSingleQuote)
                {
                    current.Append(c);
                    if (c == '\\' && next != '\0')
                    {
                        current.Append(next);
                        i++;
                    }
                    else if (c == '\'')
                    {
                        if (next == '\'')
                        {
                            current.Append(next);
                            i++;
                        }
                        else
                        {
                            inSingleQuote = false;
                        }
                    }
                    continue;
                }

                if (inDoubleQuote)
                {
                    current.Append(c);
                    if (c == '\\' && next != '\0')
                    {
                        current.Append(next);
                        i++;
                    }
                    else if (c == '"')
                    {
                        inDoubleQuote = false;
                    }
                    continue;
                }

                if (inBacktick)
                {
                    current.Append(c);
                    if (c == '`')
                        inBacktick = false;
                    continue;
                }

                if (c == '-' && next == '-')
                {
                    current.Append(c);
                    current.Append(next);
                    i++;
                    inLineComment = true;
                    continue;
                }

                if (c == '#')
                {
                    current.Append(c);
                    inLineComment = true;
                    continue;
                }

                if (c == '/' && next == '*')
                {
                    current.Append(c);
                    current.Append(next);
                    i++;
                    inBlockComment = true;
                    continue;
                }

                if (c == '\'')
                {
                    current.Append(c);
                    inSingleQuote = true;
                    continue;
                }

                if (c == '"')
                {
                    current.Append(c);
                    inDoubleQuote = true;
                    continue;
                }

                if (c == '`')
                {
                    current.Append(c);
                    inBacktick = true;
                    continue;
                }

                if (c == ';')
                {
                    var statement = current.ToString().Trim();
                    if (!String.IsNullOrWhiteSpace(statement))
                        statements.Add(statement);
                    current.Clear();
                    continue;
                }

                current.Append(c);
            }

            var tail = current.ToString().Trim();
            if (!String.IsNullOrWhiteSpace(tail))
                statements.Add(tail);

            return statements;
        }

        private static string FindWorkspaceRoot()
        {
            foreach (var start in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory }.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var directory = new DirectoryInfo(start);
                while (directory is not null)
                {
                    if (File.Exists(Path.Combine(directory.FullName, "AGENTS.md"))
                        && File.Exists(Path.Combine(directory.FullName, "MyBook", "MyBook.csproj")))
                        return directory.FullName;

                    directory = directory.Parent;
                }
            }

            return Directory.GetCurrentDirectory();
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
            IEnumerable<AccountInternalId>? internalCardNos = null,
            Action<int>? afterSaveInTransaction = null,
            bool forceValidateBeginningBalances = false)
        {
            var recordList = records.ToList();
            var accountBalanceList = accountBalances?.ToList() ?? [];
            var beginningAccountBalanceList = beginningAccountBalances?.ToList() ?? [];
            var internalCardNoList = internalCardNos?.ToList();
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
                        forceValidateBeginningBalances || ShouldValidateBeginningAccountBalances(provider),
                        internalCardNos: internalCardNoList,
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

        public bool MarkStatementProcessedOnce(
            StatementImportProvider provider,
            DateTime time,
            string statementKey,
            IEnumerable<AccountInternalId>? internalCardNos = null)
        {
            var internalCardNoList = internalCardNos?.ToList() ?? [];
            try
            {
                return ExecuteLockedTransaction(() =>
                {
                    if (IsStatementImported(provider, time, statementKey))
                        return false;

                    InsertStatementImport(provider, time, statementKey);
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
                        import.BeginningHoldings,
                        import.InternalCardNos);
                    if (!statementImportId.HasValue)
                    {
                        saved.Add(false);
                        continue;
                    }

                    savedStatementImportIds.Add(statementImportId.Value);
                    saved.Add(true);
                    shouldValidateBeginningBalances[import.Provider] = true;
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
            List<Holding>? beginningHoldings = null,
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

            var initializationRecords = BuildInitializationRecords(
                provider,
                time,
                statementKey,
                records,
                beginningAccountBalances,
                holdingAccount,
                beginningHoldings);
            var statementImportId = InsertStatementImport(provider, time, statementKey);

            if (holdingAccount is not null && holdings is not null)
            {
                ValidateBeginningAccountHoldingQuantities(provider, statementKey, holdingAccount, beginningHoldings ?? []);
                SaveAccountHoldingsCore(holdingAccount, holdings, GetAccountBalancesForAccount(accountBalances, holdingAccount));
                ValidateSavedAccountHoldingQuantities(provider, statementKey, holdingAccount, holdings);
            }

            if (hasExternalBalances)
                SaveCashHoldingsFromAccountBalances(accountBalances, holdingAccount);

            var recordsToSave = initializationRecords.Count == 0
                ? records
                : initializationRecords.Concat(records).ToList();
            SaveRecordsCore(recordsToSave, statementImportId);
            if (!hasExternalBalances)
                ApplyRecordDeltasToHoldings(recordsToSave);

            if (internalCardNos is not null)
                EnsureAccountInternalCardNos(internalCardNos);

            afterSaveInTransaction?.Invoke(statementImportId);
            return statementImportId;
        }

        public static bool IsInitializationRecord(Record record)
        {
            return record.Source.StartsWith(InitializationRecordSourcePrefix, StringComparison.Ordinal);
        }

        private List<Record> BuildInitializationRecords(
            StatementImportProvider provider,
            DateTime importTime,
            string statementKey,
            List<Record> records,
            List<AccountBalance> beginningAccountBalances,
            Account? holdingAccount,
            List<Holding>? beginningHoldings)
        {
            var result = new List<Record>();
            var (recordDate, postingDate) = GetInitializationRecordDates(records, importTime);
            var initializedAccountIds = new HashSet<int>();

            if (holdingAccount is not null && beginningHoldings is not null)
            {
                var account = GetPostingAccount(holdingAccount);
                if (!AccountHasHistory(account))
                {
                    foreach (var holding in beginningHoldings)
                    {
                        NormalizeHolding(holding);
                        var amount = holding.totalPrice;
                        var hasQuantity = !Holding.IsSingleValueAsset(holding.holdingType) && holding.quantity != 0;
                        if (amount.v == 0 && !hasQuantity)
                            continue;

                        result.Add(CreateInitialHoldingRecord(
                            provider,
                            statementKey,
                            account,
                            holding,
                            amount,
                            recordDate,
                            postingDate));
                    }

                    initializedAccountIds.Add(account.Id);
                }
            }

            foreach (var group in beginningAccountBalances
                .Where(balance =>
                {
                    if (balance.Account is null)
                        throw new InvalidOperationException("Beginning account balance account is required.");
                    return true;
                })
                .GroupBy(balance => GetPostingAccount(balance.Account!).Id))
            {
                if (initializedAccountIds.Contains(group.Key))
                    continue;

                var account = db.Queryable<Account>()
                    .Where(it => it.Id == group.Key)
                    .First()
                    ?? throw new InvalidOperationException($"Account not found: {group.Key}");
                if (AccountHasHistory(account))
                    continue;

                foreach (var balance in group)
                {
                    if (balance.v == 0)
                        continue;

                    result.Add(CreateInitialCashBalanceRecord(
                        provider,
                        statementKey,
                        account,
                        balance,
                        recordDate,
                        postingDate));
                }
            }

            return result;
        }

        private static (DateTime Date, DateTime PostingDate) GetInitializationRecordDates(
            List<Record> records,
            DateTime importTime)
        {
            if (records.Count == 0)
                return (importTime.Date, importTime.Date);

            return (
                records.Min(record => record.date).Date,
                records.Min(record => record.postingDate ?? record.date).Date);
        }

        public bool HasAccountHistory(Account account)
        {
            return AccountHasHistory(account);
        }

        private bool AccountHasHistory(Account account)
        {
            account = GetPostingAccount(account);
            return db.Queryable<Record>()
                    .Where(record => record._account_Id == account.Id)
                    .Any()
                || db.Queryable<Holding>()
                    .Where(holding => holding._account_Id == account.Id)
                    .Any();
        }

        private static Record CreateInitialHoldingRecord(
            StatementImportProvider provider,
            string statementKey,
            Account account,
            Holding holding,
            Currency amount,
            DateTime recordDate,
            DateTime postingDate)
        {
            var record = new Record
            {
                Account = account,
                Holding = new Holding(holding.code, holding.holdingType)
                {
                    Account = account,
                    desc = holding.desc,
                    displayText = holding.displayText
                },
                date = recordDate,
                postingDate = postingDate,
                updateTime = DateTime.Now,
                isInternal = true,
                HoldingQuantity = Holding.IsSingleValueAsset(holding.holdingType) ? 0 : holding.quantity,
                DestAccount = holding.displayText,
                Source = LimitRecordText(
                    $"{InitializationRecordSourcePrefix}{provider}; statementKey={statementKey}; account={account.name}; holding={holding.code}/{holding.holdingType}; quantity={holding.quantity}"),
                Reason = InitialHoldingReason
            };
            record.CopyFrom(amount);
            return record;
        }

        private static Record CreateInitialCashBalanceRecord(
            StatementImportProvider provider,
            string statementKey,
            Account account,
            AccountBalance balance,
            DateTime recordDate,
            DateTime postingDate)
        {
            var record = new Record
            {
                Account = account,
                date = recordDate,
                postingDate = postingDate,
                updateTime = DateTime.Now,
                isInternal = true,
                DestAccount = balance.t.ToString(),
                Source = LimitRecordText(
                    $"{InitializationRecordSourcePrefix}{provider}; statementKey={statementKey}; account={account.name}; currency={balance.t}"),
                Reason = InitialCashBalanceReason
            };
            record.CopyFrom(balance);
            return record;
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
            SaveRecordsCore([record], statementImportId);
            ApplyRecordDeltasToHoldings([record]);
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

            ApplyRecordDeltaToHolding(existing, -1);
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
            if (statementImport.provider == StatementImportProvider.Manual)
                existing._holding_Id = EnsureCashHolding(account, existing.t).Id;
            ValidateRecordHolding(existing, account);
            db.Updateable(existing).ExecuteCommand();
            ApplyRecordDeltaToHolding(existing, 1);
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
                && !IsUndeterminedAccount(namedAccount))
            {
                return GetPostingAccount(namedAccount);
            }

            var matchedAccount = FindAccountByInternalCardNoText(
                null,
                $"record {record.Id} internal transfer match",
                record.DestAccount);
            if (matchedAccount is null || IsUndeterminedAccount(matchedAccount))
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
                    record._holding_Id,
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
                ResolveRecordHolding(record, account);
                if (record._holding_Id <= 0)
                    throw new InvalidOperationException("Record holding is required.");
                record._statementImport_Id = statementImportId;
            }

            if (recordList.Count > 0)
                db.Insertable(recordList).ExecuteCommand();
        }

        private void ResolveRecordHolding(Record record, Account account)
        {
            if (record.Holding is null)
            {
                record._holding_Id = EnsureCashHolding(account, record.t).Id;
                return;
            }

            var holding = record.Holding;
            NormalizeHolding(holding);
            var holdingAccount = holding.Account is null ? account : GetPostingAccount(holding.Account);
            if (holdingAccount.Id != account.Id)
            {
                throw new InvalidOperationException(
                    $"Record holding account mismatch: recordAccount={account.name}, holdingAccount={holdingAccount.name}, holding={holding.code}/{holding.holdingType}");
            }

            var existing = db.Queryable<Holding>()
                .Where(it => it._account_Id == account.Id
                    && it.code == holding.code
                    && it.holdingType == holding.holdingType)
                .First();
            if (existing is null)
            {
                holding.Account = account;
                holding._account_Id = account.Id;
                holding.quantity = Holding.IsSingleValueAsset(holding.holdingType) ? 1 : 0;
                holding.currentPrice = new Currency(0, record.t);
                holding.Id = db.Insertable(holding).ExecuteReturnIdentity();
                existing = holding;
                ValidateAccountBalancesFromHoldings(account);
            }

            ValidateRecordHolding(record, account, existing);
            record._holding_Id = existing.Id;
        }

        private void ValidateRecordHolding(Record record, Account account)
        {
            var holding = db.Queryable<Holding>()
                .Where(it => it.Id == record._holding_Id)
                .First();
            if (holding is null)
                throw new InvalidOperationException($"Record holding not found: {record._holding_Id}");

            ValidateRecordHolding(record, account, holding);
        }

        private static void ValidateRecordHolding(Record record, Account account, Holding holding)
        {
            if (holding._account_Id != account.Id)
            {
                throw new InvalidOperationException(
                    $"Record holding account mismatch: recordAccount={account.name}, holding={holding.code}/{holding.holdingType}");
            }

            if (record.t != holding.currentPrice.t)
            {
                throw new InvalidOperationException(
                    $"Record holding currency mismatch: account={account.name}, holding={holding.code}/{holding.holdingType}/{holding.currentPrice.t}, record={record.t}");
            }
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

        private List<AccountBalance> GetAccountBalancesForAccount(List<AccountBalance> accountBalances, Account account)
        {
            var postingAccount = GetPostingAccount(account);
            return accountBalances
                .Where(accountBalance =>
                {
                    if (accountBalance.Account is null)
                        throw new InvalidOperationException("Account balance account is required.");

                    return GetPostingAccount(accountBalance.Account).Id == postingAccount.Id;
                })
                .ToList();
        }

        private void SaveCashHoldingsFromAccountBalances(List<AccountBalance> accountBalances, Account? explicitHoldingAccount)
        {
            if (accountBalances.Count == 0)
                return;

            var explicitHoldingAccountId = explicitHoldingAccount is null ? (int?)null : GetPostingAccount(explicitHoldingAccount).Id;
            var normalizedBalances = accountBalances
                .Select(accountBalance =>
                {
                    if (accountBalance.Account is null)
                        throw new InvalidOperationException("Account balance account is required.");

                    var account = GetPostingAccount(accountBalance.Account);
                    return new { Account = account, Balance = accountBalance };
                })
                .Where(item => item.Account.Id != explicitHoldingAccountId)
                .ToList();
            foreach (var group in normalizedBalances.GroupBy(item => item.Account.Id))
            {
                var account = group.First().Account;
                var balances = group.Select(item => item.Balance).ToList();
                var cashHoldings = group
                    .Select(item => CreateCashHoldingFromAccountBalance(item.Balance))
                    .ToList();
                SaveAccountHoldingsCore(account, cashHoldings, balances);
            }
        }

        private Holding CreateCashHoldingFromAccountBalance(AccountBalance accountBalance)
        {
            if (accountBalance.Account is null)
                throw new InvalidOperationException("Account balance account is required.");

            var account = GetPostingAccount(accountBalance.Account);
            return new Holding(accountBalance.t.ToString(), HoldingType.Cash)
            {
                Account = account,
                desc = $"{accountBalance.t} cash balance",
                displayText = accountBalance.t.ToString(),
                currentPrice = new Currency(accountBalance.v, accountBalance.t)
            };
        }

        private Holding EnsureCashHolding(Account account, CurrencyType currency)
        {
            account = GetPostingAccount(account);
            var code = currency.ToString();
            var existing = db.Queryable<Holding>()
                .Where(holding => holding._account_Id == account.Id
                    && holding.code == code
                    && holding.holdingType == HoldingType.Cash)
                .First();
            if (existing is not null)
                return existing;

            var holding = new Holding(code, HoldingType.Cash)
            {
                Account = account,
                _account_Id = account.Id,
                desc = $"{currency} cash balance",
                displayText = code,
                currentPrice = new Currency(0, currency)
            };
            holding.Id = db.Insertable(holding).ExecuteReturnIdentity();
            ValidateAccountBalancesFromHoldings(account);
            return holding;
        }

        private void ApplyRecordDeltasToHoldings(IEnumerable<Record> records)
        {
            var affectedAccountIds = new HashSet<int>();
            foreach (var record in records)
            {
                ApplyRecordDeltaToHolding(record, 1);
                affectedAccountIds.Add(record._account_Id);
            }

            foreach (var accountId in affectedAccountIds)
                ValidateAccountBalancesFromHoldings(accountId);
        }

        private void ApplyRecordDeltaToHolding(Record record, int direction)
        {
            if (record._holding_Id <= 0)
                throw new InvalidOperationException("Record holding is required.");

            var holding = db.Queryable<Holding>()
                .Where(it => it.Id == record._holding_Id)
                .First();
            if (holding is null)
                throw new InvalidOperationException($"Record holding not found: {record._holding_Id}");
            if (holding._account_Id != record._account_Id)
                throw new InvalidOperationException($"Record holding account mismatch: record={record.Id}, holding={holding.Id}");
            if (holding.currentPrice.t != record.t)
                throw new InvalidOperationException($"Record holding currency mismatch: record={record.Id}, holding={holding.code}/{holding.holdingType}");

            var oldQuantity = holding.quantity;
            var oldTotal = holding.totalPrice.v;
            var newQuantity = oldQuantity + direction * record.HoldingQuantity;
            var newTotal = oldTotal + direction * record.v;
            if (Holding.IsSingleValueAsset(holding.holdingType))
            {
                holding.quantity = 1;
                holding.currentPrice = new Currency(newTotal, record.t);
            }
            else
            {
                holding.quantity = newQuantity;
                holding.currentPrice = new Currency(newQuantity == 0 ? 0 : newTotal / newQuantity, record.t);
            }

            db.Updateable(holding).ExecuteCommand();
            ValidateAccountBalancesFromHoldings(holding._account_Id);
        }

        private void ValidateAccountBalancesFromHoldings(Account account)
        {
            ValidateAccountBalancesFromHoldings(GetPostingAccount(account).Id);
        }

        private void ValidateAccountBalancesFromHoldings(int accountId)
        {
            var holdingSums = db.Queryable<Holding>()
                .Where(holding => holding._account_Id == accountId)
                .ToList()
                .GroupBy(holding => holding.currentPrice.t)
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(holding => holding.totalPrice.v)));
            var balanceSums = db.Queryable<AccountBalance>()
                .Where(balance => balance._account_Id == accountId)
                .ToList()
                .GroupBy(balance => balance.t)
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(balance => balance.v)));

            foreach (var currency in holdingSums.Keys.Union(balanceSums.Keys))
            {
                var holdingTotal = holdingSums.TryGetValue(currency, out var holdingValue) ? holdingValue : 0;
                var balanceTotal = balanceSums.TryGetValue(currency, out var balanceValue) ? balanceValue : 0;
                if (holdingTotal != balanceTotal)
                {
                    throw new InvalidOperationException(
                        $"Account balance view mismatch: accountId={accountId}, currency={currency}, holdings={holdingTotal}, balances={balanceTotal}");
                }
            }
        }

        private void ValidateAllAccountBalancesFromHoldings()
        {
            var accountIds = db.Queryable<Holding>()
                .ToList()
                .Select(holding => holding._account_Id)
                .Concat(db.Queryable<AccountBalance>().ToList().Select(balance => balance._account_Id))
                .Distinct()
                .ToList();
            foreach (var accountId in accountIds)
                ValidateAccountBalancesFromHoldings(accountId);
        }

        public void SaveAccountHoldings(Account account, IEnumerable<Holding> holdings)
        {
            var holdingList = holdings.ToList();
            ExecuteLockedTransaction(() =>
            {
                SaveAccountHoldingsCore(account, holdingList);
            });
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
                        holding.Id,
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
                payload.HoldingId,
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
                .Where(account => account.usage != AccountUsage.Undetermined)
                .Select(account => account.Id)
                .ToHashSet();
            var monthlyRecords = records
                .Where(record => monthlyAccountIds.Contains(record._account_Id))
                .ToList();
            var balances = db.Queryable<AccountBalance>()
                .Where(balance => balance.v != 0)
                .ToList();
            var holdings = db.Queryable<Holding>().ToList();
            var cashHoldingIds = holdings
                .Where(holding => holding.holdingType == HoldingType.Cash)
                .Select(holding => holding.Id)
                .ToHashSet();
            var accountNetFlowRecords = db.Queryable<Record>()
                .Where(record => !record.isRefundMatched && record.date < today.Date.AddDays(1))
                .ToList()
                .Where(record => cashHoldingIds.Contains(record._holding_Id))
                .ToList();
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
                .Union(accountNetFlowRecords.Select(record => record.t))
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
                AccountNetFlows = BuildAccountNetFlowStatistics(
                    accountList,
                    accountNetFlowRecords,
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
            return dates
                .Select(date =>
                {
                    var result = BuildAccountBalancesAt(
                        date,
                        targetRevision,
                        today,
                        currentBalances,
                        HistoricalBalanceDateBasis.PostingDate);
                    return new AssetSummaryBalanceSet(
                        result.Date,
                        result.SnapshotTime,
                        result.HasData,
                        result.Balances);
                })
                .ToList();
        }

        public HistoricalAccountBalanceResult GetAccountBalancesAt(DateTime date)
        {
            return GetAccountBalancesAtByPostingDate(date);
        }

        public HistoricalAccountBalanceResult GetAccountBalancesAtByTransactionDate(DateTime date)
        {
            return BuildAccountBalancesAt(
                date,
                GetCurrentStatementImportRevision(),
                DateTime.Today,
                null,
                HistoricalBalanceDateBasis.TransactionDate);
        }

        public HistoricalAccountBalanceResult GetAccountBalancesAtByPostingDate(DateTime date)
        {
            return BuildAccountBalancesAt(
                date,
                GetCurrentStatementImportRevision(),
                DateTime.Today,
                null,
                HistoricalBalanceDateBasis.PostingDate);
        }

        private HistoricalAccountBalanceResult BuildAccountBalancesAt(
            DateTime date,
            int targetRevision,
            DateTime today,
            List<AccountBalance>? currentBalances,
            HistoricalBalanceDateBasis dateBasis)
        {
            var targetDate = date.Date;
            if (targetDate == today.Date)
            {
                var balances = currentBalances is null
                    ? db.Queryable<AccountBalance>().ToList().Select(CloneDashboardBalance).ToList()
                    : currentBalances.Select(CloneDashboardBalance).ToList();
                return new HistoricalAccountBalanceResult(
                    targetDate,
                    targetRevision,
                    true,
                    null,
                    null,
                    null,
                    balances);
            }

            var baseSnapshot = GetLatestBalanceSnapshotAtOrBefore(targetDate, targetRevision);
            if (baseSnapshot is null)
            {
                return new HistoricalAccountBalanceResult(
                    targetDate,
                    targetRevision,
                    false,
                    null,
                    null,
                    null,
                    []);
            }

            return new HistoricalAccountBalanceResult(
                targetDate,
                targetRevision,
                true,
                baseSnapshot.Snapshot.Id,
                baseSnapshot.Snapshot.time,
                baseSnapshot.Snapshot.maxStatementImportId,
                BuildRolledForwardBalances(baseSnapshot, targetDate, targetRevision, dateBasis));
        }

        private SnapshotData? GetLatestBalanceSnapshotAtOrBefore(DateTime date, int targetRevision)
        {
            var snapshots = db.Queryable<Snapshot>()
                .Where(it => it.maxStatementImportId >= 0
                    && it.maxStatementImportId <= targetRevision)
                .ToList();
            if (snapshots.Count == 0)
                return null;

            var snapshot = snapshots
                .Where(it => it.effectiveDate <= date.Date)
                .OrderByDescending(it => it.effectiveDate)
                .ThenByDescending(it => it.maxStatementImportId)
                .ThenByDescending(it => it.time)
                .FirstOrDefault()
                ?? snapshots
                    .OrderBy(it => it.Id)
                    .First();
            return ReadSnapshotData(snapshot);
        }

        private List<AccountBalance> BuildRolledForwardBalances(
            SnapshotData baseSnapshot,
            DateTime targetDate,
            int targetRevision,
            HistoricalBalanceDateBasis dateBasis)
        {
            var balances = baseSnapshot.AccountBalances
                .GroupBy(balance => (balance.AccountId, balance.CurrencyType))
                .ToDictionary(group => group.Key, group => group.Sum(balance => balance.Amount));
            var targetEnd = targetDate.Date.AddDays(1);
            var baseRevision = baseSnapshot.Snapshot.maxStatementImportId;
            // 快照表示某个StatementImport导入进度下的最新状态，而不是某个自然日的余额。
            // 因此只能向前滚动：从快照revision之后新增的导入里，取所有发生在目标日结束前的record。
            // 如果这些新导入的record发生在快照effectiveDate之前，也仍然要计入，因为快照创建时尚未包含它们。
            var recordQuery = db.Queryable<Record>()
                .Where(record => record._statementImport_Id > baseRevision
                    && record._statementImport_Id <= targetRevision);
            if (dateBasis == HistoricalBalanceDateBasis.TransactionDate)
                recordQuery = recordQuery.Where(record => record.date < targetEnd);

            var records = recordQuery
                .ToList()
                .Where(record => GetHistoricalBalanceRecordDate(record, dateBasis) < targetEnd)
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

        private static DateTime GetHistoricalBalanceRecordDate(
            Record record,
            HistoricalBalanceDateBasis dateBasis)
        {
            return dateBasis == HistoricalBalanceDateBasis.PostingDate
                ? record.postingDate ?? record.date
                : record.date;
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
                .Where(account => account.usage != AccountUsage.Undetermined)
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

        private List<AccountNetFlowStatistics> BuildAccountNetFlowStatistics(
            List<Account> accounts,
            List<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var accountTypesById = accounts
                .ToDictionary(account => account.Id, account => GetAccountType(account.name));
            var accountTypesByName = accounts
                .GroupBy(account => account.name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => GetAccountType(group.First().name), StringComparer.OrdinalIgnoreCase);
            var matchedRecordIds = records
                .Where(record => record.matchedRecordId.HasValue)
                .Select(record => record.matchedRecordId!.Value)
                .Distinct()
                .ToList();
            var matchedRecordAccountTypes = matchedRecordIds.Count == 0
                ? new Dictionary<int, string>()
                : db.Queryable<Record>()
                    .Where(record => matchedRecordIds.Contains(record.Id))
                    .ToList()
                    .Where(record => accountTypesById.ContainsKey(record._account_Id))
                    .ToDictionary(record => record.Id, record => accountTypesById[record._account_Id]);

            return records
                .Where(record => record.v != 0
                    && accountTypesById.ContainsKey(record._account_Id)
                    && IsExternalToAccountType(record, accountTypesById, accountTypesByName, matchedRecordAccountTypes))
                .GroupBy(record => accountTypesById[record._account_Id], StringComparer.OrdinalIgnoreCase)
                .Select(group => BuildAccountNetFlowStatistic(group.Key, group.ToList(), exchangeRates))
                .Where(statistic => statistic.HasMissingExchangeRate
                    ? statistic.CurrencyTotals.Any(total => total.Amount != 0)
                    : statistic.NetRmb != 0)
                .OrderByDescending(statistic => Math.Abs(statistic.NetRmb))
                .ThenBy(statistic => statistic.DisplayName)
                .ToList();
        }

        private static AccountNetFlowStatistics BuildAccountNetFlowStatistic(
            string accountType,
            List<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var totals = records
                .GroupBy(record => record.t)
                .Select(group =>
                {
                    var amount = group.Sum(record => record.v);
                    return new AccountNetFlowCurrencyTotal
                    {
                        Currency = group.Key,
                        Amount = amount,
                        RmbAmount = TryConvertToRmb(amount, group.Key, exchangeRates)
                    };
                })
                .Where(total => total.Amount != 0)
                .OrderByDescending(total => Math.Abs(total.RmbAmount ?? 0))
                .ThenBy(total => total.Currency)
                .ToList();

            return new AccountNetFlowStatistics
            {
                AccountPrefix = accountType,
                DisplayName = BuildAccountTypeDisplayName(accountType),
                NetRmb = RoundAccountNetFlowRmb(totals
                    .Where(total => total.RmbAmount.HasValue)
                    .Sum(total => total.RmbAmount!.Value)),
                HasMissingExchangeRate = totals.Any(total => !total.RmbAmount.HasValue),
                RecordCount = records.Count,
                CurrencyTotals = totals
            };
        }

        private static decimal RoundAccountNetFlowRmb(decimal value)
        {
            return Decimal.Round(value, 0, MidpointRounding.ToEven);
        }

        private static bool IsExternalToAccountType(
            Record record,
            Dictionary<int, string> accountTypesById,
            Dictionary<string, string> accountTypesByName,
            Dictionary<int, string> matchedRecordAccountTypes)
        {
            var accountType = accountTypesById[record._account_Id];
            var counterpartyType = ResolveCounterpartyAccountType(record, accountTypesByName, matchedRecordAccountTypes);
            return counterpartyType is null
                || !String.Equals(accountType, counterpartyType, StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveCounterpartyAccountType(
            Record record,
            Dictionary<string, string> accountTypesByName,
            Dictionary<int, string> matchedRecordAccountTypes)
        {
            if (record.matchedRecordId.HasValue
                && matchedRecordAccountTypes.TryGetValue(record.matchedRecordId.Value, out var matchedAccountType))
            {
                return matchedAccountType;
            }

            if (!TryResolveCounterpartyName(record, out var counterpartyName))
                return null;

            return accountTypesByName.TryGetValue(counterpartyName, out var accountType)
                ? accountType
                : null;
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
                .Where(match => !IsUndeterminedAccount(match.Account))
                .ToList();
            return knownMatches.Count > 0 ? knownMatches : matches;
        }

        private static bool IsUndeterminedAccount(Account account)
        {
            return account.usage == AccountUsage.Undetermined;
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
                delete from `StatementImports`
                where `statementKey` <> ''
                """);
            if (startSnapshot is null)
            {
                db.Deleteable<Holding>().ExecuteCommand();
                return;
            }

            RestoreCurrentStateFromSnapshot(ReadSnapshotData(startSnapshot));
        }

        private Snapshot? GetStartSnapshot()
        {
            return db.Queryable<Snapshot>()
                .Where(snapshot => snapshot.source == SnapshotSource.Start)
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
            var holdings = snapshot.Holdings
                .Select(CreateHoldingFromSnapshot)
                .ToList();
            ValidateRestoredHoldings(snapshot, holdings);
            RestoreHoldingsFromSnapshot(snapshot, holdings);
            ValidateCurrentBalancesMatchSnapshot(snapshot);
        }

        private static Holding CreateHoldingFromSnapshot(SnapshotHoldingData snapshotHolding)
        {
            var holding = new Holding
            {
                Id = snapshotHolding.HoldingId,
                _account_Id = snapshotHolding.AccountId,
                code = snapshotHolding.Code,
                holdingType = snapshotHolding.HoldingType,
                quantity = snapshotHolding.Quantity,
                _currentPrice_v = snapshotHolding.CurrentPrice.v,
                _currentPrice_t = snapshotHolding.CurrentPrice.t,
                displayText = snapshotHolding.DisplayText,
                desc = snapshotHolding.Description
            };
            if (holding.totalPrice != snapshotHolding.TotalPrice)
            {
                throw new InvalidOperationException(
                    $"Snapshot holding total mismatch: account={snapshotHolding.AccountName}, code={snapshotHolding.Code}, type={snapshotHolding.HoldingType}, snapshot={snapshotHolding.TotalPrice.v}/{snapshotHolding.TotalPrice.t}, restored={holding.totalPrice.v}/{holding.totalPrice.t}");
            }

            return holding;
        }

        private void RestoreHoldingsFromSnapshot(SnapshotData snapshot, List<Holding> snapshotHoldings)
        {
            var existingHoldings = db.Queryable<Holding>().ToList();
            var affectedAccountIds = existingHoldings
                .Select(holding => holding._account_Id)
                .Concat(snapshot.AccountBalances.Select(balance => balance.AccountId))
                .Concat(snapshotHoldings.Select(holding => holding._account_Id))
                .ToHashSet();
            var usedHoldingIds = new HashSet<int>();

            foreach (var snapshotHolding in snapshotHoldings)
            {
                var existing = snapshotHolding.Id > 0
                    ? existingHoldings.FirstOrDefault(holding => holding.Id == snapshotHolding.Id)
                    : null;
                existing ??= existingHoldings.FirstOrDefault(holding =>
                    holding._account_Id == snapshotHolding._account_Id
                    && holding.code == snapshotHolding.code
                    && holding.holdingType == snapshotHolding.holdingType);

                if (existing is null)
                {
                    InsertSnapshotHolding(snapshotHolding);
                    usedHoldingIds.Add(snapshotHolding.Id);
                    continue;
                }

                existing._account_Id = snapshotHolding._account_Id;
                existing.code = snapshotHolding.code;
                existing.holdingType = snapshotHolding.holdingType;
                existing.quantity = snapshotHolding.quantity;
                existing._currentPrice_v = snapshotHolding._currentPrice_v;
                existing._currentPrice_t = snapshotHolding._currentPrice_t;
                existing.displayText = snapshotHolding.displayText;
                existing.desc = snapshotHolding.desc;
                db.Updateable(existing).ExecuteCommand();
                usedHoldingIds.Add(existing.Id);
            }

            var referencedHoldingIds = db.Queryable<Record>()
                .Select(record => record._holding_Id)
                .ToList()
                .ToHashSet();
            foreach (var staleHolding in existingHoldings.Where(holding => !usedHoldingIds.Contains(holding.Id)))
            {
                if (referencedHoldingIds.Contains(staleHolding.Id))
                {
                    staleHolding.quantity = Holding.IsSingleValueAsset(staleHolding.holdingType) ? 1 : 0;
                    staleHolding.currentPrice = new Currency(0, staleHolding.currentPrice.t);
                    db.Updateable(staleHolding).ExecuteCommand();
                    continue;
                }

                db.Deleteable<Holding>()
                    .Where(holding => holding.Id == staleHolding.Id)
                    .ExecuteCommand();
            }

            foreach (var accountId in affectedAccountIds)
                ValidateAccountBalancesFromHoldings(accountId);
        }

        private void InsertSnapshotHolding(Holding holding)
        {
            if (holding.Id <= 0)
            {
                holding.Id = db.Insertable(holding).ExecuteReturnIdentity();
                return;
            }

            db.Ado.ExecuteCommand("""
                insert into `Holdings`
                    (`Id`, `code`, `holdingType`, `quantity`, `desc`, `displayText`, `_currentPrice_v`, `_currentPrice_t`, `_account_Id`)
                values
                    (@id, @code, @holdingType, @quantity, @desc, @displayText, @currentPriceValue, @currentPriceCurrency, @accountId)
                """,
                new SugarParameter("@id", holding.Id),
                new SugarParameter("@code", holding.code),
                new SugarParameter("@holdingType", holding.holdingType.ToString()),
                new SugarParameter("@quantity", holding.quantity),
                new SugarParameter("@desc", holding.desc),
                new SugarParameter("@displayText", holding.displayText),
                new SugarParameter("@currentPriceValue", holding.currentPrice.v),
                new SugarParameter("@currentPriceCurrency", holding.currentPrice.t.ToString()),
                new SugarParameter("@accountId", holding._account_Id));
        }

        private static void ValidateRestoredHoldings(SnapshotData snapshot, List<Holding> holdings)
        {
            var balanceSums = snapshot.AccountBalances
                .GroupBy(balance => (balance.AccountId, balance.CurrencyType))
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(balance => balance.Amount)));
            var holdingSums = holdings
                .GroupBy(holding => (holding._account_Id, holding.currentPrice.t))
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(holding => holding.totalPrice.v)));
            foreach (var key in balanceSums.Keys.Union(holdingSums.Keys))
            {
                var balance = balanceSums.TryGetValue(key, out var balanceValue) ? balanceValue : 0;
                var holding = holdingSums.TryGetValue(key, out var holdingValue) ? holdingValue : 0;
                if (balance != holding)
                {
                    throw new InvalidOperationException(
                        $"Snapshot holding balance mismatch: accountId={key.Item1}, currency={key.Item2}, snapshotBalance={balance}, holdings={holding}");
                }
            }
        }

        private void ValidateCurrentBalancesMatchSnapshot(SnapshotData snapshot)
        {
            var expected = snapshot.AccountBalances
                .GroupBy(balance => (balance.AccountId, balance.CurrencyType))
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(balance => balance.Amount)));
            var actual = db.Queryable<AccountBalance>()
                .ToList()
                .GroupBy(balance => (balance._account_Id, balance.t))
                .ToDictionary(group => group.Key, group => Currency.RoundMoney(group.Sum(balance => balance.v)));
            foreach (var key in expected.Keys.Union(actual.Keys))
            {
                var expectedValue = expected.TryGetValue(key, out var expectedBalance) ? expectedBalance : 0;
                var actualValue = actual.TryGetValue(key, out var actualBalance) ? actualBalance : 0;
                if (expectedValue != actualValue)
                {
                    throw new InvalidOperationException(
                        $"Restored account balance mismatch: accountId={key.Item1}, currency={key.Item2}, snapshot={expectedValue}, current={actualValue}");
                }
            }
        }

        public void CleanWiseImportedData()
        {
            ExecuteLockedTransaction(() =>
            {
                var wiseAccount = GetAccountByName("WISE");
                ClearRecordMatchesForStatementProvider(StatementImportProvider.WiseMail);
                var wiseImportIds = db.Queryable<StatementImport>()
                    .Where(import => import.provider == StatementImportProvider.WiseMail)
                    .Select(import => import.Id)
                    .ToList();
                var wiseRecords = wiseImportIds.Count == 0
                    ? []
                    : db.Queryable<Record>()
                        .Where(record => wiseImportIds.Contains(record._statementImport_Id))
                        .ToList();
                foreach (var record in wiseRecords)
                    ApplyRecordDeltaToHolding(record, -1);

                db.Ado.ExecuteCommand("""
                    delete record
                    from `Records` record
                    join `StatementImports` statementImport
                        on record.`_statementImport_Id` = statementImport.`Id`
                    where statementImport.`provider` = @provider
                    """,
                    new SugarParameter("@provider", StatementImportProvider.WiseMail.ToString()));
                DeleteUnreferencedAccountHoldings(wiseAccount.Id);
                ValidateAccountBalancesFromHoldings(wiseAccount.Id);
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

        private void DeleteUnreferencedAccountHoldings(int accountId)
        {
            var referencedHoldingIds = db.Queryable<Record>()
                .Where(record => record._account_Id == accountId)
                .Select(record => record._holding_Id)
                .ToList()
                .ToHashSet();
            var deleteIds = db.Queryable<Holding>()
                .Where(holding => holding._account_Id == accountId)
                .ToList()
                .Where(holding => !referencedHoldingIds.Contains(holding.Id))
                .Select(holding => holding.Id)
                .ToList();
            if (deleteIds.Count > 0)
                db.Deleteable<Holding>().In(deleteIds).ExecuteCommand();
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

        public List<Record> GetStatementRecords(StatementImportProvider provider, Account account)
        {
            var postingAccount = GetPostingAccount(account);
            var importIds = db.Queryable<StatementImport>()
                .Where(statementImport => statementImport.provider == provider)
                .Select(statementImport => statementImport.Id)
                .ToList();
            if (importIds.Count == 0)
                return [];

            return db.Queryable<Record>()
                .Where(record => importIds.Contains(record._statementImport_Id)
                    && record._account_Id == postingAccount.Id)
                .ToList();
        }

        public List<StatementImport> GetStatementImports(StatementImportProvider provider)
        {
            return db.Queryable<StatementImport>()
                .Where(statementImport => statementImport.provider == provider)
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

        private void DbValidateSchema()
        {
            foreach (var type in SchemaTypes)
            {
                var tableName = GetTableName(type);
                var objectType = GetDatabaseObjectType(tableName);
                if (objectType is null)
                    throw new InvalidOperationException($"Database schema mismatch: missing table {tableName}");

                var expectedObjectType = SchemaViewNames.Contains(tableName)
                    ? "VIEW"
                    : "BASE TABLE";
                if (!objectType.Equals(expectedObjectType, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"Database schema mismatch: {tableName} must be {expectedObjectType}, actual {objectType}");
                }

                var dbColumns = GetDatabaseColumnNames(tableName);
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

        private string? GetDatabaseObjectType(string tableName)
        {
            var result = db.Ado.GetDataTable("""
                select `TABLE_TYPE`
                from `information_schema`.`TABLES`
                where `TABLE_SCHEMA` = database()
                    and `TABLE_NAME` = @tableName
                """,
                new SugarParameter("@tableName", tableName));
            return result.Rows.Count == 0 ? null : result.Rows[0]["TABLE_TYPE"]?.ToString();
        }

        private HashSet<string> GetDatabaseColumnNames(string tableName)
        {
            var result = db.Ado.GetDataTable("""
                select `COLUMN_NAME`
                from `information_schema`.`COLUMNS`
                where `TABLE_SCHEMA` = database()
                    and `TABLE_NAME` = @tableName
                """,
                new SugarParameter("@tableName", tableName));
            var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (System.Data.DataRow row in result.Rows)
            {
                var name = row["COLUMN_NAME"]?.ToString();
                if (!String.IsNullOrWhiteSpace(name))
                    columns.Add(name);
            }

            return columns;
        }

        private void DbValidateAccountBalancesViewDefinition()
        {
            var createView = GetViewCreateSql("AccountBalances");
            var normalized = NormalizeSqlForComparison(createView);
            var requiredFragments = new[]
            {
                "view accountbalances as select",
                "row_number() over",
                "from holdings",
                "group by",
                "_account_id",
                "_currentprice_t",
                "holdingtype = 'ust'",
                "round",
                "quantity *",
                "_currentprice_v",
                "having",
                "amount <> 0"
            };
            foreach (var fragment in requiredFragments)
            {
                if (!normalized.Contains(fragment, StringComparison.Ordinal))
                    throw new InvalidOperationException($"Database schema mismatch: AccountBalances view definition is not the expected holdings rollup.");
            }

            if (normalized.Contains("case _currentprice_t when", StringComparison.Ordinal)
                || normalized.Contains("case holdings._currentprice_t when", StringComparison.Ordinal))
                throw new InvalidOperationException("Database schema mismatch: AccountBalances view still uses hard-coded currency id mapping.");
        }

        private string GetViewCreateSql(string viewName)
        {
            var result = db.Ado.GetDataTable($"show create view `{viewName}`");
            if (result.Rows.Count == 0)
                throw new InvalidOperationException($"Database schema mismatch: missing view {viewName}");

            foreach (System.Data.DataColumn column in result.Columns)
            {
                if (String.Equals(column.ColumnName, "Create View", StringComparison.OrdinalIgnoreCase))
                    return result.Rows[0][column]?.ToString() ?? "";
            }

            throw new InvalidOperationException($"Database schema mismatch: cannot read create SQL for view {viewName}");
        }

        private static string NormalizeSqlForComparison(string sql)
        {
            return Regex.Replace(sql.Replace("`", ""), @"\s+", " ")
                .Trim()
                .ToLowerInvariant();
        }

        private void DbValidateForeignKeys()
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

        private void ValidateBeginningAccountHoldingQuantities(
            StatementImportProvider provider,
            string statementKey,
            Account account,
            List<Holding> beginningHoldings)
        {
            account = GetAccountByName(account.name);
            var existingHoldings = db.Queryable<Holding>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            if (existingHoldings.Count == 0)
                return;

            var currentQuantities = BuildHoldingQuantityMap(existingHoldings, "current holdings");
            var beginningQuantities = BuildHoldingQuantityMap(beginningHoldings, "beginning holdings");
            ValidateHoldingQuantityMaps(
                provider,
                statementKey,
                account,
                "Beginning holding quantity mismatch",
                "current",
                currentQuantities,
                "beginning",
                beginningQuantities);
        }

        private void ValidateSavedAccountHoldingQuantities(
            StatementImportProvider provider,
            string statementKey,
            Account account,
            List<Holding> endingHoldings)
        {
            account = GetAccountByName(account.name);
            var savedHoldings = db.Queryable<Holding>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            var savedQuantities = BuildHoldingQuantityMap(savedHoldings, "saved holdings");
            var endingQuantities = BuildHoldingQuantityMap(endingHoldings, "ending holdings");
            ValidateHoldingQuantityMaps(
                provider,
                statementKey,
                account,
                "Ending holding quantity mismatch",
                "saved",
                savedQuantities,
                "ending",
                endingQuantities);
        }

        private static Dictionary<string, HoldingQuantityData> BuildHoldingQuantityMap(
            IEnumerable<Holding> holdings,
            string context)
        {
            var quantities = new Dictionary<string, HoldingQuantityData>(StringComparer.OrdinalIgnoreCase);
            foreach (var holding in holdings)
            {
                if (Holding.IsSingleValueAsset(holding.holdingType))
                    continue;

                var key = GetHoldingKey(holding);
                if (quantities.ContainsKey(key))
                    throw new InvalidOperationException($"Duplicate {context}: {holding.code}/{holding.holdingType}");

                quantities[key] = new HoldingQuantityData(holding.code, holding.holdingType, holding.quantity);
            }

            return quantities;
        }

        private static void ValidateHoldingQuantityMaps(
            StatementImportProvider provider,
            string statementKey,
            Account account,
            string message,
            string leftName,
            Dictionary<string, HoldingQuantityData> left,
            string rightName,
            Dictionary<string, HoldingQuantityData> right)
        {
            foreach (var key in left.Keys.Union(right.Keys, StringComparer.OrdinalIgnoreCase))
            {
                left.TryGetValue(key, out var leftValue);
                right.TryGetValue(key, out var rightValue);
                var code = leftValue?.Code ?? rightValue?.Code ?? key;
                var holdingType = leftValue?.HoldingType ?? rightValue?.HoldingType ?? HoldingType.NASDAQ;
                var leftQuantity = leftValue?.Quantity ?? 0;
                var rightQuantity = rightValue?.Quantity ?? 0;
                if (leftQuantity != rightQuantity)
                {
                    throw new InvalidOperationException(
                        $"{message} for {provider}: statementKey={statementKey}, account={account.name}, code={code}, type={holdingType}, {leftName}={leftQuantity}, {rightName}={rightQuantity}");
                }
            }
        }

        private void SaveAccountHoldingsCore(Account account, List<Holding> holdingList, List<AccountBalance>? expectedBalances = null)
        {
            account = GetAccountByName(account.name);

            foreach (var holding in holdingList)
            {
                NormalizeHolding(holding);
                holding.Account = account;
                holding._account_Id = account.Id;
            }

            if (expectedBalances is not null)
                ValidateAccountHoldingsBalance(account, holdingList, expectedBalances);

            var existingHoldings = db.Queryable<Holding>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            var currentKeys = holdingList.Select(GetHoldingKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var staleHolding in existingHoldings.Where(holding => !currentKeys.Contains(GetHoldingKey(holding))))
            {
                staleHolding.quantity = Holding.IsSingleValueAsset(staleHolding.holdingType) ? 1 : 0;
                staleHolding.currentPrice = new Currency(0, staleHolding.currentPrice.t);
                db.Updateable(staleHolding).ExecuteCommand();
            }

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

            ValidateAccountBalancesFromHoldings(account);
        }

        private static void NormalizeHolding(Holding holding)
        {
            if (Holding.IsSingleValueAsset(holding.holdingType))
                holding.quantity = 1;
        }

        private void ValidateAccountHoldingsBalance(Account account, List<Holding> holdings, List<AccountBalance> accountBalances)
        {
            if (accountBalances.Count == 0)
                throw new InvalidOperationException($"Missing expected account balance for holdings: {account.name}");

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
            int HoldingId,
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
            int HoldingId,
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

        private sealed record HoldingQuantityData(
            string Code,
            HoldingType HoldingType,
            int Quantity);

        private enum HistoricalBalanceDateBasis
        {
            TransactionDate,
            PostingDate
        }

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
        int HoldingId,
        int AccountId,
        string AccountName,
        string Code,
        HoldingType HoldingType,
        int Quantity,
        Currency CurrentPrice,
        Currency TotalPrice,
        string DisplayText,
        string Description);

    public sealed record HistoricalAccountBalanceResult(
        DateTime Date,
        int TargetRevision,
        bool HasData,
        int? SnapshotId,
        DateTime? SnapshotTime,
        int? SnapshotRevision,
        List<AccountBalance> Balances);

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

    sealed record BootstrapSqlBackupResult(
        string BootstrapPath,
        string BackupDirectory,
        string Hash,
        bool BootstrapChanged,
        bool BackupWritten,
        string Reason);

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
            List<Holding>? beginningHoldings = null,
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
            BeginningHoldings = beginningHoldings ?? [];
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
        public List<Holding> BeginningHoldings { get; }
        public List<AccountInternalId> InternalCardNos { get; }
    }
}
