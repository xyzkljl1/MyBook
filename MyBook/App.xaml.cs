using Microsoft.Extensions.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows;

namespace MyBook
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if (e.Args.Any(arg => arg.Equals("--clean-database", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var cleanupResult = new DatabaseUtil(config).CleanVolatileData(ReadCleanToSnapshotId(e.Args));
                WriteCleanupResult(cleanupResult);
                Shutdown();
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--create-start-snapshot", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var snapshot = new DatabaseUtil(config).CreateSnapshot(DateTime.Now, SnapshotSource.Start);
                Console.WriteLine($"Created start snapshot: {snapshot.Id} {snapshot.source} {snapshot.time:yyyy-MM-dd HH:mm:ss.ffffff} revision={snapshot.maxStatementImportId} effectiveDate={snapshot.effectiveDate:yyyy-MM-dd} key={snapshot.snapshotKey}");
                Shutdown();
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--clean-wise-data", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                new DatabaseUtil(config).CleanWiseImportedData();
                Console.WriteLine("Clean Wise imported data done.");
                Shutdown();
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--export-bootstrap-sql", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    var result = new DatabaseUtil(config).EnsureBootstrapSqlBackupIfChanged("manual export");
                    WriteBootstrapSqlBackupResult(result);
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"Export bootstrap SQL failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--rebuild-database-from-bootstrap-sql", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    DatabaseUtil.DbRebuildFromBootstrapSql(config);
                    Console.WriteLine("Rebuilt database from bootstrap SQL.");
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"Rebuild database from bootstrap SQL failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--debug-sql", StringComparison.OrdinalIgnoreCase))
                || e.Args.Any(arg => arg.StartsWith("--debug-sql=", StringComparison.OrdinalIgnoreCase))
                || e.Args.Any(arg => arg.Equals("--debug-sql-file", StringComparison.OrdinalIgnoreCase))
                || e.Args.Any(arg => arg.StartsWith("--debug-sql-file=", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var sql = ReadDebugSqlArgument(e.Args);
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    var database = new DatabaseUtil(config);
                    if (IsQueryDebugSql(sql))
                    {
                        WriteDebugSqlResult(database.QueryDebugSql(sql));
                    }
                    else
                    {
                        var affectedRows = database.ExecuteDebugSql(sql);
                        Console.WriteLine($"Debug SQL done: affectedRows={affectedRows}");
                    }
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"Debug SQL failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--debug-fetch-local-ibkr-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var mail = new MailUtil(config, database);
                mail.DebugFetchLocalIBKRReports();
                Shutdown();
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--debug-fetch-local-wise-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var mail = new MailUtil(config, database);
                mail.DebugFetchLocalWiseReports(GetArgumentValue(e.Args, "--wise-local-dir"));
                Shutdown();
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--debug-download-icbc-history-details", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var dateText = GetArgumentValue(e.Args, "--icbc-history-date");
                    if (String.IsNullOrWhiteSpace(dateText))
                        throw new ArgumentException("Missing --icbc-history-date value.");

                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    var database = new DatabaseUtil(config);
                    var mail = new MailUtil(config, database);
                    Task.Run(() => mail.DebugDownloadICBCHistoryDetails(
                            ParseDateArgument(dateText),
                            GetArgumentValue(e.Args, "--icbc-history-dir")))
                        .WaitAsync(TimeSpan.FromMinutes(5))
                        .GetAwaiter()
                        .GetResult();
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"DebugDownloadICBCHistoryDetails failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--debug-download-latest-icbc-history-detail", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    var database = new DatabaseUtil(config);
                    var mail = new MailUtil(config, database);
                    Task.Run(() => mail.DebugDownloadLatestICBCHistoryDetail(GetArgumentValue(e.Args, "--icbc-history-dir")))
                        .WaitAsync(TimeSpan.FromMinutes(5))
                        .GetAwaiter()
                        .GetResult();
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"DebugDownloadLatestICBCHistoryDetail failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--debug-fetch-local-icbc-history-details", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    var database = new DatabaseUtil(config);
                    var mail = new MailUtil(config, database);
                    mail.DebugFetchLocalICBCHistoryDetails(GetArgumentValue(e.Args, "--icbc-history-dir"));
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"DebugFetchLocalICBCHistoryDetails failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (RunMailFetchCommand(
                    e.Args,
                    "--fetch-icbc-bills",
                    "FetchICBCBills",
                    "--icbc-month",
                    TimeSpan.FromMinutes(20),
                    async (mail, month) =>
                    {
                        if (String.IsNullOrWhiteSpace(month))
                        {
                            await mail.FetchICBCBills();
                            return;
                        }

                        var targetMonth = ParseMonthArgument(month);
                        if (!await mail.FetchICBCBill(targetMonth))
                            throw new InvalidOperationException($"Missing ICBC bill for {targetMonth:yyyy-MM}");
                    }))
                return;

            if (RunMailFetchCommand(
                    e.Args,
                    "--fetch-icbc-history-details",
                    "FetchICBCHistoryDetails",
                    "--icbc-history-since",
                    TimeSpan.FromMinutes(20),
                    (mail, since) => mail.FetchICBCHistoryDetails(
                        String.IsNullOrWhiteSpace(since)
                            ? DateTime.Today.AddMonths(-5)
                            : ParseDateArgument(since))))
                return;

            if (RunMailFetchCommand(
                    e.Args,
                    "--fetch-ibkr-reports",
                    "FetchIBKRReports",
                    null,
                    TimeSpan.FromSeconds(60),
                    (mail, _) => mail.FetchIBKRReports()))
                return;

            if (RunMailFetchCommand(
                    e.Args,
                    "--fetch-wise-reports",
                    "FetchWiseReports",
                    "--wise-month",
                    TimeSpan.FromMinutes(20),
                    (mail, month) => String.IsNullOrWhiteSpace(month)
                        ? mail.FetchWiseReports()
                        : mail.FetchWiseReports(ParseMonthArgument(month))))
                return;

            if (RunMailFetchCommand(
                    e.Args,
                    "--fetch-ocbc-reports",
                    "FetchOCBCReports",
                    "--ocbc-month",
                    TimeSpan.FromMinutes(20),
                    (mail, month) => String.IsNullOrWhiteSpace(month)
                        ? mail.FetchOCBCReports()
                        : mail.FetchOCBCReports(ParseMonthArgument(month))))
                return;

            if (RunMailFetchCommand(
                    e.Args,
                    "--fetch-paypal-reports",
                    "FetchPayPalReports",
                    "--paypal-month",
                    TimeSpan.FromMinutes(10),
                    (mail, month) => String.IsNullOrWhiteSpace(month)
                        ? mail.FetchPayPalReports()
                        : mail.FetchPayPalReports(ParseMonthArgument(month))))
                return;

            var startupConfig = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            new DatabaseUtil(startupConfig).EnsureBootstrapSqlBackupIfChanged("startup");

            MainWindow = new MainWindow();
            MainWindow.Show();
        }

        private bool RunMailFetchCommand(
            string[] args,
            string flag,
            string displayName,
            string? monthArgument,
            TimeSpan timeout,
            Func<MailUtil, string?, Task> fetch)
        {
            if (!args.Any(arg => arg.Equals(flag, StringComparison.OrdinalIgnoreCase)))
                return false;

            var exitCode = 0;
            try
            {
                Console.WriteLine($"{displayName}: load config");
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                Console.WriteLine($"{displayName}: open database");
                var database = new DatabaseUtil(config);
                Console.WriteLine($"{displayName}: create mail util");
                var mail = new MailUtil(config, database);
                var month = monthArgument is null ? null : GetArgumentValue(args, monthArgument);
                Console.WriteLine(String.IsNullOrWhiteSpace(month)
                    ? $"{displayName}: start"
                    : $"{displayName}: start {month}");
                Task.Run(() => fetch(mail, month))
                    .WaitAsync(timeout)
                    .GetAwaiter()
                    .GetResult();
                Console.WriteLine($"{displayName}: done");
            }
            catch (Exception exception)
            {
                exitCode = 1;
                Console.WriteLine($"{displayName} failed: {exception.Message}");
            }

            Shutdown(exitCode);
            Environment.Exit(exitCode);
            return true;
        }

        private static void WriteCleanupResult(DatabaseCleanupResult result)
        {
            Console.WriteLine("Before: " + FormatCounts(result.BeforeCounts));
            Console.WriteLine("After: " + FormatCounts(result.AfterCounts));
            Console.WriteLine("Fixed StatementImports:");
            foreach (var statementImport in result.FixedStatementImports)
                Console.WriteLine($"{statementImport.Id} {statementImport.provider} {statementImport.time:yyyy-MM-dd HH:mm:ss.ffffff}");
        }

        private static void WriteBootstrapSqlBackupResult(BootstrapSqlBackupResult result)
        {
            Console.WriteLine($"Bootstrap SQL: {result.BootstrapPath}");
            Console.WriteLine($"Backup directory: {result.BackupDirectory}");
            Console.WriteLine($"Hash: {result.Hash}");
            Console.WriteLine($"Bootstrap changed: {result.BootstrapChanged}");
            Console.WriteLine($"Backup written: {result.BackupWritten}");
        }

        private static int? ReadCleanToSnapshotId(string[] args)
        {
            var value = GetArgumentValue(args, "--clean-to-snapshot");
            if (String.IsNullOrWhiteSpace(value))
                return null;
            if (!Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var snapshotId) || snapshotId <= 0)
                throw new ArgumentException($"Invalid --clean-to-snapshot value: {value}");
            return snapshotId;
        }

        private static string FormatCounts(Dictionary<string, int> counts)
        {
            return String.Join(", ", counts.Select(item => $"{item.Key}={item.Value}"));
        }

        private static string? GetArgumentValue(string[] args, string name)
        {
            for (var i = 0; i < args.Length; i++)
            {
                if (args[i].Equals(name, StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                    return args[i + 1];

                var prefix = name + "=";
                if (args[i].StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return args[i][prefix.Length..];
            }

            return null;
        }

        private static string ReadDebugSqlArgument(string[] args)
        {
            var file = GetArgumentValue(args, "--debug-sql-file");
            if (!String.IsNullOrWhiteSpace(file))
                return File.ReadAllText(file);

            var sql = GetArgumentValue(args, "--debug-sql");
            if (!String.IsNullOrWhiteSpace(sql))
                return sql;

            throw new ArgumentException("Missing --debug-sql or --debug-sql-file value.");
        }

        private static bool IsQueryDebugSql(string sql)
        {
            var trimmed = TrimSqlLeadingTrivia(sql);
            return trimmed.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("show", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("describe", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("desc", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("with", StringComparison.OrdinalIgnoreCase);
        }

        private static string TrimSqlLeadingTrivia(string sql)
        {
            var current = sql.TrimStart();
            while (true)
            {
                if (current.StartsWith("--", StringComparison.Ordinal))
                {
                    var newline = current.IndexOf('\n');
                    current = newline < 0 ? "" : current[(newline + 1)..].TrimStart();
                    continue;
                }

                if (current.StartsWith("/*", StringComparison.Ordinal))
                {
                    var end = current.IndexOf("*/", StringComparison.Ordinal);
                    current = end < 0 ? "" : current[(end + 2)..].TrimStart();
                    continue;
                }

                return current;
            }
        }

        private static void WriteDebugSqlResult(DataTable table)
        {
            Console.WriteLine($"Rows: {table.Rows.Count}");
            if (table.Columns.Count == 0)
                return;

            Console.WriteLine(String.Join("\t", table.Columns.Cast<DataColumn>().Select(column => column.ColumnName)));
            foreach (DataRow row in table.Rows)
            {
                Console.WriteLine(String.Join(
                    "\t",
                    table.Columns.Cast<DataColumn>().Select(column => FormatDebugSqlValue(row[column]))));
            }
        }

        private static string FormatDebugSqlValue(object? value)
        {
            if (value is null || value == DBNull.Value)
                return "NULL";
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss.ffffff", CultureInfo.InvariantCulture);
            if (value is byte[] bytes)
                return Convert.ToHexString(bytes);

            return Convert.ToString(value, CultureInfo.InvariantCulture)?
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "    ")
                ?? "";
        }

        private static DateTime ParseMonthArgument(string value)
        {
            return ParseDateArgument(value);
        }

        private static DateTime ParseDateArgument(string value)
        {
            if (DateTime.TryParseExact(
                    value,
                    ["yyyy-MM", "yyyy/MM", "yyyy-MM-dd", "yyyy/MM/dd"],
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var date))
                return date;

            throw new ArgumentException($"Invalid date argument: {value}");
        }
    }
}
