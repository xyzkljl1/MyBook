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
                var cleanupResult = new DatabaseUtil(config).CleanVolatileData();
                WriteCleanupResult(cleanupResult);
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

            if (e.Args.Any(arg => arg.Equals("--fetch-ibkr-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    Console.WriteLine("FetchIBKRReports: load config");
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    Console.WriteLine("FetchIBKRReports: open database");
                    var database = new DatabaseUtil(config);
                    Console.WriteLine("FetchIBKRReports: create mail util");
                    var mail = new MailUtil(config, database);
                    Console.WriteLine("FetchIBKRReports: start");
                    Task.Run(mail.FetchIBKRReports).WaitAsync(TimeSpan.FromSeconds(60)).GetAwaiter().GetResult();
                    Console.WriteLine("FetchIBKRReports: done");
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"FetchIBKRReports failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--fetch-wise-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    Console.WriteLine("FetchWiseReports: load config");
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    Console.WriteLine("FetchWiseReports: open database");
                    var database = new DatabaseUtil(config);
                    Console.WriteLine("FetchWiseReports: create mail util");
                    var mail = new MailUtil(config, database);
                    var wiseMonth = GetArgumentValue(e.Args, "--wise-month");
                    Console.WriteLine(String.IsNullOrWhiteSpace(wiseMonth)
                        ? "FetchWiseReports: start"
                        : $"FetchWiseReports: start {wiseMonth}");
                    Task.Run(String.IsNullOrWhiteSpace(wiseMonth)
                            ? mail.FetchWiseReports
                            : () => mail.FetchWiseReports(ParseMonthArgument(wiseMonth)))
                        .WaitAsync(TimeSpan.FromMinutes(20))
                        .GetAwaiter()
                        .GetResult();
                    Console.WriteLine("FetchWiseReports: done");
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"FetchWiseReports failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--fetch-ocbc-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    Console.WriteLine("FetchOCBCReports: load config");
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    Console.WriteLine("FetchOCBCReports: open database");
                    var database = new DatabaseUtil(config);
                    Console.WriteLine("FetchOCBCReports: create mail util");
                    var mail = new MailUtil(config, database);
                    var ocbcMonth = GetArgumentValue(e.Args, "--ocbc-month");
                    Console.WriteLine(String.IsNullOrWhiteSpace(ocbcMonth)
                        ? "FetchOCBCReports: start"
                        : $"FetchOCBCReports: start {ocbcMonth}");
                    Task.Run(String.IsNullOrWhiteSpace(ocbcMonth)
                            ? mail.FetchOCBCReports
                            : () => mail.FetchOCBCReports(ParseMonthArgument(ocbcMonth)))
                        .WaitAsync(TimeSpan.FromMinutes(20))
                        .GetAwaiter()
                        .GetResult();
                    Console.WriteLine("FetchOCBCReports: done");
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"FetchOCBCReports failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--fetch-paypal-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    Console.WriteLine("FetchPayPalReports: load config");
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    Console.WriteLine("FetchPayPalReports: open database");
                    var database = new DatabaseUtil(config);
                    Console.WriteLine("FetchPayPalReports: create mail util");
                    var mail = new MailUtil(config, database);
                    var paypalMonth = GetArgumentValue(e.Args, "--paypal-month");
                    Console.WriteLine(String.IsNullOrWhiteSpace(paypalMonth)
                        ? "FetchPayPalReports: start"
                        : $"FetchPayPalReports: start {paypalMonth}");
                    Task.Run(String.IsNullOrWhiteSpace(paypalMonth)
                            ? mail.FetchPayPalReports
                            : () => mail.FetchPayPalReports(ParseMonthArgument(paypalMonth)))
                        .WaitAsync(TimeSpan.FromMinutes(10))
                        .GetAwaiter()
                        .GetResult();
                    Console.WriteLine("FetchPayPalReports: done");
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"FetchPayPalReports failed: {exception.Message}");
                }

                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            MainWindow = new MainWindow();
            MainWindow.Show();
        }

        private static void WriteCleanupResult(DatabaseCleanupResult result)
        {
            Console.WriteLine("Before: " + FormatCounts(result.BeforeCounts));
            Console.WriteLine("After: " + FormatCounts(result.AfterCounts));
            Console.WriteLine("Fixed StatementImports:");
            foreach (var statementImport in result.FixedStatementImports)
                Console.WriteLine($"{statementImport.Id} {statementImport.provider} {statementImport.time:yyyy-MM-dd HH:mm:ss.ffffff}");
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
            if (DateTime.TryParseExact(
                    value,
                    ["yyyy-MM", "yyyy/MM", "yyyy-MM-dd", "yyyy/MM/dd"],
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var month))
                return month;

            throw new ArgumentException($"Invalid month argument: {value}");
        }
    }
}
