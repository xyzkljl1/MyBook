using Microsoft.Extensions.Configuration;
using System.Globalization;
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

            if (e.Args.Any(arg => arg.Equals("--debug-fetch-local-ibkr-reports", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var mail = new MailUtil(config, database);
                mail.DebugFetchLocalIBKRReports();
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
