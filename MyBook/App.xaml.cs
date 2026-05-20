using Microsoft.Extensions.Configuration;
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
    }
}
