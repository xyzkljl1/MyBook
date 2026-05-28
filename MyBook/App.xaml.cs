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

            if (e.Args.Any(arg => arg.Equals("--create-start-snapshot", StringComparison.OrdinalIgnoreCase)))
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var snapshot = new DatabaseUtil(config).CreateSnapshot(DateTime.Now, SnapshotSource.Start);
                Console.WriteLine($"Created start snapshot: {snapshot.Id} {snapshot.source} {snapshot.time:yyyy-MM-dd HH:mm:ss.ffffff} revision={snapshot.maxStatementImportId} effectiveDate={snapshot.effectiveDate:yyyy-MM-dd} key={snapshot.snapshotKey}");
                Shutdown();
                return;
            }

            var startupConfig = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            new DatabaseUtil(startupConfig).EnsureBootstrapSqlBackupIfChanged("startup");

            MainWindow = new MainWindow();
            MainWindow.Show();
        }

        private static void WriteBootstrapSqlBackupResult(BootstrapSqlBackupResult result)
        {
            Console.WriteLine($"Bootstrap schema SQL: {result.BootstrapPath}");
            Console.WriteLine($"Bootstrap fixed-data SQL: {result.FixedDataPath}");
            Console.WriteLine($"Backup directory: {result.BackupDirectory}");
            Console.WriteLine($"Hash: {result.Hash}");
            Console.WriteLine($"Bootstrap schema changed: {result.BootstrapChanged}");
            Console.WriteLine($"Bootstrap fixed-data changed: {result.FixedDataChanged}");
            Console.WriteLine($"Backup written: {result.BackupWritten}");
        }
    }
}
