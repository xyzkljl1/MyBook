using Microsoft.Extensions.Configuration;
using System.Windows;

namespace MyBook
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private const string SingleInstanceMutexName = @"Local\MyBook.SingleInstance";
        private const string ShowWindowEventName = @"Local\MyBook.ShowWindow";
        private Mutex? singleInstanceMutex;
        private EventWaitHandle? showWindowEvent;
        private RegisteredWaitHandle? showWindowWait;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            singleInstanceMutex = new Mutex(true, SingleInstanceMutexName, out var createdNew);
            if (!createdNew)
            {
                singleInstanceMutex.Dispose();
                singleInstanceMutex = null;
                if (e.Args.Length == 0 && TryShowRunningInstance())
                {
                    Shutdown(0);
                    return;
                }

                Console.WriteLine("MyBook is already running.");
                if (e.Args.Length == 0)
                    MessageBox.Show("MyBook is already running, but it cannot receive window activation requests.",
                        "MyBook", MessageBoxButton.OK, MessageBoxImage.Information);

                Shutdown(1);
                Environment.Exit(1);
                return;
            }

            // Create before constructing the window so activation requests during startup remain pending.
            if (e.Args.Length == 0)
                showWindowEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ShowWindowEventName);

            Func<string[], int>? runLogin = e.Args.Any(arg => arg.Equals("--bilibili-login", StringComparison.OrdinalIgnoreCase))
                ? WebUtil.BilibiliLogin.Run
                : e.Args.Any(arg => arg.Equals("--steam-login", StringComparison.OrdinalIgnoreCase)) ? SteamLogin.Run : null;
            if (runLogin is not null)
            {
                var exitCode = runLogin(e.Args);
                Shutdown(exitCode);
                Environment.Exit(exitCode);
                return;
            }

            if (e.Args.Any(arg => arg.Equals("--plaid-link", StringComparison.OrdinalIgnoreCase)))
            {
                var exitCode = 0;
                try
                {
                    var options = PlaidLink.ParseCommandLine(e.Args);
                    var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                    var database = new DatabaseUtil(config);
                    using var plaidLink = new PlaidLink(config, database);
                    var result = plaidLink.ConnectAndStoreAsync(options.CountryCodes, options.Products)
                        .GetAwaiter()
                        .GetResult();
                    Console.WriteLine($"Plaid Item connected and stored in database row {result.DatabaseId}.");
                }
                catch (Exception exception)
                {
                    exitCode = 1;
                    Console.WriteLine($"Plaid Link failed: {exception.Message}");
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

            MainWindow = new MainWindow();
            MainWindow.Show();
            if (showWindowEvent is not null)
                showWindowWait = ThreadPool.RegisterWaitForSingleObject(showWindowEvent, (_, _) =>
                {
                    if (!Dispatcher.HasShutdownStarted)
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (!Dispatcher.HasShutdownStarted && MainWindow is MyBook.MainWindow window)
                                window.RestoreFromTray();
                        }));
                }, null, Timeout.Infinite, false);
        }

        private static bool TryShowRunningInstance()
        {
            // The mutex owner may still be creating the event; wait briefly for that startup gap.
            for (var attempt = 0; attempt < 10; attempt++)
            {
                if (EventWaitHandle.TryOpenExisting(ShowWindowEventName, out var signal))
                {
                    using (signal)
                        return signal.Set();
                }
                Thread.Sleep(50);
            }
            return false;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            showWindowWait?.Unregister(null);
            showWindowEvent?.Dispose();
            try
            {
                singleInstanceMutex?.ReleaseMutex();
            }
            catch (ApplicationException)
            {
            }
            finally
            {
                singleInstanceMutex?.Dispose();
                singleInstanceMutex = null;
            }

            base.OnExit(e);
        }
    }
}
