using Microsoft.Extensions.Configuration;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MyBook
{
    partial class Fetcher : IDisposable
    {
        IConfigurationRoot? config;
        MailUtil? mail;
        PubWebUtil? pubWeb;
        GraphQLUtil? graphQL;
        KrakenUtil? kraken;
        CryptoUtil? crypto;
        WebUtil? web;
        PlaidUtil? plaid;
        WiseUtil? wise;
        DatabaseUtil? database;
        SIMUtil? sim;
        Timer? dailyTimer;
        Timer? simTimer;
        readonly SemaphoreSlim fetchLock = new(1, 1);
        readonly SemaphoreSlim simPollLock = new(1, 1);
        readonly object runtimeStatusLock = new();
        FetchRuntimeStatus runtimeStatus = new();
        const int DefaultSIMPollIntervalMinutes = 5;
        const string ImportFailureMarkerFileName = "MyBook.import-failed.tmp";
        static readonly UTF8Encoding ImportFailureMarkerEncoding = new(false);
        static readonly object importFailureMarkerLock = new();

        public void RunSchedule()
        {
            if (IsDebugBuild())
            {
                Console.WriteLine("skip scheduled fetch in DEBUG");
                ResetRuntimeStatus();
                return;
            }

            config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            database = new(config);
            mail = new(config, database);
            pubWeb = new(config, database);
            graphQL = new(config, database);
            web = new(config, database);
            wise = new(config, database);
            if (!String.IsNullOrWhiteSpace(config["plaid_client_id"])
                && !String.IsNullOrWhiteSpace(config[PlaidUtil.SelectedSecretConfigKey]))
                plaid = new(config, database);
            var krakenPub = new KrakenPubUtil();
            var krakenApiKey = config["kraken_api_key"];
            var krakenApiSecret = config["kraken_api_secret"];
            if (String.IsNullOrWhiteSpace(krakenApiKey) && String.IsNullOrWhiteSpace(krakenApiSecret))
            {
                Console.WriteLine("skip scheduled Kraken fetch: missing Kraken API credentials");
                kraken = null;
            }
            else if (String.IsNullOrWhiteSpace(krakenApiKey) || String.IsNullOrWhiteSpace(krakenApiSecret))
            {
                Console.WriteLine("skip scheduled Kraken fetch: incomplete Kraken API credentials");
                kraken = null;
            }
            else
            {
                kraken = new(config, database, krakenPub);
            }
            crypto = new(config, database, krakenPub);
            sim = new(database);
            dailyTimer?.Dispose();
            simTimer?.Dispose();
            var nextDailyRun = GetNextDailyRunTime();
            UpdateRuntimeStatus(status =>
            {
                status.IsScheduledFetchEnabled = true;
                status.NextFetchTime = nextDailyRun;
            });
            RunDailyFetchInBackground();
            dailyTimer = new Timer(
                _ => RunDailyFetchInBackground(),
                null,
                GetDueTime(nextDailyRun),
                TimeSpan.FromDays(1));
            StartSIMPolling();
            //pubWeb.Fetch(new Finance("QQQ", HoldingType.NASDAQ));
            //pubWeb.Fetch(new Finance("021282", HoldingType.CNFUND));
        }

        private static bool IsDebugBuild()
        {
#if DEBUG
            return true;
#else
            return false;
#endif
        }

        private void RunDailyFetchInBackground()
        {
            UpdateRuntimeStatus(status =>
            {
                if (status.IsScheduledFetchEnabled)
                    status.NextFetchTime = GetNextDailyRunTime();
            });

            _ = Task.Run(async () =>
            {
                try
                {
                    await RunDailyFetchAsync().ConfigureAwait(false);
                }
                catch (Exception e)
                {
                    CreateImportFailureMarker("scheduled fetch", e);
                    Console.WriteLine($"scheduled fetch fail: {e.Message}");
                }
            });
        }

        private async Task RunDailyFetchAsync()
        {
            if (mail is null)
                return;
            if (!await fetchLock.WaitAsync(0).ConfigureAwait(false))
                return;

            SetCurrentTask("每日导入");
            try
            {
                await mail.RunWithMailSessionScope(async () =>
                {
                    await RunScheduledImportTaskAsync("ICBC", StatementImportProvider.ICBCBillMail,
                        (since, limit) => mail.FetchICBCBills(since, limit), intervalDays: 27, missingAfterDays: 40).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("BOC", StatementImportProvider.BOCBillMail,
                        (since, limit) => mail.FetchBOCBills(since, limit), intervalDays: 27, missingAfterDays: 40).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("ICBC history detail", StatementImportProvider.ICBCHistoryDetailMail,
                        (since, _) => mail.FetchICBCHistoryDetails(since),
                        intervalDays: 90, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("IBKR", StatementImportProvider.IBKRReportMail,
                        (since, limit) => mail.FetchIBKRReports(since, limit), intervalDays: 1, missingAfterDays: 5).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("iFAST", StatementImportProvider.IFastMail,
                        (since, _) => mail.FetchIFastMessages(since),
                        intervalDays: 1, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("ZA", StatementImportProvider.ZAMail,
                        (since, _) => mail.FetchZAMessages(since),
                        intervalDays: 1, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("Ant", StatementImportProvider.AntMail,
                        (since, _) => mail.FetchAntMessages(since),
                        intervalDays: 1, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                    await RunScheduledImportTaskAsync("Ele", StatementImportProvider.EleMail,
                        (since, _) => mail.FetchEleMessages(since),
                        intervalDays: 1, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                }).ConfigureAwait(false);
                // FirstTrade 定时导入暂时停用；恢复时启用以下调用。
                // if (web is not null && web.IsFirstTradeConfigured)
                //     await RunScheduledImportTaskAsync("FirstTrade", StatementImportProvider.FirstTradeApi,
                //         (_, _) => web.FetchFirstTradeAsync(),
                //         intervalDays: 7, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                if (plaid is not null)
                {
                    await RunImportTaskAsync("Plaid Schwab", () => true, () => plaid.FetchSchwabAsync()).ConfigureAwait(false);
                }
                if (wise is not null && wise.IsConfigured)
                    await RunImportTaskAsync("Wise API", () => true, () => wise.FetchAsync()).ConfigureAwait(false);
                if (graphQL is not null)
                    await RunScheduledImportTaskAsync("Nexus DP", StatementImportProvider.NexusDpMonthlyReport,
                        (_, _) => graphQL.FetchNexusDpMonthlyReports(), intervalDays: 27, missingAfterDays: 40).ConfigureAwait(false);
                if (plaid is not null && database is not null)
                    await RunImportTaskAsync("PayPal", () => true,
                        () => new CombinedUtil(database, plaid, mail).FetchPayPalAsync()).ConfigureAwait(false);
                if (kraken is not null)
                    await RunImportTaskAsync(
                        "Kraken",
                        () => ShouldFetchProviderAfterDays("Kraken", StatementImportProvider.KrakenApi, 1),
                        () => kraken.FetchDailyReportsAsync()).ConfigureAwait(false);
                if (crypto is not null)
                    await RunImportTaskAsync(
                        "Crypto ETH",
                        () => ShouldFetchProviderAfterDays("Crypto ETH", StatementImportProvider.EthereumApi, 1),
                        () => crypto.FetchDailyReportsAsync()).ConfigureAwait(false);
                if (pubWeb is not null)
                    await RunImportTaskAsync("exchange rate", () => true, pubWeb.FetchExchangeRates).ConfigureAwait(false);
                if (database is not null)
                {
                    await RunImportTaskAsync(
                        "allocated expense cache",
                        () => true,
                        () =>
                        {
                            database.ProcessAllocatedExpenseDirtyRecords();
                            return Task.CompletedTask;
                        }).ConfigureAwait(false);
                    await RunImportTaskAsync(
                        "snapshot",
                        () => true,
                        () =>
                        {
                            database.CreateDailySnapshot();
                            return Task.CompletedTask;
                        }).ConfigureAwait(false);
                }
            }
            finally
            {
                UpdateRuntimeStatus(status => status.LastFetchTime = DateTime.Now);
                SetCurrentTask(null);
                fetchLock.Release();
            }
        }

        private void StartSIMPolling()
        {
            if (config is null)
                return;

            if (String.IsNullOrWhiteSpace(config["sim_imsi"]))
            {
                Console.WriteLine("skip scheduled SIM SMS polling: missing sim_imsi in config.json");
                return;
            }

            var interval = GetSIMPollInterval();
            Console.WriteLine($"scheduled SIM SMS polling every {interval.TotalMinutes:0} minute(s)");
            RunSIMPollInBackground();
            simTimer = new Timer(
                _ => RunSIMPollInBackground(),
                null,
                interval,
                interval);
        }

        private TimeSpan GetSIMPollInterval()
        {
            if (config is not null
                && Int32.TryParse(config["sim_poll_interval_minutes"], out var configuredMinutes)
                && configuredMinutes > 0)
                return TimeSpan.FromMinutes(Math.Max(1, configuredMinutes));

            return TimeSpan.FromMinutes(DefaultSIMPollIntervalMinutes);
        }

        private void RunSIMPollInBackground()
        {
            _ = Task.Run(RunSIMPollAsync);
        }

        private async Task RunSIMPollAsync()
        {
            if (config is null || sim is null)
                return;

            if (!await simPollLock.WaitAsync(0).ConfigureAwait(false))
            {
                Console.WriteLine("skip scheduled SIM SMS polling: previous poll is still running");
                return;
            }

            try
            {
                var expectedImsi = config["sim_imsi"];
                if (String.IsNullOrWhiteSpace(expectedImsi))
                {
                    Console.WriteLine("skip scheduled SIM SMS polling: missing sim_imsi in config.json");
                    return;
                }

                var result = await sim.PollConfiguredSIMMessages(expectedImsi).ConfigureAwait(false);
                foreach (var line in result.LogLines)
                    Console.WriteLine(line);
            }
            catch (Exception e)
            {
                CreateImportFailureMarker("SIM", e);
                Console.WriteLine($"scheduled SIM SMS polling fail: {e.Message}");
            }
            finally
            {
                simPollLock.Release();
            }
        }

        private Task RunScheduledImportTaskAsync(string name, StatementImportProvider provider,
            Func<DateTime, int, Task> fetch, int intervalDays, int missingAfterDays = 0, bool advanceOnEmptyQuery = false)
        {
            return RunImportTaskAsync(name, () => true, () =>
            {
                var db = database ?? throw new InvalidOperationException("Scheduled import requires a database.");
                return ImportSchedule.RunAsync(name, intervalDays, missingAfterDays,
                    () => advanceOnEmptyQuery
                        ? db.GetLatestStatementImportTimeByKeyPrefix(provider, ImportSchedule.SuccessfulQueryKey)
                            ?? db.GetStatementImportCheckpointTime(provider)
                        : db.GetLatestStatementImportTime(provider),
                    since => fetch(since, missingAfterDays),
                    !advanceOnEmptyQuery ? null : date => db.SaveStatementQueryProgress(provider, date));
            });
        }

        private bool ShouldFetchProviderAfterDays(string name, StatementImportProvider provider, int intervalDays)
        {
            if (intervalDays <= 0)
                throw new ArgumentOutOfRangeException(nameof(intervalDays));
            if (database is null)
                return true;

            var latestImportTime = database.GetLatestStatementImportTime(provider);
            if (latestImportTime is null)
                return true;

            var elapsedDays = (DateTime.Today - latestImportTime.Value.Date).TotalDays;
            if (elapsedDays >= intervalDays)
                return true;

            Console.WriteLine($"skip scheduled {name} fetch: last import {latestImportTime.Value:yyyy-MM-dd}, elapsed {elapsedDays:0} days");
            return false;
        }

        private async Task RunImportTaskAsync(string name, Func<bool> shouldRun, Func<Task> fetch)
        {
            SetCurrentTask(name);
            try
            {
                if (shouldRun())
                    await fetch().ConfigureAwait(false);
            }
            catch (Exception e)
            {
                CreateImportFailureMarker(name, e);
                Console.WriteLine($"scheduled {name} fetch fail: {e.Message}");
            }
            finally
            {
                SetCurrentTask("每日导入");
            }
        }

        public FetchRuntimeStatus GetRuntimeStatus()
        {
            lock (runtimeStatusLock)
            {
                var status = runtimeStatus.Clone();
                status.HasImportFailureMarker = File.Exists(GetImportFailureMarkerPath());
                return status;
            }
        }

        private static string GetImportFailureMarkerPath()
        {
            return Path.Combine(AppContext.BaseDirectory, ImportFailureMarkerFileName);
        }

        public void ClearImportFailureMarker()
        {
            lock (importFailureMarkerLock)
                File.Delete(GetImportFailureMarkerPath());
        }

        private static void CreateImportFailureMarker(string taskName, Exception exception)
        {
            OnImportFailed(taskName, exception);
            try
            {
                var content = String.Join(
                    Environment.NewLine,
                    "MyBook import failed",
                    DateTime.Now.ToString("O", CultureInfo.InvariantCulture),
                    taskName,
                    exception.GetType().FullName ?? exception.GetType().Name)
                    + Environment.NewLine;
                lock (importFailureMarkerLock)
                    File.WriteAllText(GetImportFailureMarkerPath(), content, ImportFailureMarkerEncoding);
            }
            catch (Exception markerException)
            {
                Console.WriteLine($"write import failure marker fail: {markerException.Message}");
            }
        }

        static partial void OnImportFailed(string taskName, Exception exception);

        private void UpdateRuntimeStatus(Action<FetchRuntimeStatus> update)
        {
            lock (runtimeStatusLock)
            {
                update(runtimeStatus);
            }
        }

        private void SetCurrentTask(string? name)
        {
            UpdateRuntimeStatus(status =>
            {
                status.CurrentTaskName = name;
                status.CurrentTaskStartedAt = name is null ? null : DateTime.Now;
            });
        }

        private void ResetRuntimeStatus()
        {
            lock (runtimeStatusLock)
            {
                runtimeStatus = new FetchRuntimeStatus();
            }
        }

        private static DateTime GetNextDailyRunTime()
        {
            var now = DateTime.Now;
            var nextRun = DateTime.Today.AddDays(1).AddMinutes(5);
            if (now >= nextRun)
                nextRun = nextRun.AddDays(1);
            return nextRun;
        }

        private static TimeSpan GetDueTime(DateTime runTime)
        {
            var dueTime = runTime - DateTime.Now;
            return dueTime > TimeSpan.Zero ? dueTime : TimeSpan.Zero;
        }

        public void Dispose()
        {
            dailyTimer?.Dispose();
            simTimer?.Dispose();
            pubWeb?.Dispose();
            fetchLock.Dispose();
            simPollLock.Dispose();
        }
    }

    public class FetchRuntimeStatus
    {
        public bool IsScheduledFetchEnabled { get; set; }
        public string? CurrentTaskName { get; set; }
        public DateTime? CurrentTaskStartedAt { get; set; }
        public DateTime? LastFetchTime { get; set; }
        public DateTime? NextFetchTime { get; set; }
        public bool HasImportFailureMarker { get; set; }

        public FetchRuntimeStatus Clone()
        {
            return (FetchRuntimeStatus)MemberwiseClone();
        }
    }
}
