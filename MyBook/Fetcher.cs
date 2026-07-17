using Microsoft.Extensions.Configuration;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MyBook
{
    class Fetcher : IDisposable
    {
        IConfigurationRoot? config;
        MailUtil? mail;
        PubWebUtil? pubWeb;
        GraphQLUtil? graphQL;
        KrakenUtil? kraken;
        CryptoUtil? crypto;
        DatabaseUtil? database;
        SIMUtil? sim;
        Timer? dailyTimer;
        Timer? simTimer;
        readonly SemaphoreSlim fetchLock = new(1, 1);
        readonly SemaphoreSlim simPollLock = new(1, 1);
        readonly object runtimeStatusLock = new();
        FetchRuntimeStatus runtimeStatus = new();
        const int MonthlyFetchIntervalDays = 27;
        const int ICBCHistoryDetailFetchIntervalDays = 90;
        const int ICBCHistoryDetailSearchWindowMonths = 5;
        const int DefaultSIMPollIntervalMinutes = 5;
        const string ImportFailureMarkerFileName = "MyBook.import-failed.tmp";
        static readonly UTF8Encoding ImportFailureMarkerEncoding = new(false);

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
                kraken = new(config, database);
            }
            crypto = new(config, database);
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
                    await RunImportTaskAsync(
                        "ICBC",
                        () => ShouldFetchMonthlyProvider("ICBC", StatementImportProvider.ICBCBillMail),
                        mail.FetchICBCBills).ConfigureAwait(false);
                    await RunImportTaskAsync(
                        "ICBC history detail",
                        () => ShouldFetchProviderAfterDays("ICBC history detail", StatementImportProvider.ICBCHistoryDetailMail, ICBCHistoryDetailFetchIntervalDays),
                        FetchICBCHistoryDetailsScheduledAsync).ConfigureAwait(false);
                    await RunImportTaskAsync("IBKR", () => true, mail.FetchIBKRReports).ConfigureAwait(false);
                    await RunImportTaskAsync(
                        "OCBC",
                        () => ShouldFetchMonthlyProvider("OCBC", StatementImportProvider.OCBCStatementMail),
                        mail.FetchOCBCReports).ConfigureAwait(false);
                }).ConfigureAwait(false);
                if (graphQL is not null)
                    await RunImportTaskAsync(
                        "Nexus DP",
                        () => ShouldFetchMonthlyProvider("Nexus DP", StatementImportProvider.NexusDpMonthlyReport),
                        graphQL.FetchNexusDpMonthlyReports).ConfigureAwait(false);
                if (kraken is not null)
                    await RunImportTaskAsync(
                        "Kraken",
                        () => ShouldFetchProviderAfterDays("Kraken", StatementImportProvider.KrakenApi, 0),
                        () => kraken.FetchDailyReportsAsync()).ConfigureAwait(false);
                if (crypto is not null)
                    await RunImportTaskAsync(
                        "Crypto ETH",
                        () => ShouldFetchProviderAfterDays("Crypto ETH", StatementImportProvider.EthereumApi, 0),
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
                Console.WriteLine($"scheduled SIM SMS polling fail: {e.Message}");
            }
            finally
            {
                simPollLock.Release();
            }
        }

        private bool ShouldFetchMonthlyProvider(string name, StatementImportProvider provider)
        {
            return ShouldFetchProviderAfterDays(name, provider, MonthlyFetchIntervalDays);
        }

        private bool ShouldFetchProviderAfterDays(string name, StatementImportProvider provider, int intervalDays)
        {
            if (database is null)
                return true;

            var latestImportTime = database.GetLatestStatementImportTime(provider);
            if (latestImportTime is null)
                return true;

            var elapsedDays = (DateTime.Today - latestImportTime.Value.Date).TotalDays;
            if (elapsedDays > intervalDays)
                return true;

            Console.WriteLine($"skip scheduled {name} fetch: last import {latestImportTime.Value:yyyy-MM-dd}, elapsed {elapsedDays:0} days");
            return false;
        }

        private async Task FetchICBCHistoryDetailsScheduledAsync()
        {
            if (mail is null || database is null)
                return;

            var today = DateTime.Today;
            await mail.FetchICBCHistoryDetails(today.AddMonths(-ICBCHistoryDetailSearchWindowMonths)).ConfigureAwait(false);
            database.MarkStatementProcessedOnce(
                StatementImportProvider.ICBCHistoryDetailMail,
                today,
                $"scheduled-empty-import-{today:yyyyMMdd}");
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
            return Path.Combine(Path.GetTempPath(), ImportFailureMarkerFileName);
        }

        private static void CreateImportFailureMarker(string taskName, Exception exception)
        {
            try
            {
                var content = String.Join(
                    Environment.NewLine,
                    "MyBook import failed",
                    DateTime.Now.ToString("O", CultureInfo.InvariantCulture),
                    taskName,
                    exception.GetType().FullName ?? exception.GetType().Name)
                    + Environment.NewLine;
                File.WriteAllText(GetImportFailureMarkerPath(), content, ImportFailureMarkerEncoding);
            }
            catch (Exception markerException)
            {
                Console.WriteLine($"write import failure marker fail: {markerException.Message}");
            }
        }

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
