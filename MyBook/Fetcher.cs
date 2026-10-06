using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MyBook
{
    partial class Fetcher : IDisposable
    {
        // RunSchedule initializes these before starting any scheduled work.
        IConfigurationRoot config = null!;
        MailUtil mail = null!;
        PubWebUtil pubWeb = null!;
        GraphQLUtil graphQL = null!;
        KrakenUtil kraken = null!;
        CryptoUtil crypto = null!;
        WebUtil web = null!;
        PlaidUtil plaid = null!;
        WiseUtil wise = null!;
        DatabaseUtil database = null!;
        SIMUtil sim = null!;
        Timer? dailyTimer;
        Timer? simTimer;
        Timer? marketPriceTimer;
        int marketPricesRunning;
        readonly ConcurrentDictionary<(string Code, HoldingType Type), MarketPrice> marketPrices = new();
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
            config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            marketPriceTimer?.Dispose();
            marketPriceTimer = new Timer(_ => RunMarketPricesInBackground(), null,
                TimeSpan.Zero, TimeSpan.FromMinutes(15));
            if (IsDebugBuild())
            {
                Console.WriteLine("skip scheduled imports in DEBUG");
                ResetRuntimeStatus();
                return;
            }

            database = new(config);
            mail = new(config, database);
            pubWeb = new(config, database);
            graphQL = new(config, database);
            web = new(config, database);
            wise = new(config, database);
            plaid = new(config, database);
            var krakenPub = new KrakenPubUtil();
            kraken = new(config, database, krakenPub);
            crypto = new(config, database, krakenPub);
            sim = new(database);
            dailyTimer?.Dispose();
            simTimer?.Dispose();
            UpdateRuntimeStatus(status => status.IsScheduledFetchEnabled = true);
            dailyTimer = new Timer(
                _ => RunDailyFetchInBackground(),
                null,
                TimeSpan.Zero,
                TimeSpan.FromDays(1));
            StartSIMPolling();
        }

        private void RunMarketPricesInBackground()
        {
            if (Interlocked.CompareExchange(ref marketPricesRunning, 1, 0) != 0)
                return;
            _ = Task.Run(async () =>
            {
                try
                {
                    // Only read the held symbols; live quotes never update accounting data.
                    var targets = new DatabaseUtil(config).GetMarketPriceTargets();
                    using var quotes = new PubWebUtil(config);
                    await quotes.FetchMarketPricesAsync(targets, new KrakenPubUtil(), CacheMarketPrice).ConfigureAwait(false);
                }
                catch (Exception e)
                {
                    CreateImportFailureMarker("market prices", e);
                    Console.WriteLine($"market prices failed: {e.Message}");
                }
                finally { Volatile.Write(ref marketPricesRunning, 0); }
            });
        }

        internal IReadOnlyList<MarketPrice> GetMarketPrices() => marketPrices.Values.ToArray();

        private void CacheMarketPrice(MarketPrice quote) => marketPrices[(quote.Code, quote.HoldingType)] = quote;

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
            UpdateRuntimeStatus(status => status.NextFetchTime = GetNextDailyRunTime(DateTime.Now));

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
            if (!await fetchLock.WaitAsync(0).ConfigureAwait(false))
                return;

            var started = false;
            try
            {
                var today = DateTime.Today;
                if (database.GetLatestStatementImportTime(StatementImportProvider.DailyFetch) >= today)
                    return;

                // Mark the attempt before importing so failures or restarts cannot repeat it today.
                database.SaveStatementProgress([StatementImportProvider.DailyFetch], today, "scheduled-attempt");
                started = true;
                SetCurrentTask("每日导入");
                await RunImportTaskAsync("Steam session refresh", web.RefreshSteamSessionAsync).ConfigureAwait(false);
                await RunImportTaskAsync("Bilibili session refresh", web.RefreshBilibiliSessionAsync).ConfigureAwait(false);
                await RunImportTaskAsync("exchange rate", pubWeb.FetchScheduledExchangeRates).ConfigureAwait(false);
                foreach (var source in PubWebUtil.KylcSources)
                    await RunImportTaskAsync(source.ToString(), () => pubWeb.FetchKylcRates(source)).ConfigureAwait(false);
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
                if (web.IsFirstTradeConfigured)
                    await RunScheduledImportTaskAsync("FirstTrade", StatementImportProvider.FirstTradeApi,
                        (_, _) => web.FetchFirstTradeAsync(),
                        intervalDays: 7, missingAfterDays: 0).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("Plaid Schwab", StatementImportProvider.PlaidSchwab,
                    (since, _) => plaid.FetchSchwabAsync(since), intervalDays: 1, missingAfterDays: 0).ConfigureAwait(false);
                if (wise.IsConfigured)
                    await RunScheduledImportTaskAsync("Wise API", StatementImportProvider.WiseApi,
                        (since, _) => wise.FetchAsync(since), intervalDays: 1, missingAfterDays: 0,
                        advanceOnEmptyQuery: true).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("Nexus DP", StatementImportProvider.NexusDpMonthlyReport,
                    (_, _) => graphQL.FetchNexusDpMonthlyReports(), intervalDays: 27, missingAfterDays: 40).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("Bilibili", StatementImportProvider.BilibiliWeb,
                    (_, _) => web.FetchBilibiliAsync(), intervalDays: 30, missingAfterDays: 35).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("Steam", StatementImportProvider.SteamWeb,
                    (since, _) => web.FetchSteamAsync(since), intervalDays: 3, missingAfterDays: 0,
                    advanceOnEmptyQuery: true).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("PayPal", CombinedUtil.PayPalProviders,
                    since => new CombinedUtil(database, plaid, mail).FetchPayPalAsync(since),
                    intervalDays: 1, missingAfterDays: 0, advanceOnEmptyQuery: true).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("Kraken", StatementImportProvider.KrakenApi,
                    (since, _) => kraken.FetchDailyReportsAsync(since), intervalDays: 1, missingAfterDays: 0).ConfigureAwait(false);
                await RunScheduledImportTaskAsync("Crypto ETH", StatementImportProvider.EthereumApi,
                    (since, _) => crypto.FetchDailyReportsAsync(since), intervalDays: 1, missingAfterDays: 0).ConfigureAwait(false);
                await RunImportTaskAsync(
                    "allocated expense cache",
                    () =>
                    {
                        database.ProcessAllocatedExpenseDirtyRecords();
                        return Task.CompletedTask;
                    }).ConfigureAwait(false);
                await RunImportTaskAsync(
                    "snapshot",
                    () =>
                    {
                        database.CreateDailySnapshot();
                        return Task.CompletedTask;
                    }).ConfigureAwait(false);
            }
            finally
            {
                if (started)
                    UpdateRuntimeStatus(status => status.LastFetchTime = DateTime.Now);
                SetCurrentTask(null);
                fetchLock.Release();
            }
        }

        private void StartSIMPolling()
        {
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
            if (Int32.TryParse(config["sim_poll_interval_minutes"], out var configuredMinutes)
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
            => RunScheduledImportTaskAsync(name, [provider], since => fetch(since[provider], missingAfterDays),
                intervalDays, missingAfterDays, advanceOnEmptyQuery);

        private Task RunScheduledImportTaskAsync(string name, IReadOnlyList<StatementImportProvider> providers,
            Func<IReadOnlyDictionary<StatementImportProvider, DateTime>, Task> fetch,
            int intervalDays, int missingAfterDays = 0, bool advanceOnEmptyQuery = false)
        {
            return RunImportTaskAsync(name, () =>
            {
                var progress = new Dictionary<StatementImportProvider, DateTime>();
                return ImportSchedule.RunAsync(name, intervalDays, missingAfterDays,
                    () =>
                    {
                        progress = providers.ToDictionary(provider => provider, provider => (advanceOnEmptyQuery
                            ? database.GetLatestStatementImportTimeByKeyPrefix(provider, ImportSchedule.SuccessfulQueryKey)
                            ?? database.GetStatementImportCheckpointTime(provider)
                            : database.GetLatestStatementImportTime(provider))?.Date
                            ?? throw new InvalidOperationException($"Missing {provider} import checkpoint."));
                        return progress.Values.Min();
                    },
                    _ => fetch(progress),
                    !advanceOnEmptyQuery ? null : date => database.SaveStatementQueryProgress(providers, date));
            });
        }

        private async Task RunImportTaskAsync(string name, Func<Task> fetch)
        {
            SetCurrentTask(name);
            try
            {
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

        private static DateTime GetNextDailyRunTime(DateTime now)
            => now.AddDays(1);

        public void Dispose()
        {
            lock (runtimeStatusLock)
            {
                dailyTimer?.Dispose();
                dailyTimer = null;
                runtimeStatus.IsScheduledFetchEnabled = false;
                runtimeStatus.NextFetchTime = null;
            }
            simTimer?.Dispose();
            marketPriceTimer?.Dispose();
            pubWeb?.Dispose();
            fetchLock.Dispose();
            simPollLock.Dispose();
        }
    }

    internal sealed record MarketPrice(string Code, HoldingType HoldingType, decimal Price,
        CurrencyType Currency, DateTimeOffset FetchedAt);

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
