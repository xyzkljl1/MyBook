using Microsoft.Extensions.Configuration;
using System;
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
        DatabaseUtil? database;
        SIMUtil? sim;
        Timer? dailyTimer;
        Timer? simTimer;
        readonly SemaphoreSlim fetchLock = new(1, 1);
        readonly SemaphoreSlim simPollLock = new(1, 1);
        const int MonthlyFetchIntervalDays = 27;
        const int ICBCHistoryDetailFetchIntervalDays = 90;
        const int ICBCHistoryDetailSearchWindowMonths = 5;
        const int DefaultSIMPollIntervalMinutes = 5;

        public void RunSchedule()
        {
            if (IsDebugBuild())
            {
                Console.WriteLine("skip scheduled fetch in DEBUG");
                return;
            }

            config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            database = new(config);
            mail = new(config, database);
            pubWeb = new(config, database);
            graphQL = new(config, database);
            sim = new(database);
            dailyTimer?.Dispose();
            simTimer?.Dispose();
            RunDailyFetchInBackground();
            dailyTimer = new Timer(
                _ => RunDailyFetchInBackground(),
                null,
                GetDelayUntilNextDailyRun(),
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
            _ = Task.Run(async () =>
            {
                try
                {
                    await RunDailyFetchAsync().ConfigureAwait(false);
                }
                catch (Exception e)
                {
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

            try
            {
                await mail.RunWithMailSessionScope(async () =>
                {
                    if (ShouldFetchMonthlyProvider("ICBC", StatementImportProvider.ICBCBillMail))
                        await TryFetchAsync("ICBC", mail.FetchICBCBills).ConfigureAwait(false);
                    if (ShouldFetchProviderAfterDays("ICBC history detail", StatementImportProvider.ICBCHistoryDetailMail, ICBCHistoryDetailFetchIntervalDays))
                        await TryFetchAsync("ICBC history detail", FetchICBCHistoryDetailsScheduledAsync).ConfigureAwait(false);
                    await TryFetchAsync("IBKR", mail.FetchIBKRReports).ConfigureAwait(false);
                    if (ShouldFetchMonthlyProvider("Wise", StatementImportProvider.WiseMail))
                        await TryFetchAsync("Wise", mail.FetchWiseReports).ConfigureAwait(false);
                    if (ShouldFetchMonthlyProvider("OCBC", StatementImportProvider.OCBCStatementMail))
                        await TryFetchAsync("OCBC", mail.FetchOCBCReports).ConfigureAwait(false);
                }).ConfigureAwait(false);
                if (graphQL is not null && ShouldFetchMonthlyProvider("Nexus DP", StatementImportProvider.NexusDpMonthlyReport))
                    await TryFetchAsync("Nexus DP", graphQL.FetchNexusDpMonthlyReports).ConfigureAwait(false);
                if (pubWeb is not null)
                    await TryFetchAsync("exchange rate", pubWeb.FetchExchangeRates).ConfigureAwait(false);
                if (database is not null)
                {
                    await TryFetchAsync("allocated expense cache", () =>
                    {
                        database.ProcessAllocatedExpenseDirtyRecords();
                        return Task.CompletedTask;
                    }).ConfigureAwait(false);
                    await TryFetchAsync("snapshot", () =>
                    {
                        database.CreateDailySnapshot();
                        return Task.CompletedTask;
                    }).ConfigureAwait(false);
                }
            }
            finally
            {
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

        private static async Task TryFetchAsync(string name, Func<Task> fetch)
        {
            try
            {
                await fetch().ConfigureAwait(false);
            }
            catch (Exception e)
            {
                Console.WriteLine($"scheduled {name} fetch fail: {e.Message}");
            }
        }

        private static TimeSpan GetDelayUntilNextDailyRun()
        {
            var now = DateTime.Now;
            var nextRun = DateTime.Today.AddDays(1).AddMinutes(5);
            if (now >= nextRun)
                nextRun = nextRun.AddDays(1);
            return nextRun - now;
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
}
