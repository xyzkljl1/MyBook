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
        StockUtil? stock;
        GraphQLUtil? graphQL;
        DatabaseUtil? database;
        Timer? dailyTimer;
        readonly SemaphoreSlim fetchLock = new(1, 1);
        const int MonthlyFetchIntervalDays = 27;
        const int ICBCHistoryDetailFetchIntervalDays = 90;
        const int ICBCHistoryDetailSearchWindowMonths = 5;

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
            stock = new(config, database);
            graphQL = new(config, database);
            dailyTimer?.Dispose();
            _ = RunDailyFetchAsync();
            dailyTimer = new Timer(
                _ => _ = RunDailyFetchAsync(),
                null,
                GetDelayUntilNextDailyRun(),
                TimeSpan.FromDays(1));
            //stock.Fetch(new Finance("QQQ", HoldingType.NASDAQ));
            //stock.Fetch(new Finance("021282", HoldingType.CNFUND));
        }

        private static bool IsDebugBuild()
        {
#if DEBUG
            return true;
#else
            return false;
#endif
        }

        private async Task RunDailyFetchAsync()
        {
            if (mail is null)
                return;
            if (!await fetchLock.WaitAsync(0))
                return;

            try
            {
                if (ShouldFetchMonthlyProvider("ICBC", StatementImportProvider.ICBCBillMail))
                    await TryFetchAsync("ICBC", mail.FetchICBCBills);
                if (ShouldFetchProviderAfterDays("ICBC history detail", StatementImportProvider.ICBCHistoryDetailMail, ICBCHistoryDetailFetchIntervalDays))
                    await TryFetchAsync("ICBC history detail", FetchICBCHistoryDetailsScheduledAsync);
                await TryFetchAsync("IBKR", mail.FetchIBKRReports);
                if (ShouldFetchMonthlyProvider("Wise", StatementImportProvider.WiseMail))
                    await TryFetchAsync("Wise", mail.FetchWiseReports);
                if (ShouldFetchMonthlyProvider("OCBC", StatementImportProvider.OCBCStatementMail))
                    await TryFetchAsync("OCBC", mail.FetchOCBCReports);
                if (graphQL is not null && ShouldFetchMonthlyProvider("Nexus DP", StatementImportProvider.NexusDpMonthlyReport))
                    await TryFetchAsync("Nexus DP", graphQL.FetchNexusDpMonthlyReports);
                if (stock is not null)
                    await TryFetchAsync("exchange rate", stock.FetchExchangeRates);
                if (database is not null)
                {
                    await TryFetchAsync("snapshot", () =>
                    {
                        database.CreateDailySnapshot();
                        return Task.CompletedTask;
                    });
                }
            }
            finally
            {
                fetchLock.Release();
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
            await mail.FetchICBCHistoryDetails(today.AddMonths(-ICBCHistoryDetailSearchWindowMonths));
            database.MarkStatementProcessedOnce(
                StatementImportProvider.ICBCHistoryDetailMail,
                today,
                $"scheduled-empty-import-{today:yyyyMMdd}");
        }

        private static async Task TryFetchAsync(string name, Func<Task> fetch)
        {
            try
            {
                await fetch();
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
            fetchLock.Dispose();
        }
    }
}
