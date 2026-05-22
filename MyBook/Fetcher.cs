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
        DatabaseUtil? database;
        Timer? dailyTimer;
        readonly SemaphoreSlim fetchLock = new(1, 1);

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
                await TryFetchAsync("ICBC", mail.FetchICBCBills);
                await TryFetchAsync("IBKR", mail.FetchIBKRReports);
                await TryFetchAsync("Wise", mail.FetchWiseReports);
                if (stock is not null)
                    await TryFetchAsync("exchange rate", stock.FetchExchangeRates);
            }
            finally
            {
                fetchLock.Release();
            }
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
            var nextRun = DateTime.Today.AddDays(1).AddHours(3);
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
