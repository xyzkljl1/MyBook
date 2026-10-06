using System.Text.Json;

namespace MyBook
{
    partial class WebUtil
    {
        public async Task FetchBilibiliAsync()
        {
            var account = database.GetAccountByName("Bilibili");
            var beginning = database.GetAccountBalance(account, CurrencyType.RMB);
            var ending = await FetchBilibiliBalance().ConfigureAwait(false);
            var now = DateTime.Now;
            var increase = ending.v - beginning.v;
            if (increase < 0)
                throw new InvalidOperationException($"Bilibili balance decreased: {beginning.v} -> {ending.v} RMB.");

            // User-approved balance-difference income; the observation time is not the video earning date.
            List<Record> records = increase == 0 ? [] : [new Record
            {
                Account = account, v = increase, t = CurrencyType.RMB,
                date = now, postingDate = now, updateTime = now, Reason = "视频收益",
                Source = FormattableString.Invariant($"Bilibili getUserBrokerage; beginning={beginning.v}; ending={ending.v}; observedAt={now:O}")
            }];
            // An unchanged balance is still a successful observation and advances the query date.
            database.SaveStatementRecordsOnce(StatementImportProvider.BilibiliWeb, now.Date, records,
                accountBalances: [new(account, ending)],
                statementKey: $"Bilibili/{account.Id}/{now:yyyyMMddHHmmssfffffff}",
                beginningAccountBalances: [new(account, beginning)], forceValidateBeginningBalances: true);
        }

        // The module manages one Bilibili login independently of the browser and financial imports.
        public async Task LoginBilibiliAsync(Func<string, Task> showQrCode, Action<string> showStatus,
            CancellationToken cancellationToken = default)
        {
            _ = database.GetAccountByName("Bilibili");
            using var sessionStore = database.OpenLoginSession(LoginProvider.Bilibili);
            using var client = new BilibiliClient();
            var session = await client.LoginAsync(showQrCode, showStatus, cancellationToken).ConfigureAwait(false);
            sessionStore.Save(JsonSerializer.Serialize(session), newSession: true);
        }

        public async Task<Currency> FetchBilibiliBalance()
        {
            using var sessionStore = database.OpenLoginSession(LoginProvider.Bilibili);
            var json = sessionStore.Read();
            BilibiliClient.Session session;
            try
            {
                session = json is null ? throw new JsonException() :
                    JsonSerializer.Deserialize<BilibiliClient.Session>(json) ?? throw new JsonException();
                if (String.IsNullOrWhiteSpace(session.Cookie)) throw new JsonException();
            }
            catch (JsonException)
            {
                throw new BilibiliException("Bilibili session unavailable; run MyBook.exe --bilibili-login to log in.");
            }
            using var client = new BilibiliClient();
            await client.RestoreAndRefreshAsync(session,
                (value, newSession) => sessionStore.Save(JsonSerializer.Serialize(value), newSession)).ConfigureAwait(false);
            return await client.FetchBalanceAsync().ConfigureAwait(false);
        }
    }
}
