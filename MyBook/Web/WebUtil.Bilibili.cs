using System.Net.Http;
using System.Net.Http.Json;
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

        public async Task<Currency> FetchBilibiliBalance()
        {
            const string requestName = "POST pay.bilibili.com/bk/brokerage/getUserBrokerage";
            var cookie = config["bilibili_cookie"];
            if (String.IsNullOrWhiteSpace(cookie))
                throw new InvalidOperationException("Missing bilibili_cookie configuration.");
            try
            {
                using var client = new HttpClient(new HttpClientHandler
                {
                    UseProxy = false, UseCookies = false, AllowAutoRedirect = false
                }) { Timeout = TimeSpan.FromSeconds(30) };
                using var request = new HttpRequestMessage(HttpMethod.Post,
                    "https://pay.bilibili.com/bk/brokerage/getUserBrokerage");
                request.Headers.Add("Cookie", cookie);
                request.Headers.Referrer = new Uri("https://pay.bilibili.com/pay-v2-web/shell_index");
                request.Headers.Add("Origin", "https://pay.bilibili.com");
                var timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                request.Content = JsonContent.Create(new { traceId = timestamp, timestamp, sdkVersion = "1.2.1" });
                using var response = await client.SendAsync(request).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException($"{requestName}: HTTP {(int)response.StatusCode}.");
                using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                var root = json.RootElement;
                var code = root.GetProperty("errno").GetInt64();
                if (code != 0)
                    throw new InvalidOperationException($"{requestName}: HTTP {(int)response.StatusCode}, business code {code}.");
                var balance = root.GetProperty("data").GetProperty("brokerage").GetDecimal();
                if (balance < 0 || Decimal.Round(balance, 2) != balance)
                    throw new InvalidOperationException($"{requestName}: invalid balance.");
                return new Currency(balance, CurrencyType.RMB);
            }
            catch (HttpRequestException e)
            {
                throw new InvalidOperationException($"{requestName}: {e.HttpRequestError}.");
            }
            catch (TaskCanceledException)
            {
                throw new InvalidOperationException($"{requestName}: timeout.");
            }
            catch (Exception e) when (e is JsonException or KeyNotFoundException or FormatException)
            {
                throw new InvalidOperationException($"{requestName}: invalid request or response format.");
            }
        }
    }
}
