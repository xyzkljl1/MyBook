using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using HtmlAgilityPack;

namespace MyBook;

partial class WebUtil
{
    // An independent web session, obtained by scanning the program's own QR code.
    internal sealed class BilibiliClient : IDisposable
    {
        private const string Passport = "https://passport.bilibili.com/x/passport-login/web/";
        private static readonly Uri CookieOrigin = new("https://www.bilibili.com/");
        private readonly CookieContainer cookies = new();
        private readonly HttpClient client;

        internal BilibiliClient(HttpMessageHandler? handler = null)
        {
            client = new HttpClient(handler ?? new HttpClientHandler
            {
                UseProxy = false, UseCookies = false, AllowAutoRedirect = false,
                AutomaticDecompression = DecompressionMethods.All
            }) { Timeout = TimeSpan.FromSeconds(30) };
        }

        internal sealed class Session
        {
            public required string Cookie { get; set; }
            public required string RefreshToken { get; set; }
            public string? PendingConfirmationToken { get; set; }
        }

        internal async Task<Session> LoginAsync(Func<string, Task> showQrCode, Action<string> showStatus,
            CancellationToken cancellationToken)
        {
            using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            deadline.CancelAfter(TimeSpan.FromMinutes(3));
            var token = deadline.Token;
            try
            {
                var data = await JsonAsync(HttpMethod.Get, Passport + "qrcode/generate", token).ConfigureAwait(false);
                var key = RequiredString(data, "qrcode_key", "qrcode/generate");
                await showQrCode(RequiredString(data, "url", "qrcode/generate")).ConfigureAwait(false);
                while (true)
                {
                    data = await JsonAsync(HttpMethod.Get, Passport + "qrcode/poll?qrcode_key=" + Uri.EscapeDataString(key), token).ConfigureAwait(false);
                    var code = RequiredInt(data, "code", "qrcode/poll");
                    switch (code)
                    {
                        case 0:
                            return CaptureSession(RequiredString(data, "refresh_token", "qrcode/poll"));
                        case 86101: showStatus("请使用哔哩哔哩 App 扫码"); break;
                        case 86090: showStatus("已扫码，请在手机上确认登录"); break;
                        case 86038: throw new BilibiliException("Bilibili QR code expired; open login again.");
                        default: throw new BilibiliException($"GET passport.bilibili.com/x/passport-login/web/qrcode/poll: business code {code}.");
                    }
                    await Task.Delay(TimeSpan.FromSeconds(2), token).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                throw new BilibiliException("Bilibili QR login timed out; open login again.");
            }
        }

        internal async Task RestoreAndRefreshAsync(Session session, Action<Session, bool> save, CancellationToken token = default)
        {
            foreach (var part in session.Cookie.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var pair = part.Trim().Split('=', 2);
                if (pair.Length != 2) throw new BilibiliException("Bilibili saved session is invalid; scan to log in again.");
                try { cookies.Add(new Cookie(pair[0], pair[1], "/", ".bilibili.com") { Secure = true }); }
                catch (CookieException) { throw new BilibiliException("Bilibili saved session is invalid; scan to log in again."); }
            }
            if (String.IsNullOrWhiteSpace(session.RefreshToken))
                throw new BilibiliException("Bilibili refresh token missing; scan to log in again.");
            await ConfirmRefreshAsync(session, save, token).ConfigureAwait(false);
            var info = await JsonAsync(HttpMethod.Get, Passport + "cookie/info", token).ConfigureAwait(false);
            if (info.ValueKind != JsonValueKind.Object || !info.TryGetProperty("refresh", out var refresh)
                || refresh.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                throw new BilibiliException("GET passport.bilibili.com/x/passport-login/web/cookie/info: invalid response format.");
            if (!refresh.GetBoolean()) return;

            var timestamp = info.TryGetProperty("timestamp", out var time) && time.ValueKind == JsonValueKind.Number && time.TryGetInt64(out var value)
                ? value : throw new BilibiliException("GET passport.bilibili.com/x/passport-login/web/cookie/info: missing timestamp.");
            using var rsa = RSA.Create();
            rsa.ImportFromPem("""
                -----BEGIN PUBLIC KEY-----
                MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDLgd2OAkcGVtoE3ThUREbio0Eg
                Uc/prcajMKXvkCKFCWhJYJcLkcM2DKKcSeFpD/j6Boy538YXnR6VhcuUJOhH2x71
                nzPjfdTcqMz7djHum0qSZA0AyCBDABUqCrfNgCiJ00Ra7GmRj+YCK1NJEuewlb40
                JNrRuoEUXpabUzGB8QIDAQAB
                -----END PUBLIC KEY-----
                """);
            var path = Convert.ToHexString(rsa.Encrypt(Encoding.UTF8.GetBytes($"refresh_{timestamp}"), RSAEncryptionPadding.OaepSHA256)).ToLowerInvariant();
            using var request = new HttpRequestMessage(HttpMethod.Get, "https://www.bilibili.com/correspond/1/" + path);
            var html = new HtmlDocument();
            html.LoadHtml(await SendAsync(request, token).ConfigureAwait(false));
            var refreshCsrf = html.GetElementbyId("1-name")?.InnerText.Trim();
            if (String.IsNullOrEmpty(refreshCsrf))
                throw new BilibiliException("GET www.bilibili.com/correspond/1/{path}: missing refresh CSRF.");
            var data = await JsonAsync(HttpMethod.Post, Passport + "cookie/refresh", token,
                new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["csrf"] = RequiredCookie("bili_jct"), ["refresh_csrf"] = refreshCsrf,
                    ["source"] = "main_web", ["refresh_token"] = session.RefreshToken
                })).ConfigureAwait(false);
            var status = RequiredInt(data, "status", "cookie/refresh");
            if (status != 0)
                throw new BilibiliException($"POST passport.bilibili.com/x/passport-login/web/cookie/refresh: HTTP 200, status {status}.");
            var next = CaptureSession(RequiredString(data, "refresh_token", "cookie/refresh"));
            next.PendingConfirmationToken = session.RefreshToken;
            // Retain a history row for each new session before confirming replacement of the old credentials.
            save(next, true);
            await ConfirmRefreshAsync(next, save, token).ConfigureAwait(false);
        }

        private async Task ConfirmRefreshAsync(Session session, Action<Session, bool> save, CancellationToken token)
        {
            if (session.PendingConfirmationToken is null) return;
            await JsonAsync(HttpMethod.Post, Passport + "confirm/refresh", token,
                new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    ["csrf"] = RequiredCookie("bili_jct"), ["refresh_token"] = session.PendingConfirmationToken
                })).ConfigureAwait(false);
            session.PendingConfirmationToken = null;
            save(session, false);
        }

        internal async Task<Currency> FetchBalanceAsync(CancellationToken token = default)
        {
            var timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            var data = await JsonAsync(HttpMethod.Post, "https://pay.bilibili.com/bk/brokerage/getUserBrokerage", token,
                JsonContent.Create(new { traceId = timestamp, timestamp, sdkVersion = "1.2.1" }), "errno").ConfigureAwait(false);
            if (data.ValueKind != JsonValueKind.Object || !data.TryGetProperty("brokerage", out var amount)
                || amount.ValueKind != JsonValueKind.Number || !amount.TryGetDecimal(out var balance)
                || balance < 0 || Decimal.Round(balance, 2) != balance)
                throw new BilibiliException("POST pay.bilibili.com/bk/brokerage/getUserBrokerage: invalid balance.");
            return new Currency(balance, CurrencyType.RMB);
        }

        private Session CaptureSession(string refreshToken)
        {
            _ = RequiredCookie("SESSDATA");
            _ = RequiredCookie("bili_jct");
            return new Session { Cookie = cookies.GetCookieHeader(CookieOrigin), RefreshToken = refreshToken };
        }

        private string RequiredCookie(string name) => cookies.GetCookies(CookieOrigin)[name]?.Value is { Length: > 0 } value
            ? value : throw new BilibiliException($"Bilibili session missing {name}; scan to log in again.");

        private async Task<JsonElement> JsonAsync(HttpMethod method, string url, CancellationToken token,
            HttpContent? content = null, string codeName = "code")
        {
            using var request = new HttpRequestMessage(method, url) { Content = content };
            var body = await SendAsync(request, token).ConfigureAwait(false);
            try
            {
                using var json = JsonDocument.Parse(body);
                var root = json.RootElement;
                var code = root.GetProperty(codeName).GetInt64();
                if (code != 0) throw new BilibiliException($"{RequestName(request)}: HTTP 200, business code {code}.");
                return root.TryGetProperty("data", out var data) ? data.Clone() : default;
            }
            catch (Exception e) when (e is JsonException or KeyNotFoundException or InvalidOperationException or FormatException)
            {
                throw new BilibiliException($"{RequestName(request)}: invalid response format.");
            }
        }

        private async Task<string> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            var name = RequestName(request);
            var cookie = cookies.GetCookieHeader(request.RequestUri!);
            if (cookie.Length > 0) request.Headers.Add("Cookie", cookie);
            if (request.RequestUri!.Host == "pay.bilibili.com")
            {
                request.Headers.Referrer = new Uri("https://pay.bilibili.com/pay-v2-web/shell_index");
                request.Headers.Add("Origin", "https://pay.bilibili.com");
            }
            try
            {
                using var response = await client.SendAsync(request, token).ConfigureAwait(false);
                if (response.StatusCode != HttpStatusCode.OK) throw new BilibiliException($"{name}: HTTP {(int)response.StatusCode}.");
                if (response.Headers.TryGetValues("Set-Cookie", out var headers))
                    foreach (var header in headers) cookies.SetCookies(request.RequestUri, header);
                return await response.Content.ReadAsStringAsync(token).ConfigureAwait(false);
            }
            catch (HttpRequestException e) { throw new BilibiliException($"{name}: {e.HttpRequestError}."); }
            catch (OperationCanceledException) when (!token.IsCancellationRequested) { throw new BilibiliException($"{name}: timeout."); }
            catch (CookieException) { throw new BilibiliException($"{name}: invalid response cookie."); }
        }

        private static string RequestName(HttpRequestMessage request) => $"{request.Method} {request.RequestUri!.Host}" +
            (request.RequestUri.AbsolutePath.StartsWith("/correspond/", StringComparison.Ordinal)
                ? "/correspond/1/{path}" : request.RequestUri.AbsolutePath);

        private static string RequiredString(JsonElement data, string name, string endpoint) =>
            data.ValueKind == JsonValueKind.Object && data.TryGetProperty(name, out var field)
            && field.ValueKind == JsonValueKind.String && !String.IsNullOrWhiteSpace(field.GetString())
                ? field.GetString()! : throw new BilibiliException($"Bilibili {endpoint}: missing {name}.");

        private static int RequiredInt(JsonElement data, string name, string endpoint) =>
            data.ValueKind == JsonValueKind.Object && data.TryGetProperty(name, out var field)
            && field.ValueKind == JsonValueKind.Number && field.TryGetInt32(out var value)
                ? value : throw new BilibiliException($"Bilibili {endpoint}: missing {name}.");

        public void Dispose() => client.Dispose();
    }

    internal sealed class BilibiliException(string message) : Exception(message);
}
