using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;

namespace MyBook
{
    partial class GraphQLUtil
    {
        private const string NexusOAuthAuthorizeEndpoint = "https://users.nexusmods.com/oauth/authorize";
        private const string NexusOAuthTokenEndpoint = "https://users.nexusmods.com/oauth/token";
        private const string NexusOAuthRedirectUri = "http://127.0.0.1:4700/callback";
        private const string NexusOAuthListenerPrefix = "http://127.0.0.1:4700/";
        private const int NexusOAuthTokenRefreshSkewSeconds = 60;
        private static readonly TimeSpan NexusOAuthAuthorizationTimeout = TimeSpan.FromMinutes(5);

        private NexusOAuthTokenSet? cachedNexusOAuthTokens;

        public async Task AuthorizeNexusOAuthAsync()
        {
            var clientId = GetRequiredNexusOAuthClientId();
            var state = CreateNexusOAuthRandomText(32);
            var codeVerifier = CreateNexusOAuthRandomText(64);
            var codeChallenge = CreateNexusOAuthCodeChallenge(codeVerifier);
            var authorizationUrl = BuildNexusOAuthAuthorizationUrl(clientId, state, codeChallenge);

            using var listener = new HttpListener();
            listener.Prefixes.Add(NexusOAuthListenerPrefix);
            listener.Start();

            Console.WriteLine($"Nexus OAuth redirect URI: {NexusOAuthRedirectUri}");
            Console.WriteLine("Opening Nexus OAuth authorization page.");
            Process.Start(new ProcessStartInfo(authorizationUrl) { UseShellExecute = true });

            using var timeout = new CancellationTokenSource(NexusOAuthAuthorizationTimeout);
            HttpListenerContext context;
            try
            {
                context = await listener.GetContextAsync().WaitAsync(timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException($"Timed out waiting for Nexus OAuth callback after {NexusOAuthAuthorizationTimeout.TotalSeconds:0}s.");
            }

            var request = context.Request;
            var response = context.Response;
            if (!String.Equals(request.Url?.AbsolutePath, "/callback", StringComparison.OrdinalIgnoreCase))
            {
                await WriteNexusOAuthBrowserResponse(response, 404, "Unknown Nexus OAuth callback path. You can close this tab.").ConfigureAwait(false);
                listener.Stop();
                throw new InvalidOperationException("Nexus OAuth callback path mismatch.");
            }

            var error = request.QueryString["error"];
            if (!String.IsNullOrWhiteSpace(error))
            {
                var description = request.QueryString["error_description"];
                await WriteNexusOAuthBrowserResponse(response, 400, "Nexus OAuth failed. You can close this tab.").ConfigureAwait(false);
                listener.Stop();
                throw new InvalidOperationException($"Nexus OAuth authorization failed: {error} {description}");
            }

            var code = request.QueryString["code"];
            if (String.IsNullOrWhiteSpace(code))
            {
                await WriteNexusOAuthBrowserResponse(response, 400, "Missing Nexus OAuth code. You can close this tab.").ConfigureAwait(false);
                listener.Stop();
                throw new InvalidOperationException("Nexus OAuth callback has no code.");
            }

            if (!String.Equals(request.QueryString["state"], state, StringComparison.Ordinal))
            {
                await WriteNexusOAuthBrowserResponse(response, 400, "Nexus OAuth state mismatch. You can close this tab.").ConfigureAwait(false);
                listener.Stop();
                throw new InvalidOperationException("Nexus OAuth callback state mismatch.");
            }

            await WriteNexusOAuthBrowserResponse(response, 200, "Nexus OAuth complete. You can close this tab.").ConfigureAwait(false);
            listener.Stop();

            var tokens = await ExchangeNexusOAuthAuthorizationCode(code, codeVerifier).ConfigureAwait(false);
            SaveNexusOAuthTokens(tokens);
        }

        private async Task<string?> GetNexusOAuthAccessToken()
        {
            var tokens = LoadNexusOAuthTokens();
            if (tokens is null)
                return null;

            if (!String.IsNullOrWhiteSpace(tokens.AccessToken)
                && (!tokens.ExpiresAt.HasValue
                    || DateTimeOffset.UtcNow.AddSeconds(NexusOAuthTokenRefreshSkewSeconds) < tokens.ExpiresAt.Value))
            {
                return tokens.AccessToken;
            }

            if (String.IsNullOrWhiteSpace(tokens.RefreshToken))
                return null;

            var refreshed = await RefreshNexusOAuthTokens(tokens.RefreshToken);
            SaveNexusOAuthTokens(refreshed);
            return refreshed.AccessToken;
        }

        private async Task<NexusOAuthTokenSet> RefreshNexusOAuthTokens(string refreshToken)
        {
            var clientId = GetRequiredNexusOAuthClientId();
            var form = new Dictionary<string, string>
            {
                ["client_id"] = clientId,
                ["grant_type"] = "refresh_token",
                ["refresh_token"] = refreshToken
            };

            var refreshed = await RequestNexusOAuthTokens(form);
            return String.IsNullOrWhiteSpace(refreshed.RefreshToken)
                ? refreshed with { RefreshToken = refreshToken }
                : refreshed;
        }

        private async Task<NexusOAuthTokenSet> ExchangeNexusOAuthAuthorizationCode(string code, string codeVerifier)
        {
            var clientId = GetRequiredNexusOAuthClientId();
            var form = new Dictionary<string, string>
            {
                ["client_id"] = clientId,
                ["grant_type"] = "authorization_code",
                ["redirect_uri"] = NexusOAuthRedirectUri,
                ["code"] = code,
                ["code_verifier"] = codeVerifier
            };

            return await RequestNexusOAuthTokens(form).ConfigureAwait(false);
        }

        private async Task<NexusOAuthTokenSet> RequestNexusOAuthTokens(Dictionary<string, string> form)
        {
            using HttpClient client = new();
            client.Timeout = RequestTimeout;
            ApplyNexusApplicationHeaders(client);
            using var content = new FormUrlEncodedContent(form);
            using var response = await client.PostAsync(NexusOAuthTokenEndpoint, content);
            var responseText = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new InvalidOperationException(
                    $"Nexus OAuth token request failed: {(int)response.StatusCode} {response.ReasonPhrase} {FormatNexusOAuthErrorResponse(responseText)}");
            }

            return ParseNexusOAuthTokenResponse(responseText);
        }

        private NexusOAuthTokenSet? LoadNexusOAuthTokens()
        {
            if (cachedNexusOAuthTokens is not null)
                return cachedNexusOAuthTokens;

            if (database is null)
                return null;

            var token = database.GetOAuthToken(OAuthTokenProvider.Nexus);
            if (token is null)
                return null;

            cachedNexusOAuthTokens = new NexusOAuthTokenSet
            {
                AccessToken = token.accessToken,
                RefreshToken = token.refreshToken,
                TokenType = token.tokenType,
                Scope = token.scope,
                ExpiresAt = token.expiresAt.HasValue
                    ? new DateTimeOffset(DateTime.SpecifyKind(token.expiresAt.Value, DateTimeKind.Utc))
                    : null
            };
            return cachedNexusOAuthTokens;
        }

        private void SaveNexusOAuthTokens(NexusOAuthTokenSet tokens)
        {
            var db = database ?? throw new InvalidOperationException("Saving Nexus OAuth token requires a database.");
            cachedNexusOAuthTokens = tokens;
            db.SaveOAuthToken(new OAuthToken
            {
                provider = OAuthTokenProvider.Nexus,
                accessToken = tokens.AccessToken,
                refreshToken = tokens.RefreshToken,
                tokenType = tokens.TokenType,
                scope = tokens.Scope,
                expiresAt = tokens.ExpiresAt?.UtcDateTime,
                updateTime = DateTime.UtcNow
            });
            Console.WriteLine("Nexus OAuth token saved to database.");
        }

        private NexusOAuthTokenSet ParseNexusOAuthTokenResponse(string responseText)
        {
            var json = JObject.Parse(responseText);
            var expiresIn = json["expires_in"]?.Value<int?>();
            return new NexusOAuthTokenSet
            {
                AccessToken = json["access_token"]?.ToString()
                    ?? throw new InvalidOperationException("Nexus OAuth token response has no access_token."),
                RefreshToken = json["refresh_token"]?.ToString() ?? "",
                TokenType = json["token_type"]?.ToString() ?? "Bearer",
                Scope = json["scope"]?.ToString() ?? "",
                ExpiresIn = expiresIn,
                ExpiresAt = expiresIn.HasValue
                    ? DateTimeOffset.UtcNow.AddSeconds(expiresIn.Value)
                    : null
            };
        }

        private string GetRequiredNexusOAuthClientId()
        {
            return OptionalConfig("nexus_oauth_client_id")
                ?? throw new InvalidOperationException("Missing nexus_oauth_client_id in config.json.");
        }

        private static string BuildNexusOAuthAuthorizationUrl(string clientId, string state, string codeChallenge)
        {
            var parameters = new Dictionary<string, string>
            {
                ["client_id"] = clientId,
                ["response_type"] = "code",
                ["scope"] = "",
                ["redirect_uri"] = NexusOAuthRedirectUri,
                ["state"] = state,
                ["code_challenge_method"] = "S256",
                ["code_challenge"] = codeChallenge
            };
            return NexusOAuthAuthorizeEndpoint + "?" + String.Join("&", parameters.Select(item =>
                $"{Uri.EscapeDataString(item.Key)}={Uri.EscapeDataString(item.Value)}"));
        }

        private static string FormatNexusOAuthErrorResponse(string responseText)
        {
            try
            {
                var json = JObject.Parse(responseText);
                var error = json["error"]?.ToString();
                var description = json["error_description"]?.ToString();
                var detail = String.Join(" ", new[] { error, description }
                    .Where(value => !String.IsNullOrWhiteSpace(value)));
                return String.IsNullOrWhiteSpace(detail)
                    ? "OAuth error response had no safe error detail."
                    : detail;
            }
            catch (JsonException)
            {
                return "OAuth error response was not JSON.";
            }
        }

        private static string CreateNexusOAuthRandomText(int byteCount)
        {
            return ToBase64Url(RandomNumberGenerator.GetBytes(byteCount));
        }

        private static string CreateNexusOAuthCodeChallenge(string codeVerifier)
        {
            return ToBase64Url(SHA256.HashData(Encoding.ASCII.GetBytes(codeVerifier)));
        }

        private static string ToBase64Url(byte[] bytes)
        {
            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private static async Task WriteNexusOAuthBrowserResponse(HttpListenerResponse response, int statusCode, string message)
        {
            response.StatusCode = statusCode;
            response.ContentType = "text/plain; charset=utf-8";
            var body = Encoding.UTF8.GetBytes(message);
            response.ContentLength64 = body.Length;
            await response.OutputStream.WriteAsync(body).ConfigureAwait(false);
            response.OutputStream.Close();
        }

    }

    public sealed record NexusOAuthTokenSet
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; init; } = "";

        [JsonProperty("refresh_token")]
        public string RefreshToken { get; init; } = "";

        [JsonProperty("token_type")]
        public string TokenType { get; init; } = "Bearer";

        [JsonProperty("scope")]
        public string Scope { get; init; } = "";

        [JsonProperty("expires_in")]
        public int? ExpiresIn { get; init; }

        [JsonProperty("expires_at")]
        public DateTimeOffset? ExpiresAt { get; init; }
    }
}
