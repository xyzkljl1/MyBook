using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;

namespace MyBook
{
    // TODO: Nexus OAuth flow is implemented but not remotely verified until a valid client id is available.
    partial class GraphQLUtil
    {
        // Do not change this URI unless the Nexus OAuth app redirect URI is changed too.
        private const int NexusOAuthCallbackPort = 4700;
        private const string NexusOAuthAuthorizeEndpoint = "https://users.nexusmods.com/oauth/authorize";
        private const string NexusOAuthTokenEndpoint = "https://users.nexusmods.com/oauth/token";
        private const int NexusOAuthTokenRefreshSkewSeconds = 60;

        private NexusOAuthTokenSet? cachedNexusOAuthTokens;

        public async Task<NexusOAuthTokenSet> AuthorizeNexusOAuthToken()
        {
            var clientId = GetRequiredNexusOAuthClientId();
            var clientSecret = OptionalConfig("nexus_oauth_client_secret");
            var codeVerifier = String.IsNullOrWhiteSpace(clientSecret) ? CreatePkceCodeVerifier() : null;
            var codeChallenge = codeVerifier is null ? null : CreatePkceCodeChallenge(codeVerifier);
            var state = Guid.NewGuid().ToString("N");
            var authorizeUrl = BuildNexusOAuthAuthorizeUrl(clientId, state, codeChallenge);

            using var listener = new HttpListener();
            listener.Prefixes.Add($"{GetNexusOAuthRedirectBaseUri()}/");
            listener.Start();

            Console.WriteLine($"Open Nexus OAuth authorization URL: {authorizeUrl}");
            TryOpenNexusOAuthAuthorizeUrl(authorizeUrl);

            var context = await listener.GetContextAsync();
            var request = context.Request;
            var error = request.QueryString["error"];
            var code = request.QueryString["code"];
            var actualState = request.QueryString["state"];
            var success = String.IsNullOrWhiteSpace(error)
                && !String.IsNullOrWhiteSpace(code)
                && String.Equals(actualState, state, StringComparison.Ordinal);
            await WriteNexusOAuthCallbackResponse(context.Response, success);

            if (!String.IsNullOrWhiteSpace(error))
                throw new InvalidOperationException($"Nexus OAuth authorization failed: {error}");
            if (String.IsNullOrWhiteSpace(code))
                throw new InvalidOperationException("Nexus OAuth callback did not include a code.");
            if (!String.Equals(actualState, state, StringComparison.Ordinal))
                throw new InvalidOperationException("Nexus OAuth callback state mismatch.");

            var form = new Dictionary<string, string>
            {
                ["grant_type"] = "authorization_code",
                ["redirect_uri"] = GetNexusOAuthRedirectUri(),
                ["scope"] = GetNexusOAuthScope(),
                ["client_id"] = clientId,
                ["code"] = code
            };
            if (codeVerifier is not null)
                form["code_verifier"] = codeVerifier;
            if (!String.IsNullOrWhiteSpace(clientSecret))
                form["client_secret"] = clientSecret;

            var tokens = await RequestNexusOAuthTokens(form);
            SaveNexusOAuthTokens(tokens);
            return tokens;
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
            var clientSecret = OptionalConfig("nexus_oauth_client_secret");
            if (!String.IsNullOrWhiteSpace(clientSecret))
                form["client_secret"] = clientSecret;

            var refreshed = await RequestNexusOAuthTokens(form);
            return String.IsNullOrWhiteSpace(refreshed.RefreshToken)
                ? refreshed with { RefreshToken = refreshToken }
                : refreshed;
        }

        private async Task<NexusOAuthTokenSet> RequestNexusOAuthTokens(Dictionary<string, string> form)
        {
            using HttpClient client = new();
            client.Timeout = RequestTimeout;
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0");
            using var content = new FormUrlEncodedContent(form);
            using var response = await client.PostAsync(NexusOAuthTokenEndpoint, content);
            var responseText = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new InvalidOperationException(
                    $"Nexus OAuth token request failed: {(int)response.StatusCode} {response.ReasonPhrase} {responseText}");
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
                    ?? throw new InvalidOperationException($"Nexus OAuth token response has no access_token: {responseText}"),
                RefreshToken = json["refresh_token"]?.ToString() ?? "",
                TokenType = json["token_type"]?.ToString() ?? "Bearer",
                Scope = json["scope"]?.ToString() ?? "",
                ExpiresIn = expiresIn,
                ExpiresAt = expiresIn.HasValue
                    ? DateTimeOffset.UtcNow.AddSeconds(expiresIn.Value)
                    : null
            };
        }

        private string BuildNexusOAuthAuthorizeUrl(string clientId, string state, string? codeChallenge)
        {
            var query = new Dictionary<string, string>
            {
                ["client_id"] = clientId,
                ["response_type"] = "code",
                ["scope"] = GetNexusOAuthScope(),
                ["redirect_uri"] = GetNexusOAuthRedirectUri(),
                ["state"] = state
            };
            if (!String.IsNullOrWhiteSpace(codeChallenge))
            {
                query["code_challenge_method"] = "S256";
                query["code_challenge"] = codeChallenge;
            }

            return $"{NexusOAuthAuthorizeEndpoint}?{BuildUrlEncodedForm(query)}";
        }

        private string GetRequiredNexusOAuthClientId()
        {
            return OptionalConfig("nexus_oauth_client_id")
                ?? throw new InvalidOperationException("Missing nexus_oauth_client_id in config.json.");
        }

        private string GetNexusOAuthRedirectBaseUri()
        {
            return $"http://127.0.0.1:{NexusOAuthCallbackPort}";
        }

        private string GetNexusOAuthRedirectUri()
        {
            return $"{GetNexusOAuthRedirectBaseUri()}/callback";
        }

        private string GetNexusOAuthScope()
        {
            return OptionalConfig("nexus_oauth_scope") ?? "";
        }

        private static string CreatePkceCodeVerifier()
        {
            return Base64UrlEncode(RandomNumberGenerator.GetBytes(64));
        }

        private static string CreatePkceCodeChallenge(string codeVerifier)
        {
            return Base64UrlEncode(SHA256.HashData(Encoding.ASCII.GetBytes(codeVerifier)));
        }

        private static string Base64UrlEncode(byte[] bytes)
        {
            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private static string BuildUrlEncodedForm(Dictionary<string, string> values)
        {
            return String.Join("&", values.Select(pair =>
                $"{Uri.EscapeDataString(pair.Key)}={Uri.EscapeDataString(pair.Value)}"));
        }

        private static void TryOpenNexusOAuthAuthorizeUrl(string authorizeUrl)
        {
            try
            {
                Process.Start(new ProcessStartInfo(authorizeUrl) { UseShellExecute = true });
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Unable to open browser automatically: {exception.Message}");
            }
        }

        private static async Task WriteNexusOAuthCallbackResponse(HttpListenerResponse response, bool success)
        {
            var html = success
                ? "<html><body>Nexus OAuth authorization completed. You can close this window.</body></html>"
                : "<html><body>Nexus OAuth authorization failed. Return to MyBook for details.</body></html>";
            var bytes = Encoding.UTF8.GetBytes(html);
            response.ContentType = "text/html; charset=utf-8";
            response.ContentLength64 = bytes.Length;
            await response.OutputStream.WriteAsync(bytes);
            response.Close();
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
