using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;

namespace MyBook
{
    // TODO: Nexus OAuth token refresh is implemented but not remotely verified until a valid client id is available.
    partial class GraphQLUtil
    {
        private const string NexusOAuthTokenEndpoint = "https://users.nexusmods.com/oauth/token";
        private const int NexusOAuthTokenRefreshSkewSeconds = 60;

        private NexusOAuthTokenSet? cachedNexusOAuthTokens;

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
            ApplyNexusApplicationHeaders(client);
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

        private string GetRequiredNexusOAuthClientId()
        {
            return OptionalConfig("nexus_oauth_client_id")
                ?? throw new InvalidOperationException("Missing nexus_oauth_client_id in config.json.");
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
