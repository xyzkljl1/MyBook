using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // TODO: Plaid is only a Sandbox evaluation helper for now; it does not import account data.
    class PlaidUtil
    {
        private const string PlaidSandboxApiBaseUrl = "https://sandbox.plaid.com";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(20);

        private readonly string clientId;
        private readonly string sandboxSecret;

        public PlaidUtil(IConfigurationRoot config)
        {
            clientId = RequiredConfig(config, "plaid_client_id");
            sandboxSecret = RequiredConfig(config, "plaid_sandbox_secret");
        }

        public async Task<List<PlaidInstitution>> SearchSandboxInstitutions(
            string query,
            IReadOnlyCollection<string> countryCodes,
            IReadOnlyCollection<string> products,
            CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(query))
                throw new ArgumentException("Plaid institution search query is empty.", nameof(query));
            if (countryCodes.Count == 0)
                throw new ArgumentException("Plaid institution search country codes are empty.", nameof(countryCodes));
            if (products.Count == 0)
                throw new ArgumentException("Plaid institution search products are empty.", nameof(products));

            var requestBody = JsonConvert.SerializeObject(new
            {
                client_id = clientId,
                secret = sandboxSecret,
                query = query.Trim(),
                products,
                country_codes = countryCodes
            });

            using var client = CreateHttpClient();
            using var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
            using var response = await client.PostAsync(
                    $"{PlaidSandboxApiBaseUrl}/institutions/search",
                    content,
                    cancellationToken)
                .ConfigureAwait(false);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (response.StatusCode != HttpStatusCode.OK)
                throw new InvalidOperationException(FormatPlaidError(response, responseText));

            var json = JObject.Parse(responseText);
            return (json["institutions"] as JArray ?? new JArray())
                .Select(ParseInstitution)
                .ToList();
        }

        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient
            {
                Timeout = RequestTimeout
            };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 PlaidSandboxVerifier");
            return client;
        }

        private static PlaidInstitution ParseInstitution(JToken token)
        {
            return new PlaidInstitution(
                token["institution_id"]?.ToString() ?? "",
                token["name"]?.ToString() ?? "",
                ReadStringArray(token["products"]),
                ReadStringArray(token["country_codes"]),
                token["oauth"]?.Value<bool>() ?? false,
                token["url"]?.ToString());
        }

        private static List<string> ReadStringArray(JToken? token)
        {
            return token is JArray array
                ? array.Select(item => item.ToString()).Where(value => !String.IsNullOrWhiteSpace(value)).ToList()
                : new List<string>();
        }

        private static string RequiredConfig(IConfigurationRoot config, string key)
        {
            var value = config[key];
            return String.IsNullOrWhiteSpace(value)
                ? throw new InvalidOperationException($"Missing {key} in config.json.")
                : value.Trim();
        }

        private static string FormatPlaidError(HttpResponseMessage response, string responseText)
        {
            try
            {
                var json = JObject.Parse(responseText);
                var errorCode = json["error_code"]?.ToString();
                var errorMessage = json["error_message"]?.ToString();
                var requestId = json["request_id"]?.ToString();
                return $"Plaid request failed: {(int)response.StatusCode} {response.ReasonPhrase}; code={errorCode}; message={errorMessage}; request_id={requestId}";
            }
            catch (JsonException)
            {
                return $"Plaid request failed: {(int)response.StatusCode} {response.ReasonPhrase}; response={responseText}";
            }
        }
    }

    sealed record PlaidInstitution(
        string InstitutionId,
        string Name,
        List<string> Products,
        List<string> CountryCodes,
        bool OAuth,
        string? Url);
}
