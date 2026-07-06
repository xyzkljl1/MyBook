using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace MyBook
{
    // Shared GraphQL API entry point; site-specific queries live in suffix files.
    partial class GraphQLUtil
    {
        private const string NexusGraphQLEndpoint = "https://api.nexusmods.com/v2/graphql";
        private const string NexusApplicationName = "MyBook";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(30);
        private static readonly string NexusApplicationVersion =
            typeof(GraphQLUtil).Assembly.GetName().Version?.ToString() ?? "0.0.0";

        private readonly IConfigurationRoot config;
        private readonly DatabaseUtil? database;
        private readonly string? nexusApiKey;

        public GraphQLUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            this.config = config;
            this.database = database;
            nexusApiKey = OptionalConfig("nexus_api_key");
        }

        private async Task<JObject> ExecuteNexusQuery(string query, object? variables = null)
        {
            using HttpClient client = new();
            client.Timeout = RequestTimeout;
            ApplyNexusApplicationHeaders(client);
            var accessToken = await GetNexusOAuthAccessToken();
            if (!String.IsNullOrWhiteSpace(accessToken))
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            }
            else if (!String.IsNullOrWhiteSpace(nexusApiKey))
            {
                client.DefaultRequestHeaders.Add("apikey", nexusApiKey);
            }
            else
            {
                throw new InvalidOperationException(
                    "Missing Nexus credentials. Authorize Nexus OAuth token in database or configure nexus_api_key.");
            }

            var requestBody = JsonConvert.SerializeObject(new
            {
                query,
                variables = variables ?? new { }
            });
            using var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
            using var response = await client.PostAsync(NexusGraphQLEndpoint, content);
            var responseText = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new InvalidOperationException(
                    $"Nexus GraphQL request failed: {(int)response.StatusCode} {response.ReasonPhrase} {responseText}");
            }

            var result = JObject.Parse(responseText);
            var errors = result["errors"] as JArray;
            if (errors is not null && errors.Count > 0)
                throw new InvalidOperationException($"Nexus GraphQL error: {FormatGraphQLErrors(errors)}");

            var data = result["data"] as JObject;
            return data ?? throw new InvalidOperationException($"Nexus GraphQL response has no data: {responseText}");
        }

        private static string FormatGraphQLErrors(JArray errors)
        {
            return String.Join("; ", errors
                .Select(error => error["message"]?.ToString())
                .Where(message => !String.IsNullOrWhiteSpace(message)));
        }

        private string? OptionalConfig(string key)
        {
            var value = config[key];
            return String.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static void ApplyNexusApplicationHeaders(HttpClient client)
        {
            client.DefaultRequestHeaders.UserAgent.ParseAdd($"{NexusApplicationName}/{NexusApplicationVersion}");
            client.DefaultRequestHeaders.TryAddWithoutValidation("Application-Name", NexusApplicationName);
            client.DefaultRequestHeaders.TryAddWithoutValidation("Application-Version", NexusApplicationVersion);
        }
    }
}
