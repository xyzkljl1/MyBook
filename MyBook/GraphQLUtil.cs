using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // Shared GraphQL API entry point; site-specific queries live in suffix files.
    partial class GraphQLUtil
    {
        private const string NexusGraphQLEndpoint = "https://api.nexusmods.com/v2/graphql";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(30);

        private readonly IConfigurationRoot config;
        private readonly DatabaseUtil? database;
        private readonly string nexusApiKey;

        public GraphQLUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            this.config = config;
            this.database = database;
            nexusApiKey = config["nexus_api_key"]
                ?? throw new InvalidOperationException("Missing nexus_api_key in config.json");
            if (String.IsNullOrWhiteSpace(nexusApiKey))
                throw new InvalidOperationException("nexus_api_key in config.json is empty");
        }

        private async Task<JObject> ExecuteNexusQuery(string query, object? variables = null)
        {
            using HttpClient client = new();
            client.Timeout = RequestTimeout;
            client.DefaultRequestHeaders.Add("apikey", nexusApiKey);
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0");

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
    }
}
