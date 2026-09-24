using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // Creates Plaid Items through Hosted Link. Access tokens are stored in
    // the private application database and are never written to command-line output.
    internal sealed class PlaidLink : IDisposable
    {
        private const string ClientName = "MyBook";
        private const string DefaultClientUserId = "mybook-primary-user";
        private const string Language = "en";
        private const int TransactionHistoryDays = 730;
        private const int HostedLinkLifetimeSeconds = 15 * 60;
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromMinutes(3);
        private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan LinkCompletionTimeout = TimeSpan.FromMinutes(15);

        private readonly string clientId;
        private readonly string secret;
        private readonly string apiBaseUrl;
        private readonly string environmentName;
        private readonly PlaidEnvironment environment;
        private readonly DatabaseUtil database;
        private readonly HttpClient client;

        public PlaidLink(IConfigurationRoot config, DatabaseUtil database)
        {
            clientId = RequiredConfig(config, "plaid_client_id");
            secret = RequiredConfig(config, PlaidUtil.SelectedSecretConfigKey);
            apiBaseUrl = PlaidUtil.SelectedApiBaseUrl;
            environment = PlaidUtil.UseProductionEnvironment
                ? PlaidEnvironment.Production
                : PlaidEnvironment.Sandbox;
            environmentName = environment.ToString().ToLowerInvariant();
            this.database = database;

            client = new HttpClient { Timeout = RequestTimeout };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 PlaidLink");
        }

        public Task<PlaidLinkResult> ConnectAndStoreAsync(
            IReadOnlyCollection<string> countryCodes,
            IReadOnlyCollection<string> products,
            CancellationToken cancellationToken = default) => ConnectAndStoreAsync(
                DefaultClientUserId,
                countryCodes,
                products,
                cancellationToken);

        public static PlaidLinkCommandOptions ParseCommandLine(IReadOnlyList<string> arguments)
        {
            var countryCodes = new List<string>();
            var products = new List<string>();
            for (var i = 0; i < arguments.Count; i++)
            {
                var argument = arguments[i];
                if (argument.Equals("--plaid-link", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (argument.Equals("--country", StringComparison.OrdinalIgnoreCase))
                {
                    AddCommandValues(arguments, ref i, countryCodes, "--country");
                    continue;
                }
                if (argument.StartsWith("--country=", StringComparison.OrdinalIgnoreCase))
                {
                    AddCommandValues(argument["--country=".Length..], countryCodes, "--country");
                    continue;
                }
                if (argument.Equals("--product", StringComparison.OrdinalIgnoreCase))
                {
                    AddCommandValues(arguments, ref i, products, "--product");
                    continue;
                }
                if (argument.StartsWith("--product=", StringComparison.OrdinalIgnoreCase))
                {
                    AddCommandValues(argument["--product=".Length..], products, "--product");
                    continue;
                }

                throw CommandLineError($"unrecognized argument {argument}");
            }

            var normalizedCountryCodes = countryCodes
                .Select(value => value.ToUpperInvariant())
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (normalizedCountryCodes.Count == 0)
                throw CommandLineError("at least one --country is required");
            if (normalizedCountryCodes.Any(value => value.Length != 2 || value.Any(character => character is < 'A' or > 'Z')))
                throw CommandLineError("each --country must be a two-letter ISO country code");

            var normalizedProducts = products
                .Select(value => value.ToLowerInvariant())
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (normalizedProducts.Count == 0)
                throw CommandLineError("at least one --product is required");
            if (normalizedProducts.Any(value => value.Length == 0
                || value[0] is < 'a' or > 'z'
                || value.Any(character => character != '_' && (character is < 'a' or > 'z'))))
            {
                throw CommandLineError("each --product must use Plaid's lowercase product name");
            }

            return new PlaidLinkCommandOptions(normalizedCountryCodes, normalizedProducts);
        }

        private static void AddCommandValues(
            IReadOnlyList<string> arguments,
            ref int argumentIndex,
            List<string> values,
            string optionName)
        {
            if (++argumentIndex >= arguments.Count || arguments[argumentIndex].StartsWith("--", StringComparison.Ordinal))
                throw CommandLineError($"{optionName} requires a value");
            AddCommandValues(arguments[argumentIndex], values, optionName);
        }

        private static void AddCommandValues(string text, List<string> values, string optionName)
        {
            var parsed = text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (parsed.Length == 0)
                throw CommandLineError($"{optionName} requires a value");
            values.AddRange(parsed);
        }

        private static ArgumentException CommandLineError(string reason) => new(
            $"{reason}. Usage: --plaid-link --country <ISO-2> [--country <ISO-2> ...] "
            + "--product <Plaid product> [--product <Plaid product> ...]");

        public async Task<PlaidLinkResult> ConnectAndStoreAsync(
            string clientUserId,
            IReadOnlyCollection<string> countryCodes,
            IReadOnlyCollection<string> products,
            CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(clientUserId))
                throw new ArgumentException("Plaid client user id is empty.", nameof(clientUserId));
            if (countryCodes.Count == 0 || countryCodes.Any(String.IsNullOrWhiteSpace))
                throw new ArgumentException("Plaid country codes are empty or invalid.", nameof(countryCodes));
            if (products.Count == 0 || products.Any(String.IsNullOrWhiteSpace))
                throw new ArgumentException("Plaid products are empty or invalid.", nameof(products));

            var createRequest = JObject.FromObject(new
            {
                client_id = clientId,
                secret,
                client_name = ClientName,
                language = Language,
                country_codes = countryCodes,
                products,
                user = new { client_user_id = clientUserId.Trim() },
                hosted_link = new { url_lifetime_seconds = HostedLinkLifetimeSeconds }
            });
            if (products.Contains("transactions", StringComparer.OrdinalIgnoreCase))
                createRequest["transactions"] = JObject.FromObject(new { days_requested = TransactionHistoryDays });

            var createResponse = await PostAsync(
                "/link/token/create",
                createRequest,
                cancellationToken).ConfigureAwait(false);
            var linkToken = RequiredText(createResponse, "link_token", "/link/token/create");
            var hostedLinkUrl = RequiredText(createResponse, "hosted_link_url", "/link/token/create");

            OpenHostedLink(hostedLinkUrl);

            using var completionTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            completionTimeout.CancelAfter(LinkCompletionTimeout);
            PlaidPublicToken publicToken;
            try
            {
                publicToken = await WaitForPublicTokenAsync(linkToken, completionTimeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                throw new PlaidLinkException(
                    $"poll POST /link/token/get: no completed Link session within {LinkCompletionTimeout.TotalMinutes:0} minutes");
            }

            var exchangeResponse = await PostAsync(
                "/item/public_token/exchange",
                new
                {
                    client_id = clientId,
                    secret,
                    public_token = publicToken.PublicToken
                },
                cancellationToken).ConfigureAwait(false);
            var accessToken = RequiredText(exchangeResponse, "access_token", "/item/public_token/exchange");
            var itemId = RequiredText(exchangeResponse, "item_id", "/item/public_token/exchange");

            var databaseId = SaveToken(new PlaidStoredItem
            {
                ItemId = itemId,
                AccessToken = accessToken,
                InstitutionId = publicToken.InstitutionId,
                InstitutionName = publicToken.InstitutionName,
                CreatedAtUtc = DateTimeOffset.UtcNow
            });

            return new PlaidLinkResult(
                databaseId,
                itemId,
                publicToken.InstitutionId,
                publicToken.InstitutionName,
                environmentName);
        }

        private async Task<PlaidPublicToken> WaitForPublicTokenAsync(
            string linkToken,
            CancellationToken cancellationToken)
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var response = await PostAsync(
                    "/link/token/get",
                    new { client_id = clientId, secret, link_token = linkToken },
                    cancellationToken).ConfigureAwait(false);
                var sessions = response["link_sessions"] as JArray;
                if (sessions is not null)
                {
                    foreach (var session in sessions.OfType<JObject>().Reverse())
                    {
                        var itemResults = session["results"]?["item_add_results"] as JArray;
                        var completedItems = itemResults?.OfType<JObject>()
                            .Where(item => !String.IsNullOrWhiteSpace(item["public_token"]?.ToString()))
                            .ToList() ?? [];
                        if (completedItems.Count > 1)
                            throw new PlaidLinkException(
                                "poll POST /link/token/get: multiple Items returned; this single-Item flow stopped before exchanging tokens");
                        if (completedItems.Count == 1)
                            return ParsePublicToken(completedItems[0]);

                        var legacyPublicToken = session["on_success"]?["public_token"]?.ToString();
                        if (!String.IsNullOrWhiteSpace(legacyPublicToken))
                        {
                            var institution = session["on_success"]?["metadata"]?["institution"];
                            return new PlaidPublicToken(
                                legacyPublicToken,
                                institution?["institution_id"]?.ToString(),
                                institution?["name"]?.ToString());
                        }

                        var finishedAt = session["finished_at"];
                        if (finishedAt is not null
                            && finishedAt.Type is not (JTokenType.Null or JTokenType.Undefined))
                        {
                            var exit = session["exit"];
                            var errorCode = exit?["error"]?["error_code"]?.ToString();
                            var errorType = exit?["error"]?["error_type"]?.ToString();
                            var exitStatus = exit?["metadata"]?["status"]?.ToString();
                            throw new PlaidLinkException(
                                $"poll POST /link/token/get: Link session exited without an Item; error_type={errorType}; error_code={errorCode}; exit_status={exitStatus}");
                        }
                    }
                }

                await Task.Delay(PollInterval, cancellationToken).ConfigureAwait(false);
            }
        }

        private static void OpenHostedLink(string hostedLinkUrl)
        {
            try
            {
                Process.Start(new ProcessStartInfo(hostedLinkUrl) { UseShellExecute = true });
            }
            catch (Exception exception) when (exception is InvalidOperationException or System.ComponentModel.Win32Exception)
            {
                throw new PlaidLinkException("open Hosted Link in the default browser: failed", exception);
            }
        }

        private static PlaidPublicToken ParsePublicToken(JObject item)
        {
            var publicToken = item["public_token"]?.ToString();
            if (String.IsNullOrWhiteSpace(publicToken))
                throw new PlaidLinkException(
                    "poll POST /link/token/get: completed Item has no public_token");
            return new PlaidPublicToken(
                publicToken,
                item["institution"]?["institution_id"]?.ToString(),
                item["institution"]?["name"]?.ToString());
        }

        private async Task<JObject> PostAsync(string path, object body, CancellationToken cancellationToken)
        {
            using var content = new StringContent(
                JsonConvert.SerializeObject(body),
                Encoding.UTF8,
                "application/json");
            using var response = await client.PostAsync(apiBaseUrl + path, content, cancellationToken)
                .ConfigureAwait(false);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                throw new PlaidLinkException(FormatPlaidError(path, response, responseText));
            try
            {
                return JObject.Parse(responseText);
            }
            catch (JsonException exception)
            {
                throw new PlaidLinkException(
                    $"POST {path}: HTTP {(int)response.StatusCode}; response is not valid JSON",
                    exception);
            }
        }

        private int SaveToken(PlaidStoredItem item)
        {
            try
            {
                var now = DateTime.UtcNow;
                return database.AddPlaidItem(new PlaidItem
                {
                    environment = environment,
                    itemId = item.ItemId,
                    accessToken = item.AccessToken,
                    institutionId = item.InstitutionId,
                    institutionName = item.InstitutionName,
                    createdAtUtc = item.CreatedAtUtc.UtcDateTime,
                    updateTimeUtc = now
                });
            }
            catch (Exception exception)
            {
                var path = Path.Combine(AppContext.BaseDirectory, $"plaid-token-recovery-{Guid.NewGuid():N}.local.json");
                try
                {
                    File.WriteAllText(path, JsonConvert.SerializeObject(new { environment = environmentName, item }, Formatting.Indented));
                }
                catch (Exception fileException)
                {
                    throw new PlaidLinkException(
                        $"persist Plaid Item: database insert failed; local token recovery file also failed ({fileException.GetType().Name})",
                        exception);
                }
                throw new PlaidLinkException(
                    $"persist Plaid Item: database insert failed; plaintext token saved for manual recovery at {path}. Delete it after recovery; do not share or commit it.",
                    exception);
            }
        }

        private static string RequiredText(JObject json, string name, string path)
        {
            var value = json[name]?.ToString();
            return String.IsNullOrWhiteSpace(value)
                ? throw new PlaidLinkException($"POST {path}: successful response has no {name}")
                : value;
        }

        private static string RequiredConfig(IConfigurationRoot config, string name)
        {
            var value = config[name];
            return String.IsNullOrWhiteSpace(value)
                ? throw new PlaidLinkException($"load configuration: missing {name}")
                : value.Trim();
        }

        private static string FormatPlaidError(string path, HttpResponseMessage response, string responseText)
        {
            try
            {
                var json = JObject.Parse(responseText);
                return $"POST {path}: HTTP {(int)response.StatusCode} {response.ReasonPhrase}; "
                    + $"error_type={json["error_type"]}; error_code={json["error_code"]}; request_id={json["request_id"]}";
            }
            catch (JsonException)
            {
                return $"POST {path}: HTTP {(int)response.StatusCode} {response.ReasonPhrase}; response=non-JSON";
            }
        }

        public void Dispose() => client.Dispose();
    }

    public sealed record PlaidLinkResult(
        int DatabaseId,
        string ItemId,
        string? InstitutionId,
        string? InstitutionName,
        string Environment);

    public sealed record PlaidLinkCommandOptions(
        IReadOnlyList<string> CountryCodes,
        IReadOnlyList<string> Products);

    sealed record PlaidPublicToken(
        string PublicToken,
        string? InstitutionId,
        string? InstitutionName);

    sealed class PlaidStoredItem
    {
        public string ItemId { get; set; } = "";
        public string AccessToken { get; set; } = "";
        public string? InstitutionId { get; set; }
        public string? InstitutionName { get; set; }
        public DateTimeOffset CreatedAtUtc { get; set; }
    }

    public sealed class PlaidLinkException : Exception
    {
        public PlaidLinkException(string message) : base(message) { }
        public PlaidLinkException(string message, Exception innerException) : base(message, innerException) { }
    }
}
