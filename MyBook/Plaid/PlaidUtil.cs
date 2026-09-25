using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;

namespace MyBook
{
    /*
     * plaid比较混乱。
     * 从my.plaid.com登陆用户侧界面，从plaid.com登陆开发者侧界面。都使用邮箱。
     * 但是注册时和连接时都需要提供手机号，注册时需要把手机号绑定到邮箱，并且一个手机号可以注册多次绑定到多个邮箱。

     * 连接分两种，连接到app和直接连接到plaid。
     * 直连到plaid需要登陆用户侧界面，然后连接，但是没什么卵用，看不了财务信息，只能解除或增加连接
     * 开发者的trial plan的10个连接数，指的是使用该账号的secret的app的连接，直接把开发者账号连接到机构是没用的
     * 直接连接实际是登陆plaid用户侧界面然后走一遍连接到app的流程
     * 如果使用了手机号，有时账号信息会加入所有手机号对应的邮箱的plaid账号的用户侧界面，有时不会
     * 不管哪种，实际账号都已经和手机号连接了，此时会出现显示没有链接某账号但无法建立连接的情况，但是可以通过删除连接重来来修复。     
     * 
     * 连接到app，则是要在其它平台如gpt、ib或自己的app，唤起一个plaid的嵌入页面操作
     * 此处不使用邮箱，要求手机号但不是必要的！可以直接全跳过最后选without saving并不影响连接
     * 
     * 注意解除连接或删除item不会释放名额！！必须自己保存好token！
     * 但是手动解除连接后仍然可以利用token修复，一个token解除连接后可以连接到同机构的其它账号，但是不能连接到其它机构。
     * 
     * 必须在建立item时指定历史查询范围！否则之后无法更改
     */
    // Shared Plaid transport and Item selection; institution-specific accounting lives in partial modules.
    partial class PlaidUtil
    {
        // This is intentionally a compile-time switch. Change it and rebuild when Sandbox access is needed.
        internal const bool UseProductionEnvironment = true;
        private const string PlaidSandboxApiBaseUrl = "https://sandbox.plaid.com";
        private const string PlaidProductionApiBaseUrl = "https://production.plaid.com";
        internal const string SelectedApiBaseUrl = UseProductionEnvironment
            ? PlaidProductionApiBaseUrl
            : PlaidSandboxApiBaseUrl;
        internal const string SelectedSecretConfigKey = UseProductionEnvironment
            ? "plaid_production_secret"
            : "plaid_sandbox_secret";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(20);
        private static readonly HttpClient httpClient = CreateHttpClient();

        private readonly string clientId;
        private readonly string secret;
        private readonly string apiBaseUrl;
        private readonly DatabaseUtil? database;
        private readonly IConfigurationRoot config;

        public PlaidUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            this.config = config;
            clientId = RequiredConfig(config, "plaid_client_id");
            secret = RequiredConfig(config, SelectedSecretConfigKey);
            apiBaseUrl = SelectedApiBaseUrl;
            this.database = database;
        }

        public Task<PlaidItem?> FindItemByInstitutionAsync(
            string institutionId,
            CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(institutionId))
                throw new ArgumentException("An institution identifier is required.", nameof(institutionId));
            cancellationToken.ThrowIfCancellationRequested();
            var database = this.database
                ?? throw new InvalidOperationException("Plaid Item selection requires a database.");
            var environment = UseProductionEnvironment ? PlaidEnvironment.Production : PlaidEnvironment.Sandbox;
            var items = database.GetPlaidItems(environment);
            return Task.FromResult(SelectItemByInstitution(items, environment, institutionId));
        }

        internal static (string? Id, string? Name) ReadItemInstitution(JObject response, string expectedItemId)
        {
            if (response["item"] is not JObject item
                || item["item_id"]?.Type != JTokenType.String
                || !String.Equals(item["item_id"]!.Value<string>(), expectedItemId, StringComparison.Ordinal))
                throw new InvalidOperationException("Plaid POST /item/get returned a missing or mismatched Item identifier.");
            if (!item.TryGetValue("institution_id", out var id)
                || id.Type is not (JTokenType.String or JTokenType.Null))
                throw new InvalidOperationException("Plaid POST /item/get returned an invalid institution identifier.");
            var institutionId = id.Type == JTokenType.Null ? null : id.Value<string>();
            if (institutionId is not null && String.IsNullOrWhiteSpace(institutionId))
                throw new InvalidOperationException("Plaid POST /item/get returned an empty institution identifier.");
            var name = item["institution_name"];
            if (name is not null && name.Type is not (JTokenType.String or JTokenType.Null))
                throw new InvalidOperationException("Plaid POST /item/get returned an invalid institution name.");
            // A null institution is valid for connections such as micro-deposit Items.
            return (institutionId, name?.Value<string>());
        }

        internal static PlaidItem? SelectItemByInstitution(
            IEnumerable<PlaidItem> items, PlaidEnvironment environment, string institutionId)
        {
            var matches = items.Where(item => item.environment == environment
                && String.Equals(item.institutionId, institutionId, StringComparison.Ordinal)).Take(2).ToList();
            if (matches.Count > 1)
                throw new InvalidOperationException("Multiple Plaid Items match this institution in the active environment; explicit resolution is required.");
            return matches.SingleOrDefault();
        }

        internal sealed class PlaidRequestException(string message) : Exception(message);

        internal static Account GetLinkedAccount(PlaidItem item, string accountType)
        {
            var account = item.Account;
            if (!item._account_Id.HasValue || account is null || account.Id != item._account_Id.Value)
                throw new PlaidRequestException("Plaid Item account foreign key is missing or invalid.");
            if (!String.Equals(account.name, accountType, StringComparison.OrdinalIgnoreCase)
                && !account.name.StartsWith(accountType + "_", StringComparison.OrdinalIgnoreCase))
                throw new PlaidRequestException("Plaid Item is linked to an unexpected account type.");
            return account;
        }

        // Shared transport for provider modules. Never include response bodies or tokens in errors.
        private async Task<JObject> PostAsync(string path, PlaidItem item, JObject? arguments, CancellationToken cancellationToken)
        {
            var body = arguments ?? new JObject();
            body["client_id"] = clientId;
            body["secret"] = secret;
            body["access_token"] = item.accessToken;
            using var content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
            try
            {
                using var response = await httpClient.PostAsync(apiBaseUrl + path, content, cancellationToken).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new PlaidRequestException($"Plaid POST {path}: " + FormatPlaidError(response,
                        await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false)));
                var json = ParseResponse(await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false));
                // Sync does not return an Item; its caller verifies identity with /item/get and /accounts/get.
                if (path != "/transactions/sync")
                {
                    var institution = ReadItemInstitution(json, item.itemId);
                    if (!String.IsNullOrWhiteSpace(item.institutionId) && institution.Id != item.institutionId)
                        throw new PlaidRequestException($"Plaid POST {path}: institution identity changed.");
                    if (json["item"]!["error"]?.Type is not (null or JTokenType.Null))
                        throw new PlaidRequestException($"Plaid POST {path}: Item reports a connection error; repair is required.");
                }
                return json;
            }
            catch (JsonException) { throw new PlaidRequestException($"Plaid POST {path}: invalid JSON."); }
            catch (HttpRequestException e)
            {
                var cause = e.GetBaseException();
                var detail = cause is System.Net.Sockets.SocketException socket
                    ? socket.SocketErrorCode.ToString() : $"{cause.GetType().Name}/0x{cause.HResult:X8}";
                throw new PlaidRequestException($"Plaid POST {path}: HTTP {e.StatusCode?.ToString() ?? "unavailable"}; category={e.HttpRequestError}; cause={detail}");
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            { throw new PlaidRequestException($"Plaid POST {path}: HTTP unavailable; request timeout ({RequestTimeout.TotalSeconds:0}s)."); }
        }

        internal static JObject ParseResponse(string text)
        {
            using var reader = new JsonTextReader(new StringReader(text))
            { FloatParseHandling = FloatParseHandling.Decimal, DateParseHandling = DateParseHandling.None, MaxDepth = 40 };
            var result = JObject.Load(reader, new JsonLoadSettings { DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Error });
            if (reader.Read()) throw new JsonException("Trailing JSON content.");
            return result;
        }

        internal sealed record InvestmentData(JObject Holdings, List<JObject> Transactions, List<JObject> Securities);

        private static string RawHash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));

        internal sealed record TransactionSyncData(JObject Accounts, List<JObject> Pages);

        internal async Task<TransactionSyncData> GetTransactionUpdatesAsync(PlaidItem item, string? cursor, CancellationToken cancellationToken)
        {
            var status = await PostAsync("/item/get", item, null, cancellationToken).ConfigureAwait(false);
            if (status["item"]?["billed_products"] is not JArray products || !products.Values<string>().Contains("transactions")
                || status["status"]?["transactions"]?["last_successful_update"]?.Type != JTokenType.String)
                throw new PlaidRequestException("Plaid POST /item/get: Transactions is not initialized; explicit authorization is required.");
            var accounts = await PostAsync("/accounts/get", item, null, cancellationToken).ConfigureAwait(false);
            var pages = new List<JObject>();
            var seenCursors = new HashSet<string>(StringComparer.Ordinal);
            if (cursor is not null) seenCursors.Add(cursor);
            while (true)
            {
                var arguments = new JObject { ["count"] = 500,
                    ["options"] = new JObject { ["include_original_description"] = true } };
                if (cursor is not null) arguments["cursor"] = cursor;
                var page = await PostAsync("/transactions/sync", item, arguments, cancellationToken).ConfigureAwait(false);
                if (page["has_more"]?.Type != JTokenType.Boolean)
                    throw new PlaidRequestException("Plaid POST /transactions/sync: missing pagination flag.");
                pages.Add(page);
                cursor = Text(page, "next_cursor");
                if (!page["has_more"]!.Value<bool>()) break;
                if (pages.Count >= 100 || !seenCursors.Add(cursor))
                    throw new PlaidRequestException("Plaid POST /transactions/sync: pagination failed to advance or exceeded limit.");
            }
            if (Text(pages[^1], "transactions_update_status") != "HISTORICAL_UPDATE_COMPLETE")
                throw new PlaidRequestException("Plaid POST /transactions/sync: historical transactions are not ready.");
            // Never commit a partial cursor. A failed page causes the next run to restart from the stored cursor.
            var verification = await PostAsync("/accounts/get", item, null, cancellationToken).ConfigureAwait(false);
            JArray Ordered(JObject response) => new(Rows(response, "accounts").OrderBy(a => Text(a, "account_id"), StringComparer.Ordinal));
            if (!JToken.DeepEquals(Ordered(accounts), Ordered(verification)))
                throw new PlaidRequestException("Plaid POST /accounts/get: accounts changed during sync; retry on the next run.");
            return new(accounts, pages);
        }

        internal async Task<InvestmentData> GetInvestmentsAsync(PlaidItem item, DateTime start, DateTime end, CancellationToken cancellationToken)
        {
            if (start > end) throw new InvalidOperationException("Plaid investment date range is invalid.");
            var holdings = await PostAsync("/investments/holdings/get", item, null, cancellationToken).ConfigureAwait(false);
            var accounts = Rows(holdings, "accounts").Select(a => Text(a, "account_id")).ToHashSet(StringComparer.Ordinal);
            var transactions = new List<JObject>();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var securities = Rows(holdings, "securities").ToList();
            int? total = null;
            do
            {
                var page = await PostAsync("/investments/transactions/get", item, new JObject
                {
                    ["start_date"] = start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    ["end_date"] = end.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    ["options"] = new JObject { ["count"] = 500, ["offset"] = transactions.Count }
                }, cancellationToken).ConfigureAwait(false);
                CheckFields(page, "accounts investment_transactions item request_id securities total_investment_transactions");
                if (page["total_investment_transactions"]?.Type != JTokenType.Integer)
                    throw new InvalidOperationException("Plaid transaction total is missing or invalid.");
                var count = page["total_investment_transactions"]!.Value<int>();
                if (count < 0 || count > 50000 || total.HasValue && total != count)
                    throw new InvalidOperationException("Plaid pagination total changed or exceeds the supported limit.");
                total = count;
                if (!accounts.SetEquals(Rows(page, "accounts").Select(a => Text(a, "account_id"))))
                    throw new InvalidOperationException("Plaid account selection changed during pagination.");
                var rows = Rows(page, "investment_transactions").ToList();
                if (rows.Count > 500 || rows.Count == 0 && transactions.Count < total)
                    throw new InvalidOperationException("Plaid transaction page is incomplete.");
                foreach (var row in rows)
                {
                    if (!ids.Add(Text(row, "investment_transaction_id")) || !accounts.Contains(Text(row, "account_id")))
                        throw new InvalidOperationException("Plaid transaction page contains duplicate or unselected entries.");
                    transactions.Add(row);
                }
                securities.AddRange(Rows(page, "securities"));
                if (transactions.Count > total) throw new InvalidOperationException("Plaid pagination exceeds its declared total.");
            } while (transactions.Count < total);
            // The API does not offer an atomic snapshot. Reject a moving holdings response instead of mixing states.
            var verification = await PostAsync("/investments/holdings/get", item, null, cancellationToken).ConfigureAwait(false);
            foreach (var field in new[] { "accounts", "holdings", "securities" })
                if (!JToken.DeepEquals(holdings[field], verification[field]))
                    throw new InvalidOperationException("Plaid investment snapshot changed during retrieval; retry on the next run.");
            return new(holdings, transactions, securities);
        }

        private static IEnumerable<JObject> Rows(JObject value, string field)
        {
            if (value[field] is not JArray rows) throw new InvalidOperationException("Plaid missing array: " + field);
            foreach (var row in rows)
                yield return row as JObject ?? throw new InvalidOperationException("Plaid invalid row in " + field);
        }

        private static string Text(JToken row, string field) => row[field]?.Type == JTokenType.String
            && !String.IsNullOrWhiteSpace(row[field]!.Value<string>()) ? row[field]!.Value<string>()!
            : throw new InvalidOperationException("Plaid missing text: " + field);
        private static string OptionalText(JToken row, string field) => row[field]?.Type is null or JTokenType.Null ? "" : Text(row, field);
        private static decimal Number(JToken row, string field)
        {
            if (row[field]?.Type is not (JTokenType.Integer or JTokenType.Float))
                throw new InvalidOperationException("Plaid invalid number: " + field);
            var value = row[field]!.Value<decimal>();
            MySqlDecimalColumnTypes.ValidateCurrencyValue(value, "Plaid " + field);
            return value;
        }
        private static DateTime Date(JToken row, string field) => DateTime.TryParseExact(Text(row, field), "yyyy-MM-dd",
            CultureInfo.InvariantCulture, DateTimeStyles.None, out var value) ? value
            : throw new InvalidOperationException("Plaid invalid date: " + field);
        private static void RequireUsd(JToken row)
        {
            if (Text(row, "iso_currency_code") != "USD" || OptionalText(row, "unofficial_currency_code") != "")
                throw new InvalidOperationException("Plaid Schwab requires explicit USD values; unsupported currency.");
        }

        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient(new SocketsHttpHandler
            {
                AllowAutoRedirect = false,
                UseCookies = false,
                PooledConnectionLifetime = TimeSpan.FromMinutes(5)
            })
            {
                Timeout = RequestTimeout
            };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 PlaidInstitutionVerifier");
            return client;
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
                var errorCode = json["error_code"]?.Type == JTokenType.String ? json["error_code"]!.Value<string>() : null;
                var category = errorCode switch
                {
                    "ITEM_LOGIN_REQUIRED" or "ITEM_LOCKED" or "INVALID_ACCESS_TOKEN" or "INVALID_API_KEYS"
                        or "PRODUCT_NOT_READY" or "PRODUCT_NOT_SUPPORTED" or "INSTITUTION_DOWN"
                        or "INSTITUTION_NOT_RESPONDING" or "RATE_LIMIT_EXCEEDED"
                        or "TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION" => errorCode,
                    _ => "unrecognized_error"
                };
                return $"HTTP {(int)response.StatusCode}; category={category}";
            }
            catch (JsonException)
            {
                return $"HTTP {(int)response.StatusCode}; category=non_json_error";
            }
        }
    }
}
