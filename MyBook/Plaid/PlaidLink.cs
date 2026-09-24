using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Net.Http;
using System.Text;

namespace MyBook
{
    // Creates and repairs Plaid Items through Hosted Link. Access tokens are stored in
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

        public async Task<IReadOnlyList<PlaidAccount>> GetAllStoredAccountsAsync(
            CancellationToken cancellationToken = default)
        {
            var state = LoadTokenState();
            if (state.Items.Count == 0)
                throw new PlaidLinkException(
                    "POST /accounts/get: no Plaid Items are stored in the database; complete Link first");

            var accounts = new List<PlaidAccount>();
            foreach (var item in state.Items)
            {
                if (String.IsNullOrWhiteSpace(item.ItemId) || String.IsNullOrWhiteSpace(item.AccessToken))
                    throw new PlaidLinkException(
                        "POST /accounts/get: database contains an incomplete Plaid Item");
                var response = await PostAsync(
                    "/accounts/get",
                    new { client_id = clientId, secret, access_token = item.AccessToken },
                    cancellationToken).ConfigureAwait(false);
                if (response["accounts"] is not JArray itemAccounts)
                    throw new PlaidLinkException(
                        "POST /accounts/get: successful response has no accounts array");
                foreach (var account in itemAccounts.OfType<JObject>())
                {
                    accounts.Add(new PlaidAccount(
                        item.ItemId,
                        RequiredText(account, "account_id", "/accounts/get"),
                        item.InstitutionId,
                        item.InstitutionName,
                        account["name"]?.ToString(),
                        account["official_name"]?.ToString(),
                        account["type"]?.ToString(),
                        account["subtype"]?.ToString(),
                        account["balances"]?["iso_currency_code"]?.ToString()
                            ?? account["balances"]?["unofficial_currency_code"]?.ToString(),
                        account["balances"]?["current"]?.Value<decimal?>(),
                        account["balances"]?["available"]?.Value<decimal?>(),
                        account["balances"]?["limit"]?.Value<decimal?>()));
                }
            }
            return accounts;
        }

        public async Task<PlaidTransactionSummary> GetTransactionSummaryAsync(
            int plaidItemDatabaseId,
            DateOnly startDate,
            DateOnly endDate,
            CancellationToken cancellationToken = default)
        {
            if (startDate > endDate)
                throw new ArgumentException("Transaction start date is after the end date.");
            var item = LoadTokenState().Items
                .SingleOrDefault(item => item.DatabaseId == plaidItemDatabaseId)
                ?? throw new PlaidLinkException(
                    "POST /transactions/get: selected database Plaid Item was not found in the active environment");
            if (String.IsNullOrWhiteSpace(item.AccessToken))
                throw new PlaidLinkException(
                    "POST /transactions/get: selected database Plaid Item has no access_token");

            const int pageSize = 500;
            var offset = 0;
            int? expectedTotal = null;
            int? accountCount = null;
            DateOnly? earliestTransactionDate = null;
            DateOnly? latestTransactionDate = null;
            var seenTransactionIds = new HashSet<string>(StringComparer.Ordinal);
            var totals = new Dictionary<string, MutablePlaidTransactionCurrencySummary>(StringComparer.Ordinal);
            var accountsById = new Dictionary<string, MutablePlaidTransactionAccountSummary>(StringComparer.Ordinal);
            var namedCounterpartyTransactions = 0;
            var paymentPartyTransactions = 0;
            var counterpartyAccountNumberTransactions = 0;
            var counterpartyIbanTransactions = 0;
            var counterpartyBicTransactions = 0;
            var counterpartyBacsTransactions = 0;
            while (true)
            {
                var response = await PostAsync(
                    "/transactions/get",
                    new
                    {
                        client_id = clientId,
                        secret,
                        access_token = item.AccessToken,
                        start_date = startDate.ToString("yyyy-MM-dd"),
                        end_date = endDate.ToString("yyyy-MM-dd"),
                        options = new { count = pageSize, offset }
                    },
                    cancellationToken).ConfigureAwait(false);
                if (response["accounts"] is not JArray accounts)
                    throw new PlaidLinkException(
                        "POST /transactions/get: successful response has no accounts array");
                if (response["transactions"] is not JArray transactions)
                    throw new PlaidLinkException(
                        "POST /transactions/get: successful response has no transactions array");
                var responseTotal = response["total_transactions"]?.Value<int?>()
                    ?? throw new PlaidLinkException(
                        "POST /transactions/get: successful response has no total_transactions");
                if (responseTotal < 0)
                    throw new PlaidLinkException(
                        "POST /transactions/get: total_transactions is negative");
                if (expectedTotal.HasValue && expectedTotal.Value != responseTotal)
                    throw new PlaidLinkException(
                        "POST /transactions/get: total_transactions changed during pagination");
                expectedTotal ??= responseTotal;
                accountCount ??= accounts.Count;
                if (accountsById.Count == 0)
                {
                    var accountOrdinal = 0;
                    foreach (var account in accounts.OfType<JObject>())
                    {
                        accountOrdinal++;
                        var accountId = RequiredText(account, "account_id", "/transactions/get");
                        var balances = account["balances"] as JObject
                            ?? throw new PlaidLinkException(
                                "POST /transactions/get: account has no balances object");
                        var accountCurrency = balances["iso_currency_code"]?.ToString()
                            ?? balances["unofficial_currency_code"]?.ToString();
                        if (String.IsNullOrWhiteSpace(accountCurrency))
                            throw new PlaidLinkException(
                                "POST /transactions/get: account balance has no currency code");
                        if (!accountsById.TryAdd(
                            accountId,
                            new MutablePlaidTransactionAccountSummary(
                                accountOrdinal,
                                accountCurrency,
                                balances["current"]?.Value<decimal?>())))
                        {
                            throw new PlaidLinkException(
                                "POST /transactions/get: duplicate account in response");
                        }
                    }
                    if (accountsById.Count != accounts.Count)
                        throw new PlaidLinkException(
                            "POST /transactions/get: account array contains an invalid entry");
                }

                foreach (var transaction in transactions.OfType<JObject>())
                {
                    var transactionId = RequiredText(transaction, "transaction_id", "/transactions/get");
                    if (!seenTransactionIds.Add(transactionId))
                        throw new PlaidLinkException(
                            "POST /transactions/get: duplicate transaction encountered during pagination");
                    var dateText = RequiredText(transaction, "date", "/transactions/get");
                    if (!DateOnly.TryParseExact(
                        dateText,
                        "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out var date)
                        || date < startDate
                        || date > endDate)
                    {
                        throw new PlaidLinkException(
                            "POST /transactions/get: transaction date is invalid or outside the requested range");
                    }
                    var amount = transaction["amount"]?.Value<decimal?>()
                        ?? throw new PlaidLinkException(
                            "POST /transactions/get: transaction has no amount");
                    var currency = transaction["iso_currency_code"]?.ToString()
                        ?? transaction["unofficial_currency_code"]?.ToString();
                    if (String.IsNullOrWhiteSpace(currency))
                        throw new PlaidLinkException(
                            "POST /transactions/get: transaction has no currency code");
                    var pending = transaction["pending"]?.Value<bool?>()
                        ?? throw new PlaidLinkException(
                            "POST /transactions/get: transaction has no pending status");
                    earliestTransactionDate = !earliestTransactionDate.HasValue || date < earliestTransactionDate.Value
                        ? date
                        : earliestTransactionDate;
                    latestTransactionDate = !latestTransactionDate.HasValue || date > latestTransactionDate.Value
                        ? date
                        : latestTransactionDate;

                    var accountId = RequiredText(transaction, "account_id", "/transactions/get");
                    if (!accountsById.TryGetValue(accountId, out var accountSummary))
                        throw new PlaidLinkException(
                            "POST /transactions/get: transaction references an account absent from the response");
                    if (!String.Equals(accountSummary.Currency, currency, StringComparison.Ordinal))
                        throw new PlaidLinkException(
                            "POST /transactions/get: transaction currency differs from its account balance currency");

                    if (!totals.TryGetValue(currency, out var currencySummary))
                    {
                        currencySummary = new MutablePlaidTransactionCurrencySummary();
                        totals.Add(currency, currencySummary);
                    }
                    if (pending)
                    {
                        currencySummary.PendingCount++;
                        currencySummary.PendingNetChange -= amount;
                        accountSummary.PendingCount++;
                        accountSummary.PendingNetChange -= amount;
                    }
                    else
                    {
                        currencySummary.PostedCount++;
                        accountSummary.PostedCount++;
                        if (amount < 0)
                        {
                            currencySummary.Inflow += -amount;
                            accountSummary.Inflow += -amount;
                        }
                        else
                        {
                            currencySummary.Outflow += amount;
                            accountSummary.Outflow += amount;
                        }
                    }

                    var counterparties = transaction["counterparties"] as JArray;
                    var hasNamedCounterparty = counterparties?.OfType<JObject>()
                        .Any(counterparty => !String.IsNullOrWhiteSpace(counterparty["name"]?.ToString())) == true;
                    var hasCounterpartyAccountNumber = false;
                    var hasCounterpartyIban = false;
                    var hasCounterpartyBic = false;
                    var hasCounterpartyBacs = false;
                    foreach (var counterparty in counterparties?.OfType<JObject>() ?? [])
                    {
                        if (counterparty["account_numbers"] is not JObject accountNumbers)
                            continue;
                        if (accountNumbers["international"] is JObject international)
                        {
                            hasCounterpartyIban |= !String.IsNullOrWhiteSpace(international["iban"]?.ToString());
                            hasCounterpartyBic |= !String.IsNullOrWhiteSpace(international["bic"]?.ToString());
                        }
                        if (accountNumbers["bacs"] is JObject bacs)
                        {
                            hasCounterpartyBacs |= !String.IsNullOrWhiteSpace(bacs["account"]?.ToString())
                                || !String.IsNullOrWhiteSpace(bacs["sort_code"]?.ToString());
                        }
                        hasCounterpartyAccountNumber |= hasCounterpartyIban || hasCounterpartyBic || hasCounterpartyBacs;
                    }
                    var paymentMeta = transaction["payment_meta"] as JObject;
                    var hasPaymentParty = paymentMeta is not null
                        && new[] { "payee", "payer", "by_order_of" }
                            .Any(field => !String.IsNullOrWhiteSpace(paymentMeta[field]?.ToString()));
                    if (hasNamedCounterparty)
                        namedCounterpartyTransactions++;
                    if (hasPaymentParty)
                        paymentPartyTransactions++;
                    if (hasCounterpartyAccountNumber)
                        counterpartyAccountNumberTransactions++;
                    if (hasCounterpartyIban)
                        counterpartyIbanTransactions++;
                    if (hasCounterpartyBic)
                        counterpartyBicTransactions++;
                    if (hasCounterpartyBacs)
                        counterpartyBacsTransactions++;
                }

                offset += transactions.Count;
                if (offset >= responseTotal)
                    break;
                if (transactions.Count == 0)
                    throw new PlaidLinkException(
                        "POST /transactions/get: pagination returned an empty page before total_transactions");
            }

            return new PlaidTransactionSummary(
                plaidItemDatabaseId,
                startDate,
                endDate,
                accountCount ?? 0,
                expectedTotal ?? 0,
                earliestTransactionDate,
                latestTransactionDate,
                totals
                    .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                    .Select(pair => new PlaidTransactionCurrencySummary(
                        pair.Key,
                        pair.Value.PostedCount,
                        pair.Value.PendingCount,
                        pair.Value.Inflow,
                        pair.Value.Outflow,
                        pair.Value.Inflow - pair.Value.Outflow,
                        pair.Value.PendingNetChange))
                    .ToList(),
                accountsById.Values
                    .OrderBy(account => account.Ordinal)
                    .Select(account => new PlaidTransactionAccountSummary(
                        account.Ordinal,
                        account.Currency,
                        account.CurrentBalance,
                        account.PostedCount,
                        account.PendingCount,
                        account.Inflow,
                        account.Outflow,
                        account.CurrentBalance.HasValue
                            ? account.CurrentBalance.Value - account.Inflow + account.Outflow
                            : null,
                        account.PendingNetChange))
                    .ToList(),
                new PlaidTransactionCounterpartyCoverage(
                    namedCounterpartyTransactions,
                    paymentPartyTransactions,
                    counterpartyAccountNumberTransactions,
                    counterpartyIbanTransactions,
                    counterpartyBicTransactions,
                    counterpartyBacsTransactions));
        }

        public async Task RepairStoredItemAsync(
            int plaidItemDatabaseId,
            string clientUserId,
            CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(clientUserId))
                throw new ArgumentException("Plaid client user id is empty.", nameof(clientUserId));
            var item = LoadTokenState().Items
                .SingleOrDefault(item => item.DatabaseId == plaidItemDatabaseId)
                ?? throw new PlaidLinkException(
                    "start Link update mode: selected database Plaid Item was not found in the active environment");
            if (String.IsNullOrWhiteSpace(item.AccessToken))
                throw new PlaidLinkException(
                    "start Link update mode: stored Item has no access_token");

            var createResponse = await PostAsync(
                "/link/token/create",
                new
                {
                    client_id = clientId,
                    secret,
                    client_name = ClientName,
                    language = Language,
                    country_codes = new[] { "US" },
                    user = new { client_user_id = clientUserId.Trim() },
                    access_token = item.AccessToken,
                    hosted_link = new { url_lifetime_seconds = HostedLinkLifetimeSeconds }
                },
                cancellationToken).ConfigureAwait(false);
            var linkToken = RequiredText(createResponse, "link_token", "/link/token/create");
            var hostedLinkUrl = RequiredText(createResponse, "hosted_link_url", "/link/token/create");
            OpenHostedLink(hostedLinkUrl);

            using var completionTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            completionTimeout.CancelAfter(LinkCompletionTimeout);
            try
            {
                await WaitForUpdateCompletionAsync(linkToken, completionTimeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                throw new PlaidLinkException(
                    $"poll POST /link/token/get for update mode: no completed Link session within {LinkCompletionTimeout.TotalMinutes:0} minutes");
            }

            var itemResponse = await PostAsync(
                "/item/get",
                new { client_id = clientId, secret, access_token = item.AccessToken },
                cancellationToken).ConfigureAwait(false);
            if (itemResponse["item"]?["error"] is JObject itemError)
                throw new PlaidLinkException(
                    "verify POST /item/get after update mode: Item remains in error; "
                    + $"error_type={itemError["error_type"]}; error_code={itemError["error_code"]}");
        }

        public async Task<PlaidInvestmentSummary> GetOnlyStoredItemInvestmentSummaryAsync(
            DateOnly startDate,
            DateOnly endDate,
            CancellationToken cancellationToken = default)
        {
            if (startDate > endDate)
                throw new ArgumentException("Investment transaction start date is after the end date.");
            var state = LoadTokenState();
            if (state.Items.Count != 1)
                throw new PlaidLinkException(
                    $"read Schwab investment data: expected exactly one stored Item; found {state.Items.Count}");
            var item = state.Items[0];
            if (String.IsNullOrWhiteSpace(item.AccessToken))
                throw new PlaidLinkException(
                    "read Schwab investment data: stored Item has no access_token");

            var itemResponse = await PostAsync(
                "/item/get",
                new { client_id = clientId, secret, access_token = item.AccessToken },
                cancellationToken).ConfigureAwait(false);
            var itemJson = itemResponse["item"] as JObject
                ?? throw new PlaidLinkException("POST /item/get: successful response has no item object");
            if (itemJson["error"] is JObject itemError)
                throw new PlaidLinkException(
                    "POST /item/get: Item is in error; "
                    + $"error_type={itemError["error_type"]}; error_code={itemError["error_code"]}");

            var accounts = await GetAllStoredAccountsAsync(cancellationToken).ConfigureAwait(false);
            var holdingsResponse = await PostAsync(
                "/investments/holdings/get",
                new { client_id = clientId, secret, access_token = item.AccessToken },
                cancellationToken).ConfigureAwait(false);
            var holdings = holdingsResponse["holdings"] as JArray
                ?? throw new PlaidLinkException(
                    "POST /investments/holdings/get: successful response has no holdings array");
            var securities = holdingsResponse["securities"] as JArray
                ?? throw new PlaidLinkException(
                    "POST /investments/holdings/get: successful response has no securities array");

            var transactionTypes = new Dictionary<string, int>(StringComparer.Ordinal);
            var transactionSubtypes = new Dictionary<string, int>(StringComparer.Ordinal);
            var transactionCurrencies = new Dictionary<string, int>(StringComparer.Ordinal);
            var transactionIds = new HashSet<string>(StringComparer.Ordinal);
            var offset = 0;
            int? expectedTotal = null;
            while (true)
            {
                var transactionsResponse = await PostAsync(
                    "/investments/transactions/get",
                    new
                    {
                        client_id = clientId,
                        secret,
                        access_token = item.AccessToken,
                        start_date = startDate.ToString("yyyy-MM-dd"),
                        end_date = endDate.ToString("yyyy-MM-dd"),
                        options = new { count = 500, offset }
                    },
                    cancellationToken).ConfigureAwait(false);
                var transactions = transactionsResponse["investment_transactions"] as JArray
                    ?? throw new PlaidLinkException(
                        "POST /investments/transactions/get: successful response has no investment_transactions array");
                var pageTotal = transactionsResponse["total_investment_transactions"]?.Value<int>()
                    ?? throw new PlaidLinkException(
                        "POST /investments/transactions/get: successful response has no total_investment_transactions");
                if (expectedTotal.HasValue && expectedTotal.Value != pageTotal)
                    throw new PlaidLinkException(
                        "POST /investments/transactions/get: total changed during pagination");
                expectedTotal = pageTotal;

                foreach (var transaction in transactions.OfType<JObject>())
                {
                    var transactionId = RequiredText(
                        transaction,
                        "investment_transaction_id",
                        "/investments/transactions/get");
                    if (!transactionIds.Add(transactionId))
                        throw new PlaidLinkException(
                            "POST /investments/transactions/get: duplicate transaction across pages");
                    AddCount(transactionTypes, transaction["type"]?.ToString());
                    AddCount(transactionSubtypes, transaction["subtype"]?.ToString());
                    AddCount(
                        transactionCurrencies,
                        transaction["iso_currency_code"]?.ToString()
                            ?? transaction["unofficial_currency_code"]?.ToString());
                }

                offset += transactions.Count;
                if (offset >= pageTotal)
                    break;
                if (transactions.Count == 0)
                    throw new PlaidLinkException(
                        "POST /investments/transactions/get: empty page before reported total");
            }

            return new PlaidInvestmentSummary(
                item.InstitutionName,
                accounts.Count,
                holdings.Count,
                securities.Count,
                CountValues(securities, "type"),
                CountCurrencies(holdings),
                transactionIds.Count,
                ToCounts(transactionTypes),
                ToCounts(transactionSubtypes),
                ToCounts(transactionCurrencies),
                ReadStrings(itemJson["products"]),
                ReadStrings(itemJson["consented_products"]),
                itemJson["update_type"]?.ToString(),
                itemJson["auth_method"]?.ToString(),
                startDate,
                endDate);
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

        private async Task WaitForUpdateCompletionAsync(
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
                        var finishedAt = session["finished_at"];
                        if (finishedAt is null
                            || finishedAt.Type is JTokenType.Null or JTokenType.Undefined)
                            continue;

                        if (session["exit"] is JObject exit)
                        {
                            var errorCode = exit["error"]?["error_code"]?.ToString();
                            var errorType = exit["error"]?["error_type"]?.ToString();
                            var exitStatus = exit["metadata"]?["status"]?.ToString();
                            throw new PlaidLinkException(
                                "poll POST /link/token/get for update mode: Link session exited; "
                                + $"error_type={errorType}; error_code={errorCode}; exit_status={exitStatus}");
                        }

                        return;
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
                throw new PlaidLinkException(
                    "persist Plaid Item: database insert failed after public-token exchange",
                    exception);
            }
        }

        private PlaidTokenState LoadTokenState()
        {
            try
            {
                var items = database.GetPlaidItems(environment)
                    .Select(item => new PlaidStoredItem
                    {
                        DatabaseId = item.Id,
                        ItemId = item.itemId,
                        AccessToken = item.accessToken,
                        InstitutionId = item.institutionId,
                        InstitutionName = item.institutionName,
                        CreatedAtUtc = new DateTimeOffset(DateTime.SpecifyKind(item.createdAtUtc, DateTimeKind.Utc))
                    })
                    .ToList();
                return new PlaidTokenState { Items = items };
            }
            catch (Exception exception)
            {
                throw new PlaidLinkException(
                    "load Plaid Item tokens: database query failed",
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

        private static List<string> ReadStrings(JToken? token) => token is JArray array
            ? array.Values<string>()
                .Where(value => !String.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToList()
            : [];

        private static List<PlaidValueCount> CountValues(JArray values, string propertyName)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var value in values.OfType<JObject>())
                AddCount(counts, value[propertyName]?.ToString());
            return ToCounts(counts);
        }

        private static List<PlaidValueCount> CountCurrencies(JArray holdings)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var holding in holdings.OfType<JObject>())
            {
                AddCount(
                    counts,
                    holding["iso_currency_code"]?.ToString()
                        ?? holding["unofficial_currency_code"]?.ToString());
            }
            return ToCounts(counts);
        }

        private static void AddCount(Dictionary<string, int> counts, string? value)
        {
            var key = String.IsNullOrWhiteSpace(value) ? "null" : value;
            counts[key] = counts.TryGetValue(key, out var count) ? count + 1 : 1;
        }

        private static List<PlaidValueCount> ToCounts(Dictionary<string, int> counts) => counts
            .OrderBy(pair => pair.Key, StringComparer.Ordinal)
            .Select(pair => new PlaidValueCount(pair.Key, pair.Value))
            .ToList();

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

    public sealed record PlaidAccount(
        string ItemId,
        string AccountId,
        string? InstitutionId,
        string? InstitutionName,
        string? Name,
        string? OfficialName,
        string? Type,
        string? Subtype,
        string? Currency,
        decimal? CurrentBalance,
        decimal? AvailableBalance,
        decimal? Limit);

    public sealed record PlaidInvestmentSummary(
        string? InstitutionName,
        int AccountCount,
        int HoldingCount,
        int SecurityCount,
        IReadOnlyList<PlaidValueCount> SecurityTypes,
        IReadOnlyList<PlaidValueCount> HoldingCurrencies,
        int InvestmentTransactionCount,
        IReadOnlyList<PlaidValueCount> TransactionTypes,
        IReadOnlyList<PlaidValueCount> TransactionSubtypes,
        IReadOnlyList<PlaidValueCount> TransactionCurrencies,
        IReadOnlyList<string> Products,
        IReadOnlyList<string> ConsentedProducts,
        string? UpdateType,
        string? AuthMethod,
        DateOnly StartDate,
        DateOnly EndDate);

    public sealed record PlaidValueCount(string Value, int Count);

    public sealed record PlaidTransactionSummary(
        int DatabaseId,
        DateOnly StartDate,
        DateOnly EndDate,
        int AccountCount,
        int TransactionCount,
        DateOnly? EarliestTransactionDate,
        DateOnly? LatestTransactionDate,
        IReadOnlyList<PlaidTransactionCurrencySummary> Currencies,
        IReadOnlyList<PlaidTransactionAccountSummary> Accounts,
        PlaidTransactionCounterpartyCoverage CounterpartyCoverage);

    public sealed record PlaidTransactionCurrencySummary(
        string Currency,
        int PostedCount,
        int PendingCount,
        decimal Inflow,
        decimal Outflow,
        decimal NetChange,
        decimal PendingNetChange);

    public sealed record PlaidTransactionAccountSummary(
        int Ordinal,
        string Currency,
        decimal? CurrentBalance,
        int PostedCount,
        int PendingCount,
        decimal Inflow,
        decimal Outflow,
        decimal? OpeningBalanceEstimate,
        decimal PendingNetChange);

    public sealed record PlaidTransactionCounterpartyCoverage(
        int NamedCounterpartyTransactions,
        int PaymentPartyTransactions,
        int AccountNumberTransactions,
        int IbanTransactions,
        int BicTransactions,
        int BacsTransactions);

    sealed record PlaidPublicToken(
        string PublicToken,
        string? InstitutionId,
        string? InstitutionName);

    sealed class PlaidTokenState
    {
        public List<PlaidStoredItem> Items { get; set; } = [];
    }

    sealed class PlaidStoredItem
    {
        public int DatabaseId { get; set; }
        public string ItemId { get; set; } = "";
        public string AccessToken { get; set; } = "";
        public string? InstitutionId { get; set; }
        public string? InstitutionName { get; set; }
        public DateTimeOffset CreatedAtUtc { get; set; }
    }

    sealed class MutablePlaidTransactionCurrencySummary
    {
        public int PostedCount { get; set; }
        public int PendingCount { get; set; }
        public decimal Inflow { get; set; }
        public decimal Outflow { get; set; }
        public decimal PendingNetChange { get; set; }
    }

    sealed class MutablePlaidTransactionAccountSummary
    {
        public MutablePlaidTransactionAccountSummary(int ordinal, string currency, decimal? currentBalance)
        {
            Ordinal = ordinal;
            Currency = currency;
            CurrentBalance = currentBalance;
        }

        public int Ordinal { get; }
        public string Currency { get; }
        public decimal? CurrentBalance { get; }
        public int PostedCount { get; set; }
        public int PendingCount { get; set; }
        public decimal Inflow { get; set; }
        public decimal Outflow { get; set; }
        public decimal PendingNetChange { get; set; }
    }

    public sealed class PlaidLinkException : Exception
    {
        public PlaidLinkException(string message) : base(message) { }
        public PlaidLinkException(string message, Exception innerException) : base(message, innerException) { }
    }
}
