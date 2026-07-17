using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net.Http;
using System.Text.RegularExpressions;

namespace MyBook
{
    partial class CryptoUtil
    {
        private const string ApiBaseUrl = "https://api.etherscan.io/v2/api";
        private const int ChainId = 1;
        private const int PageSize = 1000;
        private static readonly TimeSpan EtherscanRequestTimeout = TimeSpan.FromSeconds(20);
        private static readonly Regex AddressPattern = new("^0x[0-9a-fA-F]{40}$", RegexOptions.CultureInvariant);

        public async Task<EthereumAddressData> FetchETHAddressDataAsync(
            string address,
            CancellationToken cancellationToken = default)
        {
            ValidateEthereumAddressValue(address);
            var normalizedAddress = address.ToLowerInvariant();
            var balanceTask = FetchNativeBalanceWeiAsync(normalizedAddress, cancellationToken);
            var transactionsTask = FetchPagedAsync("txlist", normalizedAddress, cancellationToken);
            var internalTransactionsTask = FetchPagedAsync("txlistinternal", normalizedAddress, cancellationToken);
            var tokenTransfersTask = FetchPagedAsync("tokentx", normalizedAddress, cancellationToken);
            await Task.WhenAll(balanceTask, transactionsTask, internalTransactionsTask, tokenTransfersTask)
                .ConfigureAwait(false);

            return new EthereumAddressData(
                normalizedAddress,
                await balanceTask.ConfigureAwait(false),
                await transactionsTask.ConfigureAwait(false),
                await internalTransactionsTask.ConfigureAwait(false),
                await tokenTransfersTask.ConfigureAwait(false));
        }

        public async Task<string> FetchNativeBalanceWeiAsync(
            string address,
            CancellationToken cancellationToken = default)
        {
            ValidateEthereumAddressValue(address);
            var result = await GetResultAsync(new Dictionary<string, string>
            {
                ["module"] = "account",
                ["action"] = "balance",
                ["address"] = address,
                ["tag"] = "latest"
            }, cancellationToken).ConfigureAwait(false);
            return result.Type == JTokenType.String
                ? result.Value<string>()!
                : throw new InvalidOperationException("Etherscan balance response is not an integer string.");
        }

        public async Task<string> FetchTokenBalanceRawAsync(
            string address,
            string contractAddress,
            CancellationToken cancellationToken = default)
        {
            ValidateEthereumAddressValue(address);
            ValidateEthereumAddressValue(contractAddress);
            var result = await GetResultAsync(new Dictionary<string, string>
            {
                ["module"] = "account",
                ["action"] = "tokenbalance",
                ["address"] = address,
                ["contractaddress"] = contractAddress,
                ["tag"] = "latest"
            }, cancellationToken).ConfigureAwait(false);
            return result.Type == JTokenType.String
                ? result.Value<string>()!
                : throw new InvalidOperationException("Etherscan token balance response is not an integer string.");
        }

        private async Task<List<JObject>> FetchPagedAsync(
            string action,
            string address,
            CancellationToken cancellationToken)
        {
            var result = new List<JObject>();
            for (var page = 1; ; page++)
            {
                var token = await GetResultAsync(new Dictionary<string, string>
                {
                    ["module"] = "account",
                    ["action"] = action,
                    ["address"] = address,
                    ["startblock"] = "0",
                    ["endblock"] = "9999999999",
                    ["page"] = page.ToString(CultureInfo.InvariantCulture),
                    ["offset"] = PageSize.ToString(CultureInfo.InvariantCulture),
                    ["sort"] = "asc"
                }, cancellationToken, allowNoTransactions: true).ConfigureAwait(false);
                var pageItems = token is JArray array
                    ? array.OfType<JObject>().ToList()
                    : throw new InvalidOperationException($"Etherscan {action} response is not an array.");
                result.AddRange(pageItems);
                if (pageItems.Count < PageSize)
                    return result;
            }
        }

        private async Task<JToken> GetResultAsync(
            Dictionary<string, string> parameters,
            CancellationToken cancellationToken,
            bool allowNoTransactions = false)
        {
            parameters["chainid"] = ChainId.ToString(CultureInfo.InvariantCulture);
            parameters["apikey"] = EtherscanApiKey;
            var query = String.Join("&", parameters.Select(item =>
                $"{Uri.EscapeDataString(item.Key)}={Uri.EscapeDataString(item.Value)}"));

            using var client = new HttpClient { Timeout = EtherscanRequestTimeout };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 EthereumReader");
            using var response = await client.GetAsync($"{ApiBaseUrl}?{query}", cancellationToken).ConfigureAwait(false);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException($"Etherscan request failed: {(int)response.StatusCode} {response.ReasonPhrase}.");

            JObject json;
            try
            {
                json = JObject.Parse(responseText);
            }
            catch (JsonException exception)
            {
                throw new InvalidOperationException("Etherscan returned invalid JSON.", exception);
            }

            var status = json["status"]?.ToString();
            var message = json["message"]?.ToString();
            if (status == "0" && allowNoTransactions && String.Equals(message, "No transactions found", StringComparison.OrdinalIgnoreCase))
                return new JArray();
            if (status != "1")
                throw new InvalidOperationException($"Etherscan request failed: {message}; result={json["result"]}.");
            return json["result"] ?? throw new InvalidOperationException("Etherscan response has no result.");
        }

        internal static void ValidateEthereumAddressValue(string address)
        {
            if (String.IsNullOrWhiteSpace(address) || !AddressPattern.IsMatch(address))
                throw new ArgumentException("Ethereum address must contain 0x followed by 40 hexadecimal characters.", nameof(address));
        }

        private static string RequiredConfig(IConfigurationRoot config, string key)
        {
            var value = config[key];
            return String.IsNullOrWhiteSpace(value)
                ? throw new InvalidOperationException($"Missing {key} in config.json.")
                : value.Trim();
        }
    }

    sealed record EthereumAddressData(
        string Address,
        string NativeBalanceWei,
        List<JObject> Transactions,
        List<JObject> InternalTransactions,
        List<JObject> TokenTransfers);
}
