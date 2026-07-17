using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net.Http;
using System.Numerics;
using System.Text;

namespace MyBook
{
    partial class CryptoUtil
    {
        private const string BlockscoutApiBaseUrl = "https://eth.blockscout.com/api/v2";
        private const string EthereumRpcUrl = "https://ethereum-rpc.publicnode.com";
        private const string AddressAccountPrefix = "ETH_";
        private const string OfficialUsdtContract = "0xdac17f958d2ee523a2206206994597c13d831ec7";
        private const int EthDecimals = 18;
        private const int UsdtDecimals = 6;
        private static readonly TimeSpan ETHImportRequestTimeout = TimeSpan.FromSeconds(20);

        private async Task FetchETHDailyReportsAsync(CancellationToken cancellationToken = default)
        {
            var accounts = database.GetAccountsByNamePrefix(AddressAccountPrefix);
            foreach (var account in accounts)
                await FetchAccountDailyReportsAsync(account, cancellationToken).ConfigureAwait(false);
        }

        private async Task FetchAccountDailyReportsAsync(Account account, CancellationToken cancellationToken)
        {
            var address = account.name[AddressAccountPrefix.Length..].ToLowerInvariant();
            ValidateEthereumAddressValue(address);
            var checkpoint = database.GetLatestStatementImportTimeByKeyPrefix(StatementImportProvider.EthereumApi, address + ":")
                ?? database.GetStatementImportCheckpointTime(StatementImportProvider.EthereumApi)
                ?? throw new InvalidOperationException("Missing Ethereum statement import checkpoint.");
            var firstDate = checkpoint.Date.AddDays(1);
            var lastCompletedDate = DateTime.UtcNow.Date.AddDays(-1);
            if (firstDate > lastCompletedDate)
                return;

            var events = await FetchEventsAsync(address, cancellationToken).ConfigureAwait(false);
            var currentQuantities = new Dictionary<string, decimal>(StringComparer.Ordinal)
            {
                ["ETH"] = await FetchEthBalanceWeiAsync(address, cancellationToken).ConfigureAwait(false),
                ["USDT"] = await FetchUsdtBalanceRawAsync(address, cancellationToken).ConfigureAwait(false)
            };
            ValidateCurrentQuantities(events, currentQuantities);

            var quantities = new Dictionary<string, decimal>(StringComparer.Ordinal)
            {
                ["ETH"] = 0,
                ["USDT"] = 0
            };
            var imports = new List<StatementRecordHoldingImport>();
            for (var date = firstDate; date <= lastCompletedDate; date = date.AddDays(1))
            {
                var dayEnd = DateTime.SpecifyKind(date.AddDays(1), DateTimeKind.Utc);
                var dayEvents = events.Where(item => item.Time >= date && item.Time < dayEnd).ToList();
                var beginningQuantities = new Dictionary<string, decimal>(quantities, StringComparer.Ordinal);
                foreach (var item in dayEvents)
                    quantities[item.Asset] += item.QuantityRaw;

                var beginningHoldings = CreateHoldings(account, beginningQuantities);
                var endingHoldings = CreateHoldings(account, quantities);
                var records = CreateRecords(account, dayEvents);
                imports.Add(new StatementRecordHoldingImport(
                    StatementImportProvider.EthereumApi,
                    date,
                    $"{address}:{date:yyyyMMdd}",
                    account,
                    records,
                    endingHoldings,
                    CreateAccountBalances(account, quantities),
                    CreateAccountBalances(account, beginningQuantities),
                    beginningHoldings,
                    recordDate: date));
            }

            var completedQuantities = new Dictionary<string, decimal>(currentQuantities, StringComparer.Ordinal);
            foreach (var item in events.Where(item => item.Time >= DateTime.SpecifyKind(lastCompletedDate.AddDays(1), DateTimeKind.Utc)))
                completedQuantities[item.Asset] -= item.QuantityRaw;
            ValidateQuantities(quantities, completedQuantities, $"Ethereum completed balance {lastCompletedDate:yyyy-MM-dd}");
            var saved = database.SaveStatementRecordsAndHoldingsOnce(imports);
            Console.WriteLine($"Fetch Ethereum daily reports done: account={account.name}; events={events.Count}; saved={saved.Count(value => value)}");
        }

        private static List<Record> CreateRecords(Account account, List<EthereumAssetEvent> events)
        {
            return events.Select(item =>
            {
                var quantity = ToAssetQuantity(item.QuantityRaw, item.Decimals);
                var record = CreateRecord(account, item.Asset, item.Time, quantity, item.Reason, item.Source);
                record.blockchain = BlockchainType.Ethereum;
                record.blockchainTransactionHash = item.TransactionHash;
                record.blockchainEventIndex = item.EventIndex;
                record.blockchainAssetContract = item.ContractAddress;
                return record;
            }).ToList();
        }

        private static Record CreateRecord(Account account, string asset, DateTime date, decimal amount, string reason, string source)
        {
            var record = new Record
            {
                Account = account,
                Holding = CreateHolding(account, asset, 0),
                date = date,
                postingDate = date,
                updateTime = DateTime.Now,
                DestAccount = asset,
                Reason = reason,
                Source = source
            };
            record.CopyFrom(new Currency(amount, ParseAssetCurrency(asset)));
            return record;
        }

        private static List<Holding> CreateHoldings(Account account, Dictionary<string, decimal> quantities)
        {
            return quantities.Where(item => item.Value != 0)
                .Select(item => CreateHolding(account, item.Key, ToAssetQuantity(
                    item.Value, item.Key == "ETH" ? EthDecimals : UsdtDecimals)))
                .ToList();
        }

        private static Holding CreateHolding(Account account, string asset, decimal value)
        {
            return new Holding(asset, HoldingType.Crypto)
            {
                Account = account,
                desc = $"Ethereum {asset}",
                displayText = asset,
                currentPrice = new Currency(value, ParseAssetCurrency(asset))
            };
        }

        private static List<AccountBalance> CreateAccountBalances(Account account, Dictionary<string, decimal> rawQuantities)
        {
            return rawQuantities
                .Select(item => new AccountBalance(account, new Currency(
                    ToAssetQuantity(item.Value, item.Key == "ETH" ? EthDecimals : UsdtDecimals),
                    ParseAssetCurrency(item.Key))))
                .ToList();
        }

        private static CurrencyType ParseAssetCurrency(string asset)
        {
            return Enum.Parse<CurrencyType>(asset);
        }

        private static async Task<List<EthereumAssetEvent>> FetchEventsAsync(string address, CancellationToken cancellationToken)
        {
            var normalTask = FetchBlockscoutItemsAsync($"addresses/{address}/transactions", cancellationToken);
            var internalTask = FetchBlockscoutItemsAsync($"addresses/{address}/internal-transactions", cancellationToken);
            var tokenTask = FetchBlockscoutItemsAsync($"addresses/{address}/token-transfers", cancellationToken);
            await Task.WhenAll(normalTask, internalTask, tokenTask).ConfigureAwait(false);
            var result = new List<EthereumAssetEvent>();

            foreach (var item in await normalTask.ConfigureAwait(false))
            {
                if (!String.Equals(item["status"]?.ToString(), "ok", StringComparison.OrdinalIgnoreCase))
                    continue;
                var hash = RequiredText(item, "hash").ToLowerInvariant();
                var time = DateTime.Parse(RequiredText(item, "timestamp"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal);
                var from = item["from"]?["hash"]?.ToString() ?? "";
                var to = item["to"]?["hash"]?.ToString() ?? "";
                var value = ReadRaw(item["value"]);
                if (value != 0 && (AddressEquals(from, address) || AddressEquals(to, address)))
                {
                    var signed = AddressEquals(from, address) ? -value : value;
                    result.Add(new EthereumAssetEvent(hash, -1, time, "ETH", "", EthDecimals, signed,
                        "\u94fe\u4e0a\u8f6c\u8d26", $"Ethereum transaction {hash}; from={from}; to={to}; valueWei={value}"));
                }
                if (AddressEquals(from, address))
                {
                    var fee = ReadRaw(item["fee"]?["value"]);
                    if (fee != 0)
                        result.Add(new EthereumAssetEvent(hash, -2, time, "ETH", "", EthDecimals, -fee,
                            "\u624b\u7eed\u8d39", $"Ethereum gas {hash}; feeWei={fee}"));
                }
            }

            foreach (var item in await internalTask.ConfigureAwait(false))
            {
                if (item["success"]?.Value<bool>() == false)
                    continue;
                var from = item["from"]?["hash"]?.ToString() ?? "";
                var to = item["to"]?["hash"]?.ToString() ?? "";
                if (!AddressEquals(from, address) && !AddressEquals(to, address))
                    continue;
                var value = ReadRaw(item["value"]);
                if (value == 0)
                    continue;
                var hash = RequiredText(item, "transaction_hash").ToLowerInvariant();
                var index = item["index"]?.Value<int>() ?? throw new InvalidOperationException($"Ethereum internal transaction {hash} has no index.");
                var time = DateTime.Parse(RequiredText(item, "timestamp"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal);
                result.Add(new EthereumAssetEvent(hash, index, time, "ETH", "", EthDecimals,
                    AddressEquals(from, address) ? -value : value,
                    "\u94fe\u4e0a\u5185\u90e8\u8f6c\u8d26", $"Ethereum internal transaction {hash}; index={index}; from={from}; to={to}; valueWei={value}"));
            }

            foreach (var item in await tokenTask.ConfigureAwait(false))
            {
                var contract = item["token"]?["address_hash"]?.ToString().ToLowerInvariant() ?? "";
                if (contract != OfficialUsdtContract)
                    continue;
                var from = item["from"]?["hash"]?.ToString() ?? "";
                var to = item["to"]?["hash"]?.ToString() ?? "";
                if (!AddressEquals(from, address) && !AddressEquals(to, address))
                    continue;
                var value = ReadRaw(item["total"]?["value"]);
                if (value == 0)
                    continue;
                var hash = RequiredText(item, "transaction_hash").ToLowerInvariant();
                var index = item["log_index"]?.Value<int>() ?? throw new InvalidOperationException($"Ethereum token transfer {hash} has no log index.");
                var time = DateTime.Parse(RequiredText(item, "timestamp"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal);
                result.Add(new EthereumAssetEvent(hash, index, time, "USDT", contract, UsdtDecimals,
                    AddressEquals(from, address) ? -value : value,
                    "\u94fe\u4e0a\u8f6c\u8d26", $"Ethereum ERC20 transfer {hash}; logIndex={index}; contract={contract}; from={from}; to={to}; valueRaw={value}"));
            }
            return result.OrderBy(item => item.Time).ThenBy(item => item.EventIndex).ToList();
        }

        private static async Task<List<JObject>> FetchBlockscoutItemsAsync(string path, CancellationToken cancellationToken)
        {
            var result = new List<JObject>();
            string? query = null;
            do
            {
                using var client = CreateHttpClient();
                using var response = await client.GetAsync($"{BlockscoutApiBaseUrl}/{path}{query}", cancellationToken).ConfigureAwait(false);
                var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException($"Blockscout request failed: {(int)response.StatusCode} {response.ReasonPhrase}.");
                var json = JObject.Parse(text);
                result.AddRange((json["items"] as JArray ?? new JArray()).OfType<JObject>());
                var next = json["next_page_params"] as JObject;
                query = next is null ? null : "?" + String.Join("&", next.Properties().Select(property =>
                    $"{Uri.EscapeDataString(property.Name)}={Uri.EscapeDataString(property.Value.ToString())}"));
            } while (query is not null);
            return result;
        }

        private static async Task<decimal> FetchEthBalanceWeiAsync(string address, CancellationToken cancellationToken)
        {
            var result = await PostRpcAsync("eth_getBalance", new JArray(address, "latest"), cancellationToken).ConfigureAwait(false);
            return ParseHexRaw(result.ToString());
        }

        private static async Task<decimal> FetchUsdtBalanceRawAsync(string address, CancellationToken cancellationToken)
        {
            var data = "0x70a08231" + address[2..].PadLeft(64, '0');
            var result = await PostRpcAsync("eth_call", new JArray(new JObject
            {
                ["to"] = OfficialUsdtContract,
                ["data"] = data
            }, "latest"), cancellationToken).ConfigureAwait(false);
            return ParseHexRaw(result.ToString());
        }

        private static async Task<JToken> PostRpcAsync(string method, JArray parameters, CancellationToken cancellationToken)
        {
            using var client = CreateHttpClient();
            using var content = new StringContent(new JObject
            {
                ["jsonrpc"] = "2.0",
                ["method"] = method,
                ["params"] = parameters,
                ["id"] = 1
            }.ToString(), Encoding.UTF8, "application/json");
            using var response = await client.PostAsync(EthereumRpcUrl, content, cancellationToken).ConfigureAwait(false);
            var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException($"Ethereum RPC failed: {(int)response.StatusCode} {response.ReasonPhrase}.");
            var json = JObject.Parse(text);
            if (json["error"] is not null)
                throw new InvalidOperationException($"Ethereum RPC failed: {json["error"]}.");
            return json["result"] ?? throw new InvalidOperationException("Ethereum RPC response has no result.");
        }

        private static void ValidateCurrentQuantities(List<EthereumAssetEvent> events, Dictionary<string, decimal> current)
        {
            var calculated = new Dictionary<string, decimal>(StringComparer.Ordinal) { ["ETH"] = 0, ["USDT"] = 0 };
            foreach (var item in events)
                calculated[item.Asset] += item.QuantityRaw;
            ValidateQuantities(calculated, current, "Ethereum current balance");
        }

        private static void ValidateQuantities(Dictionary<string, decimal> expected, Dictionary<string, decimal> actual, string context)
        {
            foreach (var asset in expected.Keys.Union(actual.Keys, StringComparer.Ordinal))
            {
                var left = expected.TryGetValue(asset, out var expectedValue) ? expectedValue : 0;
                var right = actual.TryGetValue(asset, out var actualValue) ? actualValue : 0;
                if (left != right)
                    throw new InvalidOperationException($"{context} mismatch: asset={asset}, calculatedRaw={left}, actualRaw={right}.");
            }
        }

        private static decimal ReadRaw(JToken? token) => decimal.Parse(token?.ToString() ?? "0", CultureInfo.InvariantCulture);
        private static decimal ParseHexRaw(string value) => decimal.Parse(BigInteger.Parse("0" + value[2..], NumberStyles.AllowHexSpecifier).ToString(), CultureInfo.InvariantCulture);
        private static decimal ToAssetQuantity(decimal raw, int decimals) => raw / Pow10(decimals);
        private static decimal Pow10(int decimals)
        {
            var value = 1m;
            for (var index = 0; index < decimals; index++)
                value *= 10m;
            return value;
        }
        private static bool AddressEquals(string left, string right) => String.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        private static string RequiredText(JObject item, string name) => item[name]?.ToString() ?? throw new InvalidOperationException($"Ethereum API item has no {name}.");
        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient { Timeout = ETHImportRequestTimeout };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 CryptoUtil.ETH");
            return client;
        }
    }

    sealed record EthereumAssetEvent(
        string TransactionHash,
        int EventIndex,
        DateTime Time,
        string Asset,
        string ContractAddress,
        int Decimals,
        decimal QuantityRaw,
        string Reason,
        string Source);
}
