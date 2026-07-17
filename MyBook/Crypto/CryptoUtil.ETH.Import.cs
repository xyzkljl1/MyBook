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

            var firstDateUtc = DateTime.SpecifyKind(firstDate, DateTimeKind.Utc);
            var quantities = GetImportedRawQuantities(account);
            var events = await FetchEventsAsync(address, firstDateUtc, cancellationToken).ConfigureAwait(false);
            var currentQuantities = new Dictionary<string, decimal>(StringComparer.Ordinal)
            {
                ["ETH"] = await FetchEthBalanceWeiAsync(address, cancellationToken).ConfigureAwait(false),
                ["USDT"] = await FetchUsdtBalanceRawAsync(address, cancellationToken).ConfigureAwait(false)
            };
            ValidateCurrentQuantities(quantities, events, currentQuantities);
            var prices = await cryptoPrice.FetchDailyUsdPricesAsync(
                ["ETH", "USDT"],
                firstDate.AddDays(-1),
                lastCompletedDate,
                cancellationToken).ConfigureAwait(false);

            var imports = new List<StatementRecordHoldingImport>();
            for (var date = firstDate; date <= lastCompletedDate; date = date.AddDays(1))
            {
                var dayEnd = DateTime.SpecifyKind(date.AddDays(1), DateTimeKind.Utc);
                var dayEvents = events.Where(item => item.Time >= date && item.Time < dayEnd).ToList();
                var beginningQuantities = new Dictionary<string, decimal>(quantities, StringComparer.Ordinal);
                foreach (var item in dayEvents)
                    quantities[item.Asset] += item.QuantityRaw;

                var beginningAssetQuantities = ToAssetQuantities(beginningQuantities);
                var endingAssetQuantities = ToAssetQuantities(quantities);
                var beginningHoldings = CreateHoldings(account, beginningAssetQuantities, prices, date.AddDays(-1));
                var endingHoldings = CreateHoldings(account, endingAssetQuantities, prices, date);
                var records = CreateRecords(account, dayEvents, prices);
                AddValuationRecords(
                    records,
                    account,
                    date,
                    beginningAssetQuantities,
                    endingAssetQuantities,
                    prices,
                    "Ethereum");
                imports.Add(new StatementRecordHoldingImport(
                    StatementImportProvider.EthereumApi,
                    date,
                    $"{address}:{date:yyyyMMdd}",
                    account,
                    records,
                    endingHoldings,
                    CreateAccountBalances(account, endingHoldings),
                    CreateAccountBalances(account, beginningHoldings),
                    beginningHoldings,
                    recordDate: date));
            }

            var completedQuantities = new Dictionary<string, decimal>(currentQuantities, StringComparer.Ordinal);
            foreach (var item in events.Where(item => item.Time >= DateTime.SpecifyKind(lastCompletedDate.AddDays(1), DateTimeKind.Utc)))
                completedQuantities[item.Asset] -= item.QuantityRaw;
            ValidateQuantities(quantities, completedQuantities, $"Ethereum completed balance {lastCompletedDate:yyyy-MM-dd}");
            var saved = database.SaveStatementRecordsAndHoldingsOnce(
                imports,
                CreateLatestPrices(prices, ["ETH", "USDT"], lastCompletedDate));
            Console.WriteLine($"Fetch Ethereum daily reports done: account={account.name}; events={events.Count}; saved={saved.Count(value => value)}");
        }

        private static List<Record> CreateRecords(
            Account account,
            List<EthereumAssetEvent> events,
            CryptoPriceSet prices)
        {
            return events.Select(item =>
            {
                var quantity = ToAssetQuantity(item.QuantityRaw, item.Decimals);
                var price = prices.Get(item.Asset, item.Time.Date);
                var amount = CryptoPriceUtil.CalculateMarketValue(quantity, price.CloseUsd);
                var record = CreateRecord(
                    account,
                    item.Asset,
                    item.Time,
                    amount,
                    quantity,
                    item.Reason,
                    $"{item.Source}; valuation=Kraken {price.Asset}/USD; closeDate={price.SourceCandleDate:yyyy-MM-dd}; close={price.CloseUsd}");
                record.blockchain = BlockchainType.Ethereum;
                record.blockchainTransactionHash = item.TransactionHash;
                record.blockchainEventIndex = item.EventIndex;
                record.blockchainAssetContract = item.ContractAddress;
                return record;
            }).ToList();
        }

        private static Record CreateRecord(
            Account account,
            string asset,
            DateTime date,
            decimal amount,
            decimal holdingQuantity,
            string reason,
            string source)
        {
            var record = new Record
            {
                Account = account,
                Holding = CreateHolding(account, asset, 0, 0),
                date = date,
                postingDate = date,
                updateTime = DateTime.Now,
                HoldingQuantity = holdingQuantity,
                DestAccount = asset,
                Reason = reason,
                Source = source
            };
            record.CopyFrom(new Currency(amount, CurrencyType.USD));
            return record;
        }

        private static Dictionary<string, decimal> ToAssetQuantities(Dictionary<string, decimal> rawQuantities)
        {
            return rawQuantities.ToDictionary(
                item => item.Key,
                item => ToAssetQuantity(item.Value, item.Key == "ETH" ? EthDecimals : UsdtDecimals),
                StringComparer.Ordinal);
        }

        private static List<Holding> CreateHoldings(
            Account account,
            Dictionary<string, decimal> quantities,
            CryptoPriceSet prices,
            DateTime priceDate)
        {
            return quantities.Where(item => item.Value != 0)
                .Select(item => CreateHolding(account, item.Key, item.Value, prices.Get(item.Key, priceDate).CloseUsd))
                .ToList();
        }

        private static Holding CreateHolding(Account account, string asset, decimal quantity, decimal unitPrice)
        {
            return new Holding(asset, HoldingType.Crypto)
            {
                Account = account,
                desc = $"Ethereum {asset}",
                displayText = asset,
                quantity = quantity,
                currentPrice = new Currency(unitPrice, CurrencyType.USD)
            };
        }

        private static List<AccountBalance> CreateAccountBalances(Account account, List<Holding> holdings)
        {
            return
            [
                new AccountBalance(
                    account,
                    new Currency(Currency.RoundMoney(holdings.Sum(holding => holding.totalPrice.v)), CurrencyType.USD))
            ];
        }

        private static void AddValuationRecords(
            List<Record> records,
            Account account,
            DateTime date,
            Dictionary<string, decimal> beginningQuantities,
            Dictionary<string, decimal> endingQuantities,
            CryptoPriceSet prices,
            string sourcePrefix)
        {
            foreach (var asset in beginningQuantities.Keys.Union(endingQuantities.Keys, StringComparer.Ordinal).OrderBy(asset => asset, StringComparer.Ordinal))
            {
                var beginningQuantity = beginningQuantities.TryGetValue(asset, out var beginning) ? beginning : 0;
                var endingQuantity = endingQuantities.TryGetValue(asset, out var ending) ? ending : 0;
                var previousPrice = prices.Get(asset, date.AddDays(-1));
                var currentPrice = prices.Get(asset, date);
                var beginningValue = CryptoPriceUtil.CalculateMarketValue(beginningQuantity, previousPrice.CloseUsd);
                var repricedBeginningValue = CryptoPriceUtil.CalculateMarketValue(beginningQuantity, currentPrice.CloseUsd);
                var priceChange = repricedBeginningValue - beginningValue;
                if (priceChange != 0)
                {
                    records.Add(CreateRecord(
                        account,
                        asset,
                        date,
                        priceChange,
                        0,
                        "持仓价格变动",
                        $"{sourcePrefix} valuation; asset={asset}; quantity={beginningQuantity}; previousClose={previousPrice.CloseUsd}; currentClose={currentPrice.CloseUsd}; previousDate={previousPrice.SourceCandleDate:yyyy-MM-dd}; currentDate={currentPrice.SourceCandleDate:yyyy-MM-dd}"));
                }

                var eventValue = records
                    .Where(record => record.Holding?.code == asset && record.HoldingQuantity != 0)
                    .Sum(record => record.v);
                var endingValue = CryptoPriceUtil.CalculateMarketValue(endingQuantity, currentPrice.CloseUsd);
                var transactionPriceImpact = endingValue - repricedBeginningValue - eventValue;
                if (transactionPriceImpact != 0)
                {
                    records.Add(CreateRecord(
                        account,
                        asset,
                        date,
                        transactionPriceImpact,
                        0,
                        "交易价格影响",
                        $"{sourcePrefix} valuation; asset={asset}; endingValue={endingValue}; repricedBeginning={repricedBeginningValue}; eventValue={eventValue}; close={currentPrice.CloseUsd}; closeDate={currentPrice.SourceCandleDate:yyyy-MM-dd}"));
                }
            }
        }

        private static List<Finance> CreateLatestPrices(CryptoPriceSet prices, IEnumerable<string> assets, DateTime date)
        {
            return assets.Distinct(StringComparer.Ordinal).Select(asset =>
            {
                var price = prices.Get(asset, date);
                return new Finance(CryptoPriceUtil.GetBaseAsset(asset), HoldingType.Crypto)
                {
                    currentPrice = new Currency(price.CloseUsd, CurrencyType.USD),
                    currentPriceTime = new DateTimeOffset(DateTime.SpecifyKind(date.AddDays(1), DateTimeKind.Utc)).ToUnixTimeSeconds()
                };
            }).ToList();
        }

        private async Task<List<EthereumAssetEvent>> FetchEventsAsync(
            string address,
            DateTime fromUtc,
            CancellationToken cancellationToken)
        {
            var normalTask = FetchBlockscoutItemsAsync($"addresses/{address}/transactions", fromUtc, cancellationToken);
            var internalTask = FetchBlockscoutItemsAsync($"addresses/{address}/internal-transactions", fromUtc, cancellationToken);
            var tokenTask = FetchBlockscoutItemsAsync($"addresses/{address}/token-transfers", fromUtc, cancellationToken);
            await Task.WhenAll(normalTask, internalTask, tokenTask).ConfigureAwait(false);
            var result = new List<EthereumAssetEvent>();

            foreach (var item in await normalTask.ConfigureAwait(false))
                result.AddRange(ParseNormalTransactionEvents(item, address));

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

        private static List<EthereumAssetEvent> ParseNormalTransactionEvents(JObject item, string address)
        {
            var status = item["status"]?.ToString() ?? "";
            var succeeded = String.Equals(status, "ok", StringComparison.OrdinalIgnoreCase);
            var failed = String.Equals(status, "error", StringComparison.OrdinalIgnoreCase);
            if (!succeeded && !failed)
                return [];

            var result = new List<EthereumAssetEvent>();
            var hash = RequiredText(item, "hash").ToLowerInvariant();
            var time = DateTime.Parse(RequiredText(item, "timestamp"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal);
            var from = item["from"]?["hash"]?.ToString() ?? "";
            var to = item["to"]?["hash"]?.ToString() ?? "";
            var value = ReadRaw(item["value"]);
            if (succeeded && value != 0 && (AddressEquals(from, address) || AddressEquals(to, address)))
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
                        "\u624b\u7eed\u8d39", $"Ethereum gas {hash}; status={status}; feeWei={fee}"));
            }
            return result;
        }

        private static Dictionary<string, decimal> CreateEmptyRawQuantities()
        {
            return new Dictionary<string, decimal>(StringComparer.Ordinal)
            {
                ["ETH"] = 0,
                ["USDT"] = 0
            };
        }

        private Dictionary<string, decimal> GetImportedRawQuantities(Account account)
        {
            var result = CreateEmptyRawQuantities();
            foreach (var holding in database.GetCurrentAccountHoldings(account))
            {
                if (holding.holdingType != HoldingType.Crypto)
                {
                    if (holding.totalPrice.v != 0)
                        throw new InvalidOperationException($"Unsupported Ethereum holding: {holding.code}/{holding.holdingType}.");
                    continue;
                }

                if (holding.currentPrice.t != CurrencyType.USD)
                    throw new InvalidOperationException($"Ethereum crypto holding must be valued in USD: {holding.code}/{holding.currentPrice.t}.");
                result[holding.code] = holding.code switch
                {
                    "ETH" => ToRawQuantity(holding.quantity, EthDecimals),
                    "USDT" => ToRawQuantity(holding.quantity, UsdtDecimals),
                    _ => throw new InvalidOperationException($"Unsupported Ethereum crypto holding: {holding.code}.")
                };
            }
            return result;
        }

        private async Task<List<JObject>> FetchBlockscoutItemsAsync(
            string path,
            DateTime fromUtc,
            CancellationToken cancellationToken)
        {
            var result = new List<JObject>();
            string? query = null;
            DateTime? previousTime = null;
            do
            {
                using var response = await sharedHttpClient.GetAsync($"{BlockscoutApiBaseUrl}/{path}{query}", cancellationToken).ConfigureAwait(false);
                var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException($"Blockscout request failed: {(int)response.StatusCode} {response.ReasonPhrase}.");
                var json = JObject.Parse(text);
                var items = (json["items"] as JArray ?? new JArray()).OfType<JObject>().ToList();
                var reachedEarlierEvent = false;
                foreach (var item in items)
                {
                    var time = DateTime.Parse(RequiredText(item, "timestamp"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal);
                    if (previousTime.HasValue && time > previousTime.Value)
                        throw new InvalidOperationException($"Blockscout {path} response is not ordered newest first.");
                    previousTime = time;
                    if (time >= fromUtc)
                        result.Add(item);
                    else
                        reachedEarlierEvent = true;
                }
                if (reachedEarlierEvent)
                    break;

                var next = json["next_page_params"] as JObject;
                query = next is null ? null : "?" + String.Join("&", next.Properties().Select(property =>
                    $"{Uri.EscapeDataString(property.Name)}={Uri.EscapeDataString(property.Value.ToString())}"));
            } while (query is not null);
            return result;
        }

        private async Task<decimal> FetchEthBalanceWeiAsync(string address, CancellationToken cancellationToken)
        {
            var result = await PostRpcAsync("eth_getBalance", new JArray(address, "latest"), cancellationToken).ConfigureAwait(false);
            return ParseHexRaw(result.ToString());
        }

        private async Task<decimal> FetchUsdtBalanceRawAsync(string address, CancellationToken cancellationToken)
        {
            var data = "0x70a08231" + address[2..].PadLeft(64, '0');
            var result = await PostRpcAsync("eth_call", new JArray(new JObject
            {
                ["to"] = OfficialUsdtContract,
                ["data"] = data
            }, "latest"), cancellationToken).ConfigureAwait(false);
            return ParseHexRaw(result.ToString());
        }

        private async Task<JToken> PostRpcAsync(string method, JArray parameters, CancellationToken cancellationToken)
        {
            using var content = new StringContent(new JObject
            {
                ["jsonrpc"] = "2.0",
                ["method"] = method,
                ["params"] = parameters,
                ["id"] = 1
            }.ToString(), Encoding.UTF8, "application/json");
            using var response = await sharedHttpClient.PostAsync(EthereumRpcUrl, content, cancellationToken).ConfigureAwait(false);
            var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException($"Ethereum RPC failed: {(int)response.StatusCode} {response.ReasonPhrase}.");
            var json = JObject.Parse(text);
            if (json["error"] is not null)
                throw new InvalidOperationException($"Ethereum RPC failed: {json["error"]}.");
            return json["result"] ?? throw new InvalidOperationException("Ethereum RPC response has no result.");
        }

        private static void ValidateCurrentQuantities(
            Dictionary<string, decimal> beginning,
            List<EthereumAssetEvent> events,
            Dictionary<string, decimal> current)
        {
            var calculated = new Dictionary<string, decimal>(beginning, StringComparer.Ordinal);
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
        private static decimal ToRawQuantity(decimal quantity, int decimals)
        {
            var raw = quantity * Pow10(decimals);
            if (Decimal.Truncate(raw) != raw)
                throw new InvalidOperationException($"Ethereum quantity has more than {decimals} decimal places: {quantity}.");
            return raw;
        }
        private static decimal Pow10(int decimals)
        {
            var value = 1m;
            for (var index = 0; index < decimals; index++)
                value *= 10m;
            return value;
        }
        private static bool AddressEquals(string left, string right) => String.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        private static string RequiredText(JObject item, string name) => item[name]?.ToString() ?? throw new InvalidOperationException($"Ethereum API item has no {name}.");
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
