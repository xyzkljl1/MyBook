using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;

namespace MyBook
{
    partial class KrakenUtil
    {
        private const string KrakenApiBaseUrl = "https://api.kraken.com";
        private const string EthereumUsdtContract = "0xdac17f958d2ee523a2206206994597c13d831ec7";
        private const int LedgerPageSize = 50;
        private const int FundingPageSize = 50;
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(20);
        private static long lastNonce;

        private readonly string apiKey;
        private readonly byte[] apiSecret;
        private readonly DatabaseUtil? database;
        private readonly SemaphoreSlim requestLock = new(1, 1);

        public KrakenUtil(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            apiKey = RequiredConfig(config, "kraken_api_key");
            var secret = RequiredConfig(config, "kraken_api_secret");
            try
            {
                apiSecret = Convert.FromBase64String(secret);
            }
            catch (FormatException exception)
            {
                throw new InvalidOperationException("kraken_api_secret is not valid Base64.", exception);
            }
            this.database = database;
        }

        public async Task FetchDailyReportsAsync(CancellationToken cancellationToken = default)
        {
            var db = database ?? throw new InvalidOperationException("FetchDailyReportsAsync requires a database.");
            var account = db.GetAccountByName("KRAKEN");
            var latestReportDate = db.GetLatestStatementImportTime(StatementImportProvider.KrakenApi);
            if (!latestReportDate.HasValue)
                throw new InvalidOperationException("Missing Kraken statement import checkpoint.");

            var firstDate = latestReportDate.Value.Date.AddDays(1);
            var lastCompletedUtcDate = DateTime.UtcNow.Date.AddDays(-1);
            if (firstDate > lastCompletedUtcDate)
            {
                Console.WriteLine($"Fetch Kraken daily reports: no completed UTC day after {latestReportDate:yyyy-MM-dd}");
                return;
            }

            await ValidateApiPermissionsAsync(cancellationToken).ConfigureAwait(false);

            var rangeStartUtc = DateTime.SpecifyKind(firstDate, DateTimeKind.Utc);
            var rangeEndUtc = DateTime.SpecifyKind(lastCompletedUtcDate.AddDays(1), DateTimeKind.Utc);
            var currentBalances = (await FetchExtendedBalancesAsync(cancellationToken).ConfigureAwait(false))
                .ToDictionary(balance => balance.Asset, balance => balance.Balance, StringComparer.Ordinal);
            var ledgers = await FetchLedgersAsync(rangeStartUtc, DateTime.UtcNow, cancellationToken).ConfigureAwait(false);
            var fundingTransactions = new List<KrakenFundingTransaction>();
            foreach (var asset in ledgers.Where(ledger => ledger.Type is "deposit" or "withdrawal")
                         .Select(ledger => NormalizeAsset(ledger.Asset))
                         .Where(asset => asset is "ETH" or "USDT")
                         .Distinct(StringComparer.Ordinal))
            {
                fundingTransactions.AddRange(await FetchFundingTransactionsAsync(
                    "/0/private/DepositStatus",
                    "deposits",
                    asset,
                    false,
                    DateTime.UnixEpoch,
                    DateTime.UtcNow,
                    cancellationToken).ConfigureAwait(false));
                if (ledgers.Any(ledger => ledger.Type == "withdrawal" && NormalizeAsset(ledger.Asset) == asset))
                {
                    fundingTransactions.AddRange(await FetchFundingTransactionsAsync(
                        "/0/private/WithdrawStatus",
                        "withdrawals",
                        asset,
                        true,
                        DateTime.UnixEpoch,
                        DateTime.UtcNow,
                        cancellationToken).ConfigureAwait(false));
                }
            }

            var previousQuantities = DeriveQuantitiesAt(currentBalances, ledgers, rangeStartUtc);
            ValidateQuantityChainToCurrent(previousQuantities, ledgers, rangeStartUtc, currentBalances);

            var imports = new List<StatementRecordHoldingImport>();
            for (var date = firstDate; date <= lastCompletedUtcDate; date = date.AddDays(1))
            {
                var dayStartUtc = DateTime.SpecifyKind(date, DateTimeKind.Utc);
                var dayEndUtc = dayStartUtc.AddDays(1);
                var dayLedgers = ledgers.Where(ledger => ledger.Time >= dayStartUtc && ledger.Time < dayEndUtc).ToList();
                var endingQuantities = new Dictionary<string, decimal>(previousQuantities, StringComparer.Ordinal);
                foreach (var ledger in dayLedgers)
                    endingQuantities[ledger.Asset] = GetQuantity(endingQuantities, ledger.Asset) + ledger.Amount - ledger.Fee;

                var beginningHoldings = CreateHoldings(account, previousQuantities);
                var endingHoldings = CreateHoldings(account, endingQuantities);
                var records = CreateDailyRecords(account, date, dayLedgers, fundingTransactions);
                imports.Add(new StatementRecordHoldingImport(
                        StatementImportProvider.KrakenApi,
                        date,
                        $"daily-{date:yyyyMMdd}",
                        account,
                        records,
                        endingHoldings,
                        CreateAccountBalances(account, endingQuantities),
                        CreateAccountBalances(account, previousQuantities),
                        beginningHoldings,
                        recordDate: date));

                previousQuantities = endingQuantities;
            }

            var quantitiesAtCompletedEnd = DeriveQuantitiesAt(currentBalances, ledgers, rangeEndUtc);
            ValidateQuantitiesEqual(previousQuantities, quantitiesAtCompletedEnd, $"Kraken ending balance {lastCompletedUtcDate:yyyy-MM-dd}");
            var saved = db.SaveStatementRecordsAndHoldingsOnce(imports);
            Console.WriteLine(
                $"Fetch Kraken daily reports done: range={firstDate:yyyy-MM-dd}..{lastCompletedUtcDate:yyyy-MM-dd}, "
                + $"ledgers={ledgers.Count}, saved={saved.Count(value => value)}, skipped={saved.Count(value => !value)}");
        }

        public async Task<List<KrakenExtendedBalance>> FetchExtendedBalancesAsync(
            CancellationToken cancellationToken = default)
        {
            var result = await PostPrivateAsync(
                "/0/private/BalanceEx",
                new Dictionary<string, string>(),
                cancellationToken).ConfigureAwait(false);
            return result.Properties()
                .Select(property => ParseExtendedBalance(property.Name, property.Value))
                .OrderBy(balance => balance.Asset, StringComparer.Ordinal)
                .ToList();
        }

        public async Task<HashSet<string>> FetchApiPermissionsAsync(CancellationToken cancellationToken = default)
        {
            var result = await PostPrivateAsync(
                "/0/private/GetApiKeyInfo",
                new Dictionary<string, string>(),
                cancellationToken).ConfigureAwait(false);
            return (result["permissions"] as JArray)?.Values<string>()
                .Where(permission => !String.IsNullOrWhiteSpace(permission))
                .Select(permission => permission!)
                .ToHashSet(StringComparer.Ordinal)
                ?? [];
        }

        public async Task<List<KrakenFundingTransaction>> FetchRecentDepositsAsync(
            string asset,
            CancellationToken cancellationToken = default)
        {
            return await FetchFundingTransactionsAsync(
                "/0/private/DepositStatus",
                "deposits",
                asset,
                false,
                null,
                null,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<List<KrakenFundingTransaction>> FetchRecentWithdrawalsAsync(
            string asset,
            CancellationToken cancellationToken = default)
        {
            return await FetchFundingTransactionsAsync(
                "/0/private/WithdrawStatus",
                "withdrawals",
                asset,
                true,
                null,
                null,
                cancellationToken).ConfigureAwait(false);
        }

        private async Task<List<KrakenFundingTransaction>> FetchFundingTransactionsAsync(
            string path,
            string collectionName,
            string asset,
            bool isWithdrawal,
            DateTime? startUtc,
            DateTime? endUtc,
            CancellationToken cancellationToken)
        {
            var result = new List<KrakenFundingTransaction>();
            var seenCursors = new HashSet<string>(StringComparer.Ordinal);
            string? cursor = "true";
            while (cursor is not null)
            {
                var parameters = new Dictionary<string, string>
                {
                    ["asset"] = asset,
                    ["cursor"] = cursor,
                    ["limit"] = FundingPageSize.ToString(CultureInfo.InvariantCulture)
                };
                if (startUtc.HasValue)
                    parameters["start"] = ToUnixSeconds(startUtc.Value).ToString(CultureInfo.InvariantCulture);
                if (endUtc.HasValue)
                    parameters["end"] = ToUnixSeconds(endUtc.Value).ToString(CultureInfo.InvariantCulture);

                var page = await PostPrivateResultAsync(path, parameters, cancellationToken).ConfigureAwait(false);
                var (rows, nextCursor) = ParseFundingPage(page, collectionName);
                result.AddRange(rows.Select(item => ParseFundingTransaction(item, isWithdrawal)));
                if (String.IsNullOrWhiteSpace(nextCursor))
                {
                    if (page is JArray && rows.Count >= FundingPageSize)
                        throw new InvalidOperationException($"Kraken {path} did not return a cursor for a full page.");
                    break;
                }
                if (!seenCursors.Add(nextCursor))
                    throw new InvalidOperationException($"Kraken {path} returned a repeated pagination cursor.");
                cursor = nextCursor;
            }
            return result;
        }

        private static (List<JObject> Rows, string? NextCursor) ParseFundingPage(JToken page, string collectionName)
        {
            if (page is JArray array)
            {
                var cursor = array.Last?.Type == JTokenType.String ? array.Last.ToString() : null;
                return (array.OfType<JObject>().ToList(), cursor);
            }
            if (page is not JObject pageObject)
                throw new InvalidOperationException("Kraken funding status response has an unsupported result type.");

            var rows = pageObject[collectionName] as JArray
                ?? pageObject["items"] as JArray
                ?? throw new InvalidOperationException($"Kraken funding status response has no {collectionName} array.");
            var nextCursor = pageObject["cursor"]?.ToString() ?? pageObject["next_cursor"]?.ToString();
            return (rows.OfType<JObject>().ToList(), nextCursor);
        }

        private async Task ValidateApiPermissionsAsync(CancellationToken cancellationToken)
        {
            var permissions = await FetchApiPermissionsAsync(cancellationToken).ConfigureAwait(false);
            var required = new[] { "query-funds", "query-ledger" };
            var missing = required.Where(permission => !permissions.Contains(permission)).ToList();
            if (missing.Count > 0)
                throw new InvalidOperationException($"Kraken API key is missing permissions: {String.Join(", ", missing)}.");
        }

        private async Task<List<KrakenLedgerEntry>> FetchLedgersAsync(
            DateTime startUtc,
            DateTime endUtc,
            CancellationToken cancellationToken)
        {
            var result = new List<KrakenLedgerEntry>();
            for (var offset = 0; ; offset += LedgerPageSize)
            {
                var page = await PostPrivateAsync(
                    "/0/private/Ledgers",
                    new Dictionary<string, string>
                    {
                        ["type"] = "all",
                        ["start"] = ToUnixSeconds(startUtc.AddTicks(-1)).ToString(CultureInfo.InvariantCulture),
                        ["end"] = ToUnixSeconds(endUtc).ToString(CultureInfo.InvariantCulture),
                        ["ofs"] = offset.ToString(CultureInfo.InvariantCulture)
                    },
                    cancellationToken).ConfigureAwait(false);
                var ledgerObject = page["ledger"] as JObject ?? new JObject();
                result.AddRange(ledgerObject.Properties().Select(property => ParseLedger(property.Name, property.Value)));
                var count = page["count"]?.Value<int>() ?? result.Count;
                if (result.Count >= count || ledgerObject.Count < LedgerPageSize)
                    break;
            }

            return result.OrderBy(ledger => ledger.Time).ThenBy(ledger => ledger.Id, StringComparer.Ordinal).ToList();
        }

        private static List<Record> CreateDailyRecords(
            Account account,
            DateTime date,
            List<KrakenLedgerEntry> ledgers,
            List<KrakenFundingTransaction> fundingTransactions)
        {
            var records = new List<Record>();
            foreach (var ledger in ledgers)
            {
                var ledgerRecord = AddRecord(
                    records,
                    account,
                    ledger.Asset,
                    date,
                    ledger.Amount,
                    GetLedgerReason(ledger.Type, ledger.Subtype),
                    $"Kraken ledger {ledger.Id}; refid={ledger.ReferenceId}; type={ledger.Type}; subtype={ledger.Subtype}; amount={ledger.Amount}; asset={ledger.Asset}",
                    includeZero: ledger.Amount != 0m,
                    isInternal: IsAssetExchangeLedger(ledger.Type));
                if (ledgerRecord is not null && ledger.Type is "deposit" or "withdrawal")
                    ApplyFundingTransaction(ledgerRecord, ledger, fundingTransactions);
                AddRecord(
                    records,
                    account,
                    ledger.Asset,
                    date,
                    -ledger.Fee,
                    "\u624b\u7eed\u8d39",
                    $"Kraken ledger fee {ledger.Id}; refid={ledger.ReferenceId}; fee={ledger.Fee}; asset={ledger.Asset}",
                    includeZero: ledger.Fee != 0m);
            }

            return records;
        }

        private static Record? AddRecord(
            List<Record> records,
            Account account,
            string rawAsset,
            DateTime date,
            decimal amount,
            string reason,
            string source,
            bool includeZero = false,
            bool isInternal = false)
        {
            if (amount == 0 && !includeZero)
                return null;
            var displayAsset = NormalizeAsset(rawAsset);
            var record = new Record
            {
                Account = account,
                Holding = CreateHolding(account, displayAsset, 0m),
                date = date.Date,
                postingDate = date.Date,
                updateTime = DateTime.Now,
                HoldingQuantity = 0,
                DestAccount = displayAsset,
                Source = source,
                Reason = reason,
                isInternal = isInternal
            };
            record.CopyFrom(new Currency(amount, ParseAssetCurrency(displayAsset)));
            records.Add(record);
            return record;
        }

        private static void ApplyFundingTransaction(
            Record record,
            KrakenLedgerEntry ledger,
            List<KrakenFundingTransaction> fundingTransactions)
        {
            var matches = fundingTransactions.Where(item =>
                    String.Equals(item.ReferenceId, ledger.ReferenceId, StringComparison.Ordinal)
                    && String.Equals(NormalizeAsset(item.Asset), NormalizeAsset(ledger.Asset), StringComparison.Ordinal)
                    && item.IsWithdrawal == (ledger.Type == "withdrawal")
                    && (item.IsWithdrawal ? -item.Amount : item.Amount) == ledger.Amount
                    && String.Equals(item.Status, "Success", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count != 1)
                throw new InvalidOperationException($"Kraken deposit funding match is not unique: ledger={ledger.Id}; refid={ledger.ReferenceId}; matches={matches.Count}.");

            var funding = matches[0];
            record.Source += $"; fundingMethod={funding.Method}; fundingStatus={funding.Status}";
            if (!IsEthereumMainnetFundingMethod(funding.Method))
                return;

            CryptoUtil.ValidateEthereumAddressValue(funding.Address);
            if (!funding.TransactionHash.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                || funding.TransactionHash.Length != 66)
            {
                throw new InvalidOperationException($"Kraken deposit has invalid Ethereum transaction hash: refid={funding.ReferenceId}; txid={funding.TransactionHash}.");
            }

            var asset = NormalizeAsset(ledger.Asset);
            record.blockchain = BlockchainType.Ethereum;
            record.blockchainTransactionHash = funding.TransactionHash.ToLowerInvariant();
            record.blockchainEventIndex = null;
            record.blockchainAssetContract = asset == "USDT" ? EthereumUsdtContract : "";
            record.Source += $"; txid={funding.TransactionHash}; address={funding.Address}";
        }

        private static bool IsEthereumMainnetFundingMethod(string method)
        {
            return method is "Ether (Hex)" or "Ethereum" or "Ethereum (ERC20)" or "Tether USD (ERC20)"
                || method.EndsWith(" - Ethereum (Unified)", StringComparison.Ordinal);
        }

        private static List<Holding> CreateHoldings(
            Account account,
            Dictionary<string, decimal> quantities)
        {
            return quantities
                .Where(item => item.Value != 0)
                .Select(item => CreateHolding(
                    account,
                    NormalizeAsset(item.Key),
                    item.Value))
                .OrderBy(holding => holding.code, StringComparer.Ordinal)
                .ToList();
        }

        private static Holding CreateHolding(Account account, string asset, decimal quantity)
        {
            return new Holding(asset, asset == "USD" ? HoldingType.Cash : HoldingType.Crypto)
            {
                Account = account,
                desc = asset == "USD" ? "USD cash balance" : $"Kraken {asset}",
                displayText = asset,
                currentPrice = new Currency(quantity, ParseAssetCurrency(asset))
            };
        }

        private static List<AccountBalance> CreateAccountBalances(Account account, Dictionary<string, decimal> quantities)
        {
            return quantities
                .Select(item => new AccountBalance(account, new Currency(item.Value, ParseAssetCurrency(NormalizeAsset(item.Key)))))
                .ToList();
        }

        private static CurrencyType ParseAssetCurrency(string asset)
        {
            var normalized = NormalizeAsset(asset);
            var suffixIndex = normalized.IndexOf('.', StringComparison.Ordinal);
            var currencyCode = suffixIndex < 0 ? normalized : normalized[..suffixIndex];
            return Enum.TryParse<CurrencyType>(currencyCode, out var currency)
                ? currency
                : throw new InvalidOperationException($"Unsupported Kraken currency: {asset}.");
        }

        private async Task<JObject> PostPrivateAsync(
            string path,
            IReadOnlyDictionary<string, string> parameters,
            CancellationToken cancellationToken)
        {
            var result = await PostPrivateResultAsync(path, parameters, cancellationToken).ConfigureAwait(false);
            return result as JObject
                ?? throw new InvalidOperationException("Kraken response does not contain an object result.");
        }

        private async Task<JToken> PostPrivateResultAsync(
            string path,
            IReadOnlyDictionary<string, string> parameters,
            CancellationToken cancellationToken)
        {
            await requestLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                for (var attempt = 1; ; attempt++)
                {
                    try
                    {
                        var nonce = NextNonce();
                        var fields = new List<KeyValuePair<string, string>> { new("nonce", nonce) };
                        fields.AddRange(parameters);
                        var postData = String.Join("&", fields.Select(field =>
                            $"{Uri.EscapeDataString(field.Key)}={Uri.EscapeDataString(field.Value)}"));
                        using var request = new HttpRequestMessage(HttpMethod.Post, KrakenApiBaseUrl + path)
                        {
                            Content = new StringContent(postData, Encoding.UTF8, "application/x-www-form-urlencoded")
                        };
                        request.Headers.Add("API-Key", apiKey);
                        request.Headers.Add("API-Sign", CreateSignature(path, nonce, postData));

                        using var client = CreateHttpClient();
                        using var response = await client.SendAsync(request, cancellationToken).ConfigureAwait(false);
                        var responseText = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                        if (!response.IsSuccessStatusCode)
                            throw new InvalidOperationException($"Kraken request failed: {(int)response.StatusCode} {response.ReasonPhrase}.");

                        var json = JObject.Parse(responseText);
                        ThrowKrakenErrors(json, "Kraken request failed");
                        return json["result"]
                            ?? throw new InvalidOperationException("Kraken response does not contain a result.");
                    }
                    catch (HttpRequestException) when (attempt < 3)
                    {
                        await Task.Delay(TimeSpan.FromMilliseconds(500 * attempt), cancellationToken).ConfigureAwait(false);
                    }
                }
            }
            catch (JsonException exception)
            {
                throw new InvalidOperationException("Kraken returned invalid JSON.", exception);
            }
            finally
            {
                requestLock.Release();
            }
        }

        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient { Timeout = RequestTimeout };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 KrakenDailyImporter");
            return client;
        }

        private string CreateSignature(string path, string nonce, string postData)
        {
            var messageHash = SHA256.HashData(Encoding.UTF8.GetBytes(nonce + postData));
            var pathBytes = Encoding.UTF8.GetBytes(path);
            var signedData = new byte[pathBytes.Length + messageHash.Length];
            Buffer.BlockCopy(pathBytes, 0, signedData, 0, pathBytes.Length);
            Buffer.BlockCopy(messageHash, 0, signedData, pathBytes.Length, messageHash.Length);
            return Convert.ToBase64String(HMACSHA512.HashData(apiSecret, signedData));
        }

        private static KrakenExtendedBalance ParseExtendedBalance(string asset, JToken value)
        {
            var balance = ReadDecimal(value["balance"]);
            var credit = ReadDecimal(value["credit"]);
            var creditUsed = ReadDecimal(value["credit_used"]);
            var holdTrade = ReadDecimal(value["hold_trade"]);
            return new KrakenExtendedBalance(asset, balance, credit, creditUsed, holdTrade, balance + credit - creditUsed - holdTrade);
        }

        private static KrakenLedgerEntry ParseLedger(string id, JToken value)
        {
            var unixTime = ReadDecimal(value["time"]);
            return new KrakenLedgerEntry(
                id,
                value["refid"]?.ToString() ?? "",
                DateTimeOffset.FromUnixTimeMilliseconds((long)(unixTime * 1000m)).UtcDateTime,
                value["type"]?.ToString() ?? "",
                value["subtype"]?.ToString() ?? "",
                value["asset"]?.ToString() ?? throw new InvalidOperationException($"Kraken ledger {id} is missing asset."),
                ReadDecimal(value["amount"]),
                ReadDecimal(value["fee"]),
                ReadDecimal(value["balance"]));
        }

        private static KrakenFundingTransaction ParseFundingTransaction(JObject value, bool isWithdrawal)
        {
            var unixTime = ReadDecimal(value["time"]);
            return new KrakenFundingTransaction(
                value["method"]?.ToString() ?? "",
                value["refid"]?.ToString() ?? "",
                value["txid"]?.ToString() ?? "",
                value["info"]?.ToString() ?? "",
                value["asset"]?.ToString() ?? "",
                ReadDecimal(value["amount"]),
                ReadDecimal(value["fee"]),
                DateTimeOffset.FromUnixTimeMilliseconds((long)(unixTime * 1000m)).UtcDateTime,
                value["status"]?.ToString() ?? "",
                isWithdrawal);
        }

        private static Dictionary<string, decimal> DeriveQuantitiesAt(
            Dictionary<string, decimal> current,
            IEnumerable<KrakenLedgerEntry> ledgers,
            DateTime timeUtc)
        {
            var result = new Dictionary<string, decimal>(current, StringComparer.Ordinal);
            foreach (var ledger in ledgers.Where(ledger => ledger.Time >= timeUtc))
                result[ledger.Asset] = GetQuantity(result, ledger.Asset) - ledger.Amount + ledger.Fee;
            return result;
        }

        private static void ValidateQuantityChainToCurrent(
            Dictionary<string, decimal> beginning,
            IEnumerable<KrakenLedgerEntry> ledgers,
            DateTime startUtc,
            Dictionary<string, decimal> current)
        {
            var calculated = new Dictionary<string, decimal>(beginning, StringComparer.Ordinal);
            foreach (var ledger in ledgers.Where(ledger => ledger.Time >= startUtc))
            {
                calculated[ledger.Asset] = GetQuantity(calculated, ledger.Asset) + ledger.Amount - ledger.Fee;
                if (calculated[ledger.Asset] != ledger.Balance)
                {
                    throw new InvalidOperationException(
                        $"Kraken ledger balance mismatch: id={ledger.Id}, asset={ledger.Asset}, calculated={calculated[ledger.Asset]}, reported={ledger.Balance}.");
                }
            }
            ValidateQuantitiesEqual(calculated, current, "Kraken current BalanceEx");
        }

        private static void ValidateQuantitiesEqual(
            Dictionary<string, decimal> expected,
            Dictionary<string, decimal> actual,
            string context)
        {
            foreach (var asset in expected.Keys.Union(actual.Keys, StringComparer.Ordinal))
            {
                var expectedValue = GetQuantity(expected, asset);
                var actualValue = GetQuantity(actual, asset);
                if (expectedValue != actualValue)
                    throw new InvalidOperationException($"{context} quantity mismatch: asset={asset}, expected={expectedValue}, actual={actualValue}.");
            }
        }

        private static decimal GetQuantity(Dictionary<string, decimal> quantities, string asset)
        {
            return quantities.TryGetValue(asset, out var value) ? value : 0m;
        }

        private static string GetLedgerReason(string type, string subtype)
        {
            return type switch
            {
                "trade" or "spend" or "receive" => "\u8d44\u4ea7\u5151\u6362",
                "deposit" => "\u5145\u503c",
                "withdrawal" => "\u63d0\u73b0",
                "transfer" => "\u8f6c\u8d26",
                "staking" => "\u8d28\u62bc\u6536\u76ca",
                "dividend" => "\u80a1\u606f",
                "rollover" => "\u5229\u606f",
                "margin" => "\u4fdd\u8bc1\u91d1",
                "adjustment" => "\u8d26\u52a1\u8c03\u6574",
                "credit" => "\u4fe1\u7528",
                "settled" => "\u7ed3\u7b97",
                "sale" => "\u51fa\u552e",
                _ => String.IsNullOrWhiteSpace(subtype) ? type : $"{type}/{subtype}"
            };
        }

        private static bool IsAssetExchangeLedger(string type)
        {
            return type is "trade" or "spend" or "receive";
        }

        private static string NormalizeAsset(string asset)
        {
            var suffixIndex = asset.IndexOf('.', StringComparison.Ordinal);
            var baseAsset = suffixIndex < 0 ? asset : asset[..suffixIndex];
            var suffix = suffixIndex < 0 ? "" : asset[suffixIndex..];
            var normalized = baseAsset switch
            {
                "XXBT" or "XBT" => "BTC",
                "ZUSD" => "USD",
                _ when baseAsset.StartsWith("X", StringComparison.Ordinal) && baseAsset.Length == 4 => baseAsset[1..],
                _ when baseAsset.StartsWith("Z", StringComparison.Ordinal) && baseAsset.Length == 4 => baseAsset[1..],
                _ => baseAsset
            };
            return normalized + suffix;
        }

        private static long ToUnixSeconds(DateTime utc)
        {
            return new DateTimeOffset(DateTime.SpecifyKind(utc, DateTimeKind.Utc)).ToUnixTimeSeconds();
        }

        private static decimal ReadDecimal(JToken? token)
        {
            var value = token?.ToString();
            return String.IsNullOrWhiteSpace(value)
                ? 0m
                : Decimal.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        private static void ThrowKrakenErrors(JObject json, string prefix)
        {
            var errors = json["error"] as JArray;
            if (errors is { Count: > 0 })
                throw new InvalidOperationException($"{prefix}: {String.Join("; ", errors.Values<string>())}");
        }

        private static string NextNonce()
        {
            while (true)
            {
                var previous = Interlocked.Read(ref lastNonce);
                var candidate = Math.Max(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() * 1000, previous + 1);
                if (Interlocked.CompareExchange(ref lastNonce, candidate, previous) == previous)
                    return candidate.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static string RequiredConfig(IConfigurationRoot config, string key)
        {
            var value = config[key];
            return String.IsNullOrWhiteSpace(value)
                ? throw new InvalidOperationException($"Missing {key} in config.json.")
                : value.Trim();
        }
    }

    sealed record KrakenExtendedBalance(
        string Asset,
        decimal Balance,
        decimal Credit,
        decimal CreditUsed,
        decimal HoldTrade,
        decimal AvailableBalance);

    sealed record KrakenLedgerEntry(
        string Id,
        string ReferenceId,
        DateTime Time,
        string Type,
        string Subtype,
        string Asset,
        decimal Amount,
        decimal Fee,
        decimal Balance);

    sealed record KrakenFundingTransaction(
        string Method,
        string ReferenceId,
        string TransactionHash,
        string Address,
        string Asset,
        decimal Amount,
        decimal Fee,
        DateTime Time,
        string Status,
        bool IsWithdrawal);
}
