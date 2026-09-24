using System.Buffers.Binary;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MyBook
{
    partial class WebUtil
    {
        private static readonly SemaphoreSlim firstTradeLock = new(1, 1);

        public bool IsFirstTradeConfigured => new[] { "firsttrade_username", "firsttrade_password", "firsttrade_totp_secret" }
            .Any(key => !String.IsNullOrWhiteSpace(config[key]));

        public async Task FetchFirstTradeAsync(CancellationToken cancellationToken = default)
        {
            if (!await firstTradeLock.WaitAsync(0, cancellationToken).ConfigureAwait(false))
                return;
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromMinutes(5));
            var stage = "configuration";
            try
            {
                string Required(string key) => !String.IsNullOrWhiteSpace(config[key])
                    ? config[key]! : throw new FirstTradeException("incomplete local configuration");
                var username = Required("firsttrade_username");
                var password = Required("firsttrade_password");
                var totpSecret = Required("firsttrade_totp_secret");
                stage = "read database checkpoint";
                var imports = database.GetStatementImports(StatementImportProvider.FirstTradeApi);
                var checkpoint = database.GetStatementImportCheckpointTime(StatementImportProvider.FirstTradeApi)
                    ?? throw new FirstTradeException("initial database checkpoint is missing");
                stage = "acquire database session";
                using var sessionStore = database.OpenFirstTradeSession(username);
                using var client = new FirstTradeClient(username, password, totpSecret, sessionStore, proxy: config["mail_proxy"]);
                stage = "login";
                await client.LoginAsync(timeout.Token).ConfigureAwait(false);
                stage = "read account data";
                var capture = await client.FetchAsync(null, timeout.Token, account =>
                {
                    database.GetAccountByTypeAndId("FIRSTTRADE", account);
                    var last = imports.Where(i => i.statementKey.StartsWith(FirstTradeAccountKey(account), StringComparison.Ordinal))
                        .Select(i => (DateTime?)i.time).Max();
                    return last.HasValue && last.Value.Date.AddDays(-7) > checkpoint.Date
                        ? last.Value.Date.AddDays(-7) : checkpoint.Date;
                }).ConfigureAwait(false);
                stage = "validate and import account data";
                timeout.Token.ThrowIfCancellationRequested();
                sessionStore.EnsureLock();
                var recordCount = ImportFirstTradeCapture(capture);
                Console.WriteLine($"FirstTrade imported {capture.Accounts.Count} account(s), {recordCount} record(s); cash, positions and account values validated.");
            }
            catch (FirstTradeException) { throw; }
            catch (OperationCanceledException)
            {
                throw new FirstTradeException($"{stage}: cancelled or timed out");
            }
            catch (Exception)
            {
                // Never propagate HTTP, JSON or database exception details containing private data.
                throw new FirstTradeException($"{stage}: failed; private details suppressed");
            }
            finally { firstTradeLock.Release(); }
        }

        private static string FirstTradeAccountKey(string account) =>
            "FirstTrade:" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(account))) + ":";

        private int ImportFirstTradeCapture(FirstTradeCapture capture)
        {
            var imports = new List<StatementRecordHoldingImport>();
            foreach (var item in capture.Accounts)
            {
                var account = database.GetAccountByTypeAndId("FIRSTTRADE", item.Account);
                imports.Add(BuildFirstTradeImport(capture, item, account,
                    database.GetCurrentAccountHoldings(account),
                    database.GetStatementRecords(StatementImportProvider.FirstTradeApi, account),
                    database.GetKnownEquityHoldingType));
            }
            var saved = database.SaveStatementRecordsAndHoldingsOnce(imports);
            return imports.Where((_, index) => saved[index]).Sum(import => import.Records.Count);
        }

        internal static StatementRecordHoldingImport BuildFirstTradeImport(FirstTradeCapture capture,
            FirstTradeAccountCapture item, Account account, List<Holding> beginning,
            List<Record> previousRecords, Func<string, HoldingType> resolveEquity)
        {
            const string transactionPrefix = "FirstTrade transaction|";
            var time = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(capture.CompletedAtUtc, "Eastern Standard Time").DateTime;
            var key = FirstTradeAccountKey(item.Account) + capture.CompletedAtUtc.ToUnixTimeMilliseconds().ToString(CultureInfo.InvariantCulture);
            if (account.relativeBalance || account.isCredit || beginning.Any(h => h.currentPrice.t != CurrencyType.USD))
                throw new FirstTradeException("account must be an absolute-balance USD investment account");
            if (beginning.Any(h => h.holdingType is not (HoldingType.Cash or HoldingType.NASDAQ or HoldingType.ARCA)))
                throw new FirstTradeException("unsupported existing holding type");
            if (!item.Balances.TryGetProperty("result", out var balance) || balance.ValueKind != JsonValueKind.Object)
                throw new FirstTradeException("missing balance result");
            if (FirstTradeText(balance, "account") != item.Account)
                throw new FirstTradeException("balance account differs from account list");
            foreach (var field in new[] { "short_stock_value", "long_option_value", "short_option_value", "fixed_income_value", "mutual_funds_value" })
                if (balance.TryGetProperty(field, out _) && FirstTradeNumber(balance, field) != 0)
                    throw new FirstTradeException("unsupported non-equity asset balance");
            var endingCash = FirstTradeNumber(balance, "cash_balance") + FirstTradeNumber(balance, "margin_balance");
            var holdings = new List<Holding>
            {
                new("USD", HoldingType.Cash) { Account = account, currentPrice = new Currency(endingCash, CurrencyType.USD) }
            };
            var equities = new Dictionary<string, Holding>(StringComparer.Ordinal);
            foreach (var row in item.PositionPages.SelectMany(page => page.GetProperty("items").EnumerateArray()))
            {
                if (FirstTradeNumber(row, "sec_type") != 1)
                    throw new FirstTradeException("unsupported security type in positions");
                var symbol = FirstTradeText(row, "symbol");
                var holding = new Holding(symbol, resolveEquity(symbol))
                {
                    Account = account, quantity = FirstTradeNumber(row, "quantity"),
                    currentPrice = new Currency(FirstTradeNumber(row, "last"), CurrencyType.USD),
                    desc = FirstTradeText(row, "company_name"), displayText = FirstTradeText(row, "company_name")
                };
                if (holding.quantity < 0 || holding.currentPrice.v < 0 || !equities.TryAdd(symbol, holding))
                    throw new FirstTradeException("negative or duplicate equity position");
                FirstTradeEqual(holding.totalPrice.v, FirstTradeNumber(row, "market_value"), "position quantity x price");
                holdings.Add(holding);
            }
            // 由于 Firstrade 自身的余额页与持仓页使用不同的行情供应商，总额和分项之和可能对不上。
            // Only these two FirstTrade aggregate checks allow a difference below USD 100.
            // Holdings and records always use detail values; never create a residual adjustment.
            var equityTotal = holdings.Where(h => h.holdingType != HoldingType.Cash).Sum(h => h.totalPrice.v);
            if (Math.Abs(equityTotal - FirstTradeNumber(balance, "long_stock_value")) >= 100m)
                throw new FirstTradeException("equity subtotal: difference must be less than USD 100");
            var endingTotal = holdings.Sum(h => h.totalPrice.v);
            if (Math.Abs(endingTotal - FirstTradeNumber(balance, "total_account_value")) >= 100m)
                throw new FirstTradeException("account total: difference must be less than USD 100");

            var transactions = new List<FirstTradeTransaction>();
            var occurrences = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var row in item.HistoryPages.SelectMany(page => page.GetProperty("items").EnumerateArray()))
            {
                if (!DateTime.TryParseExact(FirstTradeText(row, "report_date"), "yyyy-MM-dd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var date) || date < item.HistoryFrom.Date || date > item.HistoryThrough.Date)
                    throw new FirstTradeException("invalid or out-of-range transaction date");
                var type = FirstTradeText(row, "trans_str");
                var description = FirstTradeText(row, "description");
                var subaccount = FirstTradeText(row, "account_type");
                var symbol = FirstTradeText(row, "symbol", allowEmpty: true);
                var quantity = FirstTradeNumber(row, "quantity");
                var price = FirstTradeNumber(row, "trade_price");
                var amount = FirstTradeNumber(row, "amount");
                if (subaccount is not ("Cash" or "Margin"))
                    throw new FirstTradeException("unsupported transaction subaccount");
                var canonical = JsonSerializer.Serialize(new { date, type, description, subaccount, symbol, quantity, price, amount });
                var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical)));
                occurrences.TryGetValue(hash, out var occurrence);
                occurrences[hash] = ++occurrence;
                transactions.Add(new(hash + ":" + occurrence, date, type, description, subaccount, symbol, quantity, price, amount));
            }
            // The API has no transaction ID. Preserve multiplicity and reject removals/revisions in the overlap window.
            var incoming = transactions.Select(t => t.Key).ToHashSet(StringComparer.Ordinal);
            var previous = previousRecords.Where(r => r.Source.StartsWith(transactionPrefix, StringComparison.Ordinal)
                && r.Source.Contains("|cash|", StringComparison.Ordinal)).ToList();
            foreach (var record in previous.Where(r => r.date.Date >= item.HistoryFrom.Date && r.date.Date <= item.HistoryThrough.Date))
                if (!incoming.Contains(record.Source.Split('|')[1]))
                    throw new FirstTradeException("previously imported transaction was removed or revised in the overlap window");
            var known = previous.Select(r => r.Source.Split('|')[1]).ToHashSet(StringComparer.Ordinal);
            var transfers = transactions.Where(t => t.Type == "OTHER").ToList();
            foreach (var group in transfers.GroupBy(t => (t.Date, Amount: Math.Abs(t.Amount))))
            {
                if (group.Key.Amount == 0 || group.Any(t => t.Description is not ("XFER CASH TO MARGIN" or "XFER MARGIN TO CASH"))
                    || group.Sum(t => t.Amount) != 0 || group.Count(t => t.Subaccount == "Cash") != group.Count(t => t.Subaccount == "Margin")
                    || group.Where(t => t.Subaccount == "Cash").Sum(t => t.Amount) != -group.Where(t => t.Subaccount == "Margin").Sum(t => t.Amount))
                    throw new FirstTradeException("unrecognized or unpaired cash/margin transfer");
            }
            var records = new List<Record>();
            var cash = beginning.Where(h => h.holdingType == HoldingType.Cash).Sum(h => h.totalPrice.v);
            var quantities = beginning.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => h.code, h => h.quantity, StringComparer.Ordinal);
            var values = beginning.Where(h => h.holdingType != HoldingType.Cash).ToDictionary(h => h.code, h => h.totalPrice.v, StringComparer.Ordinal);
            foreach (var tx in transactions.OrderBy(t => t.Date).ThenBy(t => t.Key, StringComparer.Ordinal))
            {
                if (known.Contains(tx.Key)) continue;
                string reason;
                var trade = tx.Type is "BOUGHT" or "SOLD";
                var isInternal = trade || tx.Type == "OTHER";
                if (trade)
                {
                    if (tx.Quantity <= 0 || tx.Price <= 0 || tx.Symbol.Length == 0
                        || (tx.Type == "BOUGHT" ? tx.Amount >= 0 : tx.Amount <= 0))
                        throw new FirstTradeException("invalid equity trade");
                    // Settlement cents are explicit; unexplained commissions must not become valuation records.
                    FirstTradeEqual(Math.Abs(tx.Amount), Decimal.Round(tx.Quantity * tx.Price, 2, MidpointRounding.AwayFromZero),
                        "trade settlement amount (separate fee detail required on mismatch)");
                    reason = tx.Type == "BOUGHT" ? "买入" : "卖出";
                    var quantity = tx.Type == "BOUGHT" ? tx.Quantity : -tx.Quantity;
                    var security = equities.GetValueOrDefault(tx.Symbol)
                        ?? new Holding(tx.Symbol, resolveEquity(tx.Symbol)) { Account = account, currentPrice = new Currency(0, CurrencyType.USD) };
                    quantities[tx.Symbol] = quantities.GetValueOrDefault(tx.Symbol) + quantity;
                    values[tx.Symbol] = values.GetValueOrDefault(tx.Symbol) - tx.Amount;
                    AddRecord(-tx.Amount, reason, transactionPrefix + tx.Key + "|asset|" + tx.Description,
                        tx.Date, true, security, quantity);
                }
                else
                {
                    if (tx.Quantity != 0 || tx.Price != 0)
                        throw new FirstTradeException("non-trade transaction changes security quantity or price");
                    reason = tx.Type switch
                    {
                        "DEPOSIT" when tx.Amount > 0 => "转入",
                        "WITHDRAWAL" when tx.Amount < 0 => "转出",
                        "INTEREST" when tx.Description.StartsWith("FULLYPAID LENDING REBATE", StringComparison.Ordinal) => "证券出借收益",
                        "INTEREST" => "现金利息",
                        "DIVIDEND" => "股息",
                        "FEE" when tx.Amount < 0 => "手续费",
                        "TAX" when tx.Amount < 0 => "税费",
                        "OTHER" => "现金保证金划转",
                        _ => throw new FirstTradeException("unsupported account history transaction type")
                    };
                }
                cash += tx.Amount;
                AddRecord(tx.Amount, reason, transactionPrefix + tx.Key + "|cash|" + tx.Type + " " + tx.Subaccount + " " + tx.Description,
                    tx.Date, isInternal, null, 0);
            }
            FirstTradeEqual(cash, endingCash, "cash balance from transaction details");
            foreach (var symbol in quantities.Keys.Union(equities.Keys).Union(values.Keys))
            {
                var ending = equities.GetValueOrDefault(symbol);
                FirstTradeEqual(quantities.GetValueOrDefault(symbol), ending?.quantity ?? 0, "security quantity from transactions");
                var change = (ending?.totalPrice.v ?? 0) - values.GetValueOrDefault(symbol);
                if (change == 0) continue;
                var security = ending ?? new Holding(symbol, resolveEquity(symbol)) { Account = account, currentPrice = new Currency(0, CurrencyType.USD) };
                AddRecord(change, "持仓价格变动", "FirstTrade valuation|" + key + "|" + symbol,
                    time.Date, false, security, 0);
            }
            var beginningTotal = beginning.Sum(h => h.totalPrice.v);
            FirstTradeEqual(beginningTotal + records.Sum(r => r.v), endingTotal, "beginning value plus records");
            return new StatementRecordHoldingImport(StatementImportProvider.FirstTradeApi, time, key, account, records, holdings,
                [new AccountBalance(account, new Currency(endingTotal, CurrencyType.USD))],
                [new AccountBalance(account, new Currency(beginningTotal, CurrencyType.USD))], beginning,
                sourceDataJson: JsonSerializer.Serialize(new FirstTradeCapture
                {
                    Version = capture.Version, StartedAtUtc = capture.StartedAtUtc,
                    CompletedAtUtc = capture.CompletedAtUtc, Accounts = [item]
                }));

            void AddRecord(decimal amount, string reason, string source, DateTime date, bool isInternal, Holding? holding, decimal quantity)
            {
                records.Add(new Record
                {
                    Account = account, v = amount, t = CurrencyType.USD, date = date, postingDate = date,
                    updateTime = DateTime.Now, Reason = reason, Source = source[..Math.Min(1024, source.Length)],
                    isInternal = isInternal, Holding = holding, HoldingQuantity = quantity
                });
            }
        }

        private sealed record FirstTradeTransaction(string Key, DateTime Date, string Type, string Description,
            string Subaccount, string Symbol, decimal Quantity, decimal Price, decimal Amount);

        private static decimal FirstTradeNumber(JsonElement row, string field)
        {
            if (!row.TryGetProperty(field, out var value)
                || !(value.ValueKind == JsonValueKind.Number && value.TryGetDecimal(out _)
                    || value.ValueKind == JsonValueKind.String && Decimal.TryParse(value.GetString(), NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                        CultureInfo.InvariantCulture, out _)))
                throw new FirstTradeException("missing or invalid numeric field: " + field);
            var result = value.ValueKind == JsonValueKind.Number ? value.GetDecimal() : Decimal.Parse(value.GetString()!, CultureInfo.InvariantCulture);
            MySqlDecimalColumnTypes.ValidateCurrencyValue(result, "FirstTrade " + field);
            return result;
        }

        private static string FirstTradeText(JsonElement row, string field, bool allowEmpty = false)
        {
            if (!row.TryGetProperty(field, out var value) || value.ValueKind != JsonValueKind.String
                || !allowEmpty && String.IsNullOrWhiteSpace(value.GetString()))
                throw new FirstTradeException("missing or invalid text field: " + field);
            return value.GetString()!;
        }

        private static void FirstTradeEqual(decimal actual, decimal expected, string stage)
        {
            if (actual != expected) throw new FirstTradeException(stage + ": exact validation failed");
        }

        internal sealed class FirstTradeCapture
        {
            public int Version { get; set; } = 1;
            public DateTimeOffset StartedAtUtc { get; set; }
            public DateTimeOffset CompletedAtUtc { get; set; }
            public List<FirstTradeAccountCapture> Accounts { get; set; } = new();
        }

        internal sealed class FirstTradeAccountCapture
        {
            public string Account { get; set; } = "";
            public DateTime HistoryFrom { get; set; }
            public DateTime HistoryThrough { get; set; }
            public JsonElement Summary { get; set; }
            public JsonElement Balances { get; set; }
            public List<JsonElement> PositionPages { get; set; } = new();
            public List<JsonElement> HistoryPages { get; set; } = new();
        }

        internal sealed class FirstTradeException(string message) : Exception("FirstTrade: " + message);

        internal sealed class FirstTradeClient : IDisposable
        {
            private enum Endpoint { Bootstrap, Login, VerifyCode, Accounts, Balances, Positions, History }
            private static readonly Uri Origin = new("https://api3x.firstrade.com/");
            // Shared client value from firstrade-api/urls.py, not a personal session credential.
            private const string ClientToken = "833w3XuIFycv18ybi";
            private readonly HttpClient client;
            private readonly string username, password;
            private readonly byte[] totpKey;
            private string? ftat, sid;
            private int loginAttempts;
            private long responseBytes;
            private readonly Func<DateTimeOffset> utcNow;
            private readonly DatabaseUtil.FirstTradeSessionLease sessionStore;
            private readonly CookieContainer? cookies;
            private FirstTradeSessionState session = new();

            internal FirstTradeClient(string username, string password, string totpSecret, DatabaseUtil.FirstTradeSessionLease sessionStore,
                HttpMessageHandler? handler = null, string? proxy = null,
                Func<DateTimeOffset>? utcNow = null)
            {
                this.username = username;
                this.password = password;
                this.sessionStore = sessionStore;
                this.utcNow = utcNow ?? (() => DateTimeOffset.UtcNow);
                var transport = handler ?? CreateHandler(proxy);
                cookies = (transport as HttpClientHandler)?.CookieContainer;
                client = new HttpClient(transport)
                    { Timeout = TimeSpan.FromSeconds(30), MaxResponseContentBufferSize = 16 * 1024 * 1024 };
                client.DefaultRequestHeaders.UserAgent.ParseAdd("okhttp/4.9.2");
                try { totpKey = DecodeTotpSecret(totpSecret); }
                catch { client.Dispose(); throw; }
                try
                {
                    var stored = sessionStore.Read();
                    if (stored is not null)
                    {
                        if (stored.Length > 1024 * 1024)
                            throw new FirstTradeException("invalid database session state");
                        session = JsonSerializer.Deserialize<FirstTradeSessionState>(stored)
                            ?? throw new FirstTradeException("invalid database session state");
                        if (session.Version != 1 || session.Failures is < 0 or > 24 || session.Cookies is null)
                            throw new FirstTradeException("invalid database session state");
                    }
                    ftat = session.Ftat;
                    sid = session.Sid;
                    if (cookies is not null)
                        foreach (var cookie in session.Cookies.Where(c => c.Expires == DateTime.MinValue || c.Expires.ToUniversalTime() > this.utcNow().UtcDateTime))
                            cookies.Add(new Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain)
                                { Secure = cookie.Secure, HttpOnly = cookie.HttpOnly, Expires = cookie.Expires });
                }
                catch
                {
                    client.Dispose();
                    CryptographicOperations.ZeroMemory(totpKey);
                    throw new FirstTradeException("database session state unavailable or invalid; login was not attempted");
                }
            }

            internal static HttpClientHandler CreateHandler(string? proxy)
            {
                IWebProxy? configuredProxy = null;
                if (!String.IsNullOrWhiteSpace(proxy))
                {
                    if (!Uri.TryCreate(proxy.Trim(), UriKind.Absolute, out var uri)
                        || uri.Scheme is not ("http" or "socks5" or "socks") || uri.Port <= 0
                        || !String.IsNullOrEmpty(uri.UserInfo) || uri.AbsolutePath.Trim('/').Length > 0
                        || uri.Query.Length > 0 || uri.Fragment.Length > 0)
                        throw new FirstTradeException("invalid mail_proxy configuration");
                    if (uri.Scheme == "socks") uri = new UriBuilder(uri) { Scheme = "socks5" }.Uri;
                    configuredProxy = new WebProxy(uri);
                }
                return new HttpClientHandler
                {
                    AllowAutoRedirect = false, UseProxy = configuredProxy is not null, Proxy = configuredProxy,
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate, UseCookies = true
                };
            }

            internal async Task LoginAsync(CancellationToken token)
            {
                EnsureRequestsAllowed();
                if (!String.IsNullOrEmpty(ftat) && !String.IsNullOrEmpty(sid)) return;
                if (utcNow() < session.NextLoginUtc)
                    throw new FirstTradeException($"login cooldown active until {session.NextLoginUtc:O}");
                if (++loginAttempts > 2)
                    throw new FirstTradeException("automatic login retry limit reached");
                token.ThrowIfCancellationRequested();
                // Reserve the failure delay before any request, so crashes/restarts cannot reset the budget.
                session.Failures = Math.Min(24, session.Failures + 1);
                session.NextLoginUtc = utcNow().AddHours(Math.Min(24, Math.Pow(2, Math.Min(5, session.Failures - 1))));
                session.Ftat = ftat = null;
                session.Sid = sid = null;
                await SaveSessionAsync().ConfigureAwait(false);
                sid = null;
                await RequestAsync(Endpoint.Bootstrap, null, null, token).ConfigureAwait(false);
                var login = await RequestAsync(Endpoint.Login, null, new()
                {
                    ["username"] = username, ["password"] = password
                }, token).ConfigureAwait(false);
                if (HasText(login, "ftat") && HasText(login, "sid") && !HasText(login, "t_token")
                    && (!login.TryGetProperty("mfa", out var mfa) || mfa.ValueKind == JsonValueKind.False))
                {
                    await AcceptSessionAsync(login).ConfigureAwait(false);
                    return;
                }
                if (!login.TryGetProperty("mfa", out var challengeType) || challengeType.ValueKind != JsonValueKind.True)
                    throw new FirstTradeException("login did not return an authenticator MFA challenge");
                var challenge = RequiredText(login, "t_token");
                // Avoid sending a code just before its 30-second time step expires.
                var remainingMilliseconds = 30000 - DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() % 30000;
                if (remainingMilliseconds < 2000)
                    await Task.Delay(TimeSpan.FromMilliseconds(remainingMilliseconds + 100), token).ConfigureAwait(false);
                var code = CreateTotpCode(totpKey, DateTimeOffset.UtcNow);
                var verified = await RequestAsync(Endpoint.VerifyCode, null, new()
                {
                    ["mfaCode"] = code,
                    ["remember_for"] = "30", ["t_token"] = challenge
                }, token).ConfigureAwait(false);
                await AcceptSessionAsync(verified).ConfigureAwait(false);
            }

            private async Task AcceptSessionAsync(JsonElement json)
            {
                ftat = RequiredText(json, "ftat");
                sid = RequiredText(json, "sid");
                session.Failures = 0;
                session.NextLoginUtc = utcNow().AddMinutes(15);
                await SaveSessionAsync().ConfigureAwait(false);
            }

            private void EnsureRequestsAllowed()
            {
                try { sessionStore.EnsureLock(); }
                catch { throw new FirstTradeException("database session lock lost; requests stopped"); }
                if (utcNow() < session.BlockedUntilUtc)
                    throw new FirstTradeException($"requests paused until {session.BlockedUntilUtc:O} after access denial or rate limiting");
            }

            private Task SaveSessionAsync()
            {
                session.Ftat = ftat;
                session.Sid = sid;
                if (cookies is not null)
                    session.Cookies = cookies.GetAllCookies().Cast<Cookie>().Where(c => !c.Expired)
                        .Select(c => new FirstTradeCookie(c.Name, c.Value, c.Path, c.Domain, c.Secure, c.HttpOnly, c.Expires)).ToList();
                try
                {
                    sessionStore.Save(JsonSerializer.Serialize(session));
                }
                catch { throw new FirstTradeException("cannot persist database session/login cooldown; requests stopped"); }
                return Task.CompletedTask;
            }

            internal async Task<FirstTradeCapture> FetchAsync(FirstTradeCapture? previous, CancellationToken token,
                Func<string, DateTime>? historyFrom = null)
            {
                var result = new FirstTradeCapture { StartedAtUtc = DateTimeOffset.UtcNow };
                var today = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(result.StartedAtUtc, "Eastern Standard Time").Date;
                var accounts = await ReadAsync(Endpoint.Accounts, null, token).ConfigureAwait(false);
                var items = RequiredItems(accounts);
                if (items.GetArrayLength() == 0)
                    throw new FirstTradeException("account list is empty");
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var item in items.EnumerateArray())
                {
                    var account = RequiredText(item, "account");
                    if (!seen.Add(account))
                        throw new FirstTradeException("duplicate account in response");
                    var last = previous?.Accounts.SingleOrDefault(a => a.Account == account);
                    var from = historyFrom?.Invoke(account) ?? (last is null ? today.AddYears(-3) : last.HistoryThrough.Date.AddDays(-7));
                    if (from > today)
                        throw new FirstTradeException("history checkpoint is in the future");
                    var query = "account=" + Uri.EscapeDataString(account);
                    var balances = await ReadAsync(Endpoint.Balances, query, token).ConfigureAwait(false);
                    if (!balances.EnumerateObject().Any(p => p.Name != "error"))
                        throw new FirstTradeException("empty balance response");
                    var positions = await ReadPagesAsync(Endpoint.Positions, query, 200, token).ConfigureAwait(false);
                    var range = query + "&range=cust&range_arr%5B%5D=" + from.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        + "&range_arr%5B%5D=" + today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    var history = await ReadPagesAsync(Endpoint.History, range, 1000, token).ConfigureAwait(false);
                    result.Accounts.Add(new FirstTradeAccountCapture
                    {
                        Account = account, HistoryFrom = from, HistoryThrough = today, Summary = item.Clone(),
                        Balances = balances, PositionPages = positions, HistoryPages = history
                    });
                }
                result.CompletedAtUtc = DateTimeOffset.UtcNow;
                await SaveSessionAsync().ConfigureAwait(false);
                return result;
            }

            private async Task<List<JsonElement>> ReadPagesAsync(Endpoint endpoint, string query, int pageSize, CancellationToken token)
            {
                var pages = new List<JsonElement>();
                var hashes = new HashSet<string>(StringComparer.Ordinal);
                int count = 0;
                for (var page = 1; page <= 100; page++)
                {
                    var response = await ReadAsync(endpoint, query + $"&per_page={pageSize}&page={page}", token).ConfigureAwait(false);
                    var items = RequiredItems(response);
                    int length = items.GetArrayLength();
                    count += length;
                    if (length > pageSize || length > 0 && !hashes.Add(Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(items.GetRawText())))))
                        throw new FirstTradeException($"{endpoint}: invalid or repeated page");
                    pages.Add(response);
                    // Honor explicit totals when supplied; otherwise follow the requested page size.
                    int? total = PageInteger(response, "total") ?? PageInteger(response, "total_records");
                    int? lastPage = PageInteger(response, "last_page") ?? PageInteger(response, "total_pages");
                    bool? hasMore = null;
                    if (response.TryGetProperty("has_more", out var more))
                    {
                        if (more.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                            throw new FirstTradeException($"{endpoint}: invalid pagination flag");
                        hasMore = more.GetBoolean();
                    }
                    if (lastPage == 0 && count == 0) lastPage = 1;
                    if (total.HasValue && (total < count || total > count && length == 0))
                        throw new FirstTradeException($"{endpoint}: inconsistent pagination total");
                    var needsMore = total > count || lastPage > page || hasMore == true;
                    var ended = total == count || lastPage == page || hasMore == false;
                    if (needsMore && (ended || length == 0) || lastPage < page)
                        throw new FirstTradeException($"{endpoint}: inconsistent pagination metadata");
                    if (ended || !needsMore && length < pageSize)
                        return pages;
                }
                throw new FirstTradeException($"{endpoint}: pagination safety limit reached");
            }

            private static int? PageInteger(JsonElement json, string name)
            {
                if (!json.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null) return null;
                if (Int32.TryParse(value.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var number)) return number;
                throw new FirstTradeException("invalid pagination metadata");
            }

            private async Task<JsonElement> ReadAsync(Endpoint endpoint, string? query, CancellationToken token)
            {
                try { return await RequestAsync(endpoint, query, null, token).ConfigureAwait(false); }
                catch (FirstTradeSessionExpiredException)
                {
                    await LoginAsync(token).ConfigureAwait(false);
                    return await RequestAsync(endpoint, query, null, token).ConfigureAwait(false);
                }
            }

            private async Task<JsonElement> RequestAsync(Endpoint endpoint, string? query, Dictionary<string, string>? form, CancellationToken token)
            {
                EnsureRequestsAllowed();
                // Closed endpoint/method allowlist: there is no arbitrary URL or account mutation API.
                var (path, method) = endpoint switch
                {
                    Endpoint.Bootstrap => ("/", HttpMethod.Get),
                    Endpoint.Login => ("/sess/login", HttpMethod.Post),
                    Endpoint.VerifyCode => ("/sess/verify_pin", HttpMethod.Post),
                    Endpoint.Accounts => ("/private/acct_list", HttpMethod.Get),
                    Endpoint.Balances => ("/private/balances", HttpMethod.Get),
                    Endpoint.Positions => ("/private/positions", HttpMethod.Get),
                    Endpoint.History => ("/private/account_history", HttpMethod.Get),
                    _ => throw new FirstTradeException("endpoint is not allowed")
                };
                if ((method == HttpMethod.Post) != (form is not null))
                    throw new FirstTradeException("request method is not allowed");
                using var request = new HttpRequestMessage(method, new Uri(Origin, path + (query is null ? "" : "?" + query)));
                if (endpoint != Endpoint.Bootstrap)
                    request.Headers.Add("access-token", ClientToken);
                if (ftat is not null) request.Headers.Add("ftat", ftat);
                if (sid is not null) request.Headers.Add("sid", sid);
                if (form is not null) request.Content = new FormUrlEncodedContent(form);
                try
                {
                    using var response = await client.SendAsync(request, token).ConfigureAwait(false);
                    if (response.StatusCode is HttpStatusCode.Forbidden or HttpStatusCode.TooManyRequests)
                    {
                        var blockedUntil = utcNow().AddHours(1);
                        try
                        {
                            var retry = response.Headers.RetryAfter;
                            var retryUntil = retry?.Date ?? (retry?.Delta is TimeSpan delay ? utcNow().Add(delay) : (DateTimeOffset?)null);
                            if (retryUntil > blockedUntil) blockedUntil = retryUntil.Value;
                        }
                        catch (FormatException) { }
                        catch (ArgumentOutOfRangeException) { blockedUntil = DateTimeOffset.MaxValue; }
                        if (blockedUntil > session.BlockedUntilUtc) session.BlockedUntilUtc = blockedUntil;
                        await SaveSessionAsync().ConfigureAwait(false);
                        throw new FirstTradeException($"{endpoint}: HTTP {(int)response.StatusCode}; requests paused until {session.BlockedUntilUtc:O}; no automatic re-login");
                    }
                    if (response.StatusCode == HttpStatusCode.Unauthorized && method == HttpMethod.Get && endpoint != Endpoint.Bootstrap)
                    {
                        ftat = sid = null;
                        await SaveSessionAsync().ConfigureAwait(false);
                        throw new FirstTradeSessionExpiredException();
                    }
                    if (!response.IsSuccessStatusCode)
                        throw new FirstTradeException($"{endpoint}: HTTP {(int)response.StatusCode}");
                    if (endpoint == Endpoint.Bootstrap) return default;
                    var bytes = await response.Content.ReadAsByteArrayAsync(token).ConfigureAwait(false);
                    responseBytes += bytes.Length;
                    if (responseBytes > 64L * 1024 * 1024)
                        throw new FirstTradeException("capture response size safety limit reached");
                    using var json = JsonDocument.Parse(bytes);
                    if (json.RootElement.ValueKind != JsonValueKind.Object)
                        throw new FirstTradeException($"{endpoint}: unexpected response structure");
                    if (json.RootElement.TryGetProperty("error", out var error) && error.ValueKind != JsonValueKind.Null
                        && error.ValueKind != JsonValueKind.False && error.GetRawText() != "0" && error.GetRawText() != "\"\"")
                        throw new FirstTradeException($"{endpoint}: server rejected request");
                    return json.RootElement.Clone();
                }
                catch (FirstTradeException) { throw; }
                catch (FirstTradeSessionExpiredException) { throw; }
                catch (OperationCanceledException) { throw new FirstTradeException($"{endpoint}: cancelled or timed out"); }
                catch (HttpRequestException e) { throw new FirstTradeException($"{endpoint}: network failure ({e.HttpRequestError})"); }
                catch (Exception) { throw new FirstTradeException($"{endpoint}: transport or response failure; private details suppressed"); }
            }

            private static JsonElement RequiredItems(JsonElement json) => json.TryGetProperty("items", out var items) && items.ValueKind == JsonValueKind.Array
                ? items : throw new FirstTradeException("response does not contain an items array");
            private static string? TextOrNull(JsonElement json, string key) => json.TryGetProperty(key, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
            private static bool HasText(JsonElement json, string key) => !String.IsNullOrWhiteSpace(TextOrNull(json, key));
            private static string RequiredText(JsonElement json, string key) => HasText(json, key) ? TextOrNull(json, key)! : throw new FirstTradeException("required response field is missing");

            internal static byte[] DecodeTotpSecret(string secret)
            {
                const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
                var value = new string((secret ?? "").Where(c => !Char.IsWhiteSpace(c)).ToArray()).ToUpperInvariant().TrimEnd('=');
                if (value.Length < 16 || value.Length > 256 || value.Length % 8 is 1 or 3 or 6)
                    throw new FirstTradeException("invalid firsttrade_totp_secret; expected a Base32 authenticator key, not a code or recovery code");
                var key = new byte[value.Length * 5 / 8];
                int buffer = 0, bits = 0, index = 0;
                foreach (var c in value)
                {
                    var digit = alphabet.IndexOf(c);
                    if (digit < 0)
                    {
                        CryptographicOperations.ZeroMemory(key);
                        throw new FirstTradeException("invalid firsttrade_totp_secret Base32 format");
                    }
                    buffer = (buffer << 5) | digit;
                    bits += 5;
                    if (bits >= 8)
                    {
                        bits -= 8;
                        key[index++] = (byte)(buffer >> bits);
                        buffer &= (1 << bits) - 1;
                    }
                }
                if (buffer != 0)
                {
                    CryptographicOperations.ZeroMemory(key);
                    throw new FirstTradeException("invalid firsttrade_totp_secret trailing bits");
                }
                return key;
            }

            internal static string CreateTotpCode(byte[] key, DateTimeOffset time)
            {
                var seconds = time.ToUnixTimeSeconds();
                if (seconds < 0) throw new FirstTradeException("invalid system time for TOTP");
                Span<byte> counter = stackalloc byte[8];
                BinaryPrimitives.WriteInt64BigEndian(counter, seconds / 30);
                Span<byte> hash = stackalloc byte[20];
                HMACSHA1.HashData(key, counter, hash);
                var offset = hash[^1] & 15;
                var number = BinaryPrimitives.ReadUInt32BigEndian(hash.Slice(offset, 4)) & 0x7fffffff;
                CryptographicOperations.ZeroMemory(hash);
                return (number % 1000000).ToString("D6", CultureInfo.InvariantCulture);
            }

            public void Dispose()
            {
                client.Dispose();
                CryptographicOperations.ZeroMemory(totpKey);
                ftat = sid = null;
                session.Ftat = session.Sid = null;
                session.Cookies.Clear();
            }
            private sealed class FirstTradeSessionState
            {
                public int Version { get; set; } = 1;
                public string? Ftat { get; set; }
                public string? Sid { get; set; }
                public DateTimeOffset NextLoginUtc { get; set; }
                public DateTimeOffset BlockedUntilUtc { get; set; }
                public int Failures { get; set; }
                public List<FirstTradeCookie> Cookies { get; set; } = [];
            }
            private sealed record FirstTradeCookie(string Name, string Value, string Path, string Domain, bool Secure, bool HttpOnly, DateTime Expires);
            private sealed class FirstTradeSessionExpiredException : Exception;
        }
    }
}
