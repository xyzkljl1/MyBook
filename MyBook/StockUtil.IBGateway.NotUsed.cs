using MailKit.Net.Proxy;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Globalization;
using System.Threading;
using System.Text.RegularExpressions;
using IBApi;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyBook
{

    // IB Gateway 来源：历史股票、债券和持仓获取实现，当前不使用，仅保留参考。
    [Obsolete("Not used. Prefer IBKR activity report mail parsing instead.")]
    class StockUtilIBGatewayNotUsed
    {
        private const string IbGatewayHost = "127.0.0.1";
        private const int DefaultIbGatewayPort = 5679;
        private const int IbGatewayClientId = 5700;
        private const int IbGatewayDelayedMarketDataType = 3;
        private static readonly TimeSpan IbGatewayReadyTimeout = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan IbGatewayPriceTimeout = TimeSpan.FromSeconds(12);
        private static readonly TimeSpan IbGatewayPositionsTimeout = TimeSpan.FromSeconds(20);
        private static readonly TimeSpan IbGatewayContractDetailsTimeout = TimeSpan.FromSeconds(8);
        private static int nextIbGatewayRequestId = 0;

        private readonly int ibGatewayPort;
        private readonly DatabaseUtil? database;
        public StockUtilIBGatewayNotUsed(IConfigurationRoot config, DatabaseUtil? database = null)
        {
            // 为了支持gbk编码
            System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            if (!int.TryParse(config["ib_gateway_port"] ?? config["tws_port"], out ibGatewayPort))
                ibGatewayPort = DefaultIbGatewayPort;
            this.database = database;
        }
        public async Task<Currency?> Fetch(Holding holding)
        {
            Currency? ret = null;
            switch(holding.holdingType)
            {
                case HoldingType.NASDAQ:
                    ret = new Currency(await FetchIbGatewayStock(holding.code, "NASDAQ"), CurrencyType.USD);
                    break;
                case HoldingType.ARCA:
                    ret = new Currency(await FetchIbGatewayStock(holding.code, "ARCA"), CurrencyType.USD);
                    break;
                case HoldingType.UST:
                    ret = new Currency(await FetchIbGatewayBond(holding.code), CurrencyType.USD);
                    break;
                case HoldingType.SHANGHAI:
                    ret = new Currency(await FetchShanghaiStock(holding.code), CurrencyType.RMB);
                    break;
                case HoldingType.CNFUND:
                    ret = new Currency(await FetchCNFund(holding.code), CurrencyType.RMB);
                    break;
                case HoldingType.Cash:
                    ret = await FetchCurrencyToRmb(holding.currentPrice.t);
                    break;
                case HoldingType.Accrued:
                    break;
            }
            ret = ret==null||ret.v < 0 ? null : ret;
            if (ret is not null)
                holding.currentPrice = ret;
            return ret;
        }
        public Task<decimal> FetchIbGatewayStock(string code, string primaryExchange = "NASDAQ")
        {
            var symbol = code.Trim().ToUpperInvariant();
            var exchange = primaryExchange.Trim().ToUpperInvariant();
            var contract = new Contract
            {
                Symbol = symbol,
                SecType = "STK",
                Exchange = "SMART",
                PrimaryExch = exchange,
                Currency = "USD"
            };
            return FetchIbGatewayLatestPrice(contract, $"{symbol}:{exchange}");
        }
        public Task<decimal> FetchIbGatewayBond(string code)
        {
            var contract = CreateIbGatewayBondContract(code);
            var bondId = code.Trim().ToUpperInvariant();
            return FetchIbGatewayLatestPrice(contract, bondId);
        }
        private static Contract CreateIbGatewayBondContract(string code)
        {
            var bondId = code.Trim().ToUpperInvariant();
            var contract = new Contract
            {
                SecType = "BOND",
                Exchange = "SMART",
                Currency = "USD"
            };

            if (Int32.TryParse(bondId, out var conId))
            {
                contract.ConId = conId;
            }
            else if (Regex.IsMatch(bondId, @"^[A-Z]{2}[A-Z0-9]{9}\d$"))
            {
                contract.SecIdType = "ISIN";
                contract.SecId = bondId;
            }
            else
            {
                contract.Symbol = bondId;
            }

            return contract;
        }
        private async Task<decimal> FetchIbGatewayLatestPrice(Contract contract, string desc)
        {
            var tickerId = Interlocked.Increment(ref nextIbGatewayRequestId);
            var wrapper = new IbGatewayPriceWrapper(tickerId, desc);
            var signal = new EReaderMonitorSignal();
            var client = new EClientSocket(wrapper, signal);
            Thread? readerThread = null;

            try
            {
                client.eConnect(IbGatewayHost, ibGatewayPort, IbGatewayClientId + tickerId);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}: {e.Message}");
                return -1;
            }

            if (!client.IsConnected())
            {
                Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}");
                return -1;
            }

            try
            {
                var reader = new EReader(client, signal);
                reader.Start();
                readerThread = new Thread(() =>
                {
                    while (client.IsConnected() && !wrapper.IsCompleted)
                    {
                        signal.waitForSignal();
                        reader.processMsgs();
                    }
                });
                readerThread.IsBackground = true;
                readerThread.Start();

                client.startApi();
                var readyTask = await Task.WhenAny(wrapper.ReadyTask, Task.Delay(IbGatewayReadyTimeout));
                if (readyTask != wrapper.ReadyTask)
                {
                    Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}: timeout waiting for API ready");
                    return -1;
                }

                client.reqMarketDataType(IbGatewayDelayedMarketDataType);
                client.reqMktData(tickerId, contract, "", false, false, new List<TagValue>());

                var completedTask = await Task.WhenAny(wrapper.PriceTask, Task.Delay(IbGatewayPriceTimeout));
                if (completedTask == wrapper.PriceTask)
                {
                    var price = await wrapper.PriceTask;
                    if (price > 0)
                        Console.WriteLine($"{desc}:{price}");
                    return price;
                }

                var fallbackPrice = wrapper.FallbackPrice;
                if (fallbackPrice.HasValue)
                {
                    Console.WriteLine($"{desc}:{fallbackPrice.Value}");
                    return fallbackPrice.Value;
                }

                Console.WriteLine($"fail to fetch IB Gateway price {desc}: timeout");
                return -1;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch IB Gateway price {desc}: {e.Message}");
                return -1;
            }
            finally
            {
                try
                {
                    if (client.IsConnected())
                    {
                        client.cancelMktData(tickerId);
                        client.eDisconnect(false);
                    }
                }
                catch
                {
                }

                signal.issueSignal();
                if (readerThread is not null && readerThread.IsAlive)
                    readerThread.Join(1000);
            }
        }
        private async Task<string?> FetchIbGatewayBondDisplayText(string code)
        {
            var requestId = Interlocked.Increment(ref nextIbGatewayRequestId);
            var wrapper = new IbGatewayContractDetailsWrapper(requestId, code);
            var signal = new EReaderMonitorSignal();
            var client = new EClientSocket(wrapper, signal);
            Thread? readerThread = null;

            try
            {
                client.eConnect(IbGatewayHost, ibGatewayPort, IbGatewayClientId + requestId);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}: {e.Message}");
                return null;
            }

            if (!client.IsConnected())
            {
                Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}");
                return null;
            }

            try
            {
                var reader = new EReader(client, signal);
                reader.Start();
                readerThread = new Thread(() =>
                {
                    while (client.IsConnected() && !wrapper.IsCompleted)
                    {
                        signal.waitForSignal();
                        reader.processMsgs();
                    }
                });
                readerThread.IsBackground = true;
                readerThread.Start();

                client.startApi();
                var readyTask = await Task.WhenAny(wrapper.ReadyTask, Task.Delay(IbGatewayReadyTimeout));
                if (readyTask != wrapper.ReadyTask)
                {
                    Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}: timeout waiting for API ready");
                    return null;
                }

                client.reqContractDetails(requestId, CreateIbGatewayBondContract(code));
                var completedTask = await Task.WhenAny(wrapper.DisplayTextTask, Task.Delay(IbGatewayContractDetailsTimeout));
                if (completedTask == wrapper.DisplayTextTask)
                    return await wrapper.DisplayTextTask;

                Console.WriteLine($"fail to fetch IB Gateway contract details {code}: timeout");
                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch IB Gateway contract details {code}: {e.Message}");
                return null;
            }
            finally
            {
                try
                {
                    if (client.IsConnected())
                        client.eDisconnect(false);
                }
                catch
                {
                }

                signal.issueSignal();
                if (readerThread is not null && readerThread.IsAlive)
                    readerThread.Join(1000);
            }
        }
        public async Task<List<Holding>> Fetch(Account account)
        {
            if (!account.name.Contains("IBKR", StringComparison.OrdinalIgnoreCase))
                return new List<Holding>();

            var holdings = await FetchIbGatewayPositions(account);
            if (database is null)
            {
                Console.WriteLine("skip saving IB Gateway positions: database is not configured for StockUtil");
                return holdings;
            }

            database.SaveAccountHoldings(account, holdings);
            return holdings;
        }
        private async Task<List<Holding>> FetchIbGatewayPositions(Account account)
        {
            var requestId = Interlocked.Increment(ref nextIbGatewayRequestId);
            var wrapper = new IbGatewayPositionsWrapper(account);
            var signal = new EReaderMonitorSignal();
            var client = new EClientSocket(wrapper, signal);
            Thread? readerThread = null;

            try
            {
                client.eConnect(IbGatewayHost, ibGatewayPort, IbGatewayClientId + requestId);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}: {e.Message}");
                return new List<Holding>();
            }

            if (!client.IsConnected())
            {
                Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}");
                return new List<Holding>();
            }

            try
            {
                var reader = new EReader(client, signal);
                reader.Start();
                readerThread = new Thread(() =>
                {
                    while (client.IsConnected() && !wrapper.IsCompleted)
                    {
                        signal.waitForSignal();
                        reader.processMsgs();
                    }
                });
                readerThread.IsBackground = true;
                readerThread.Start();

                client.startApi();
                var readyTask = await Task.WhenAny(wrapper.ReadyTask, Task.Delay(IbGatewayReadyTimeout));
                if (readyTask != wrapper.ReadyTask)
                {
                    Console.WriteLine($"fail to connect IB Gateway {IbGatewayHost}:{ibGatewayPort}: timeout waiting for API ready");
                    return new List<Holding>();
                }

                client.reqPositions();
                var completedTask = await Task.WhenAny(wrapper.PositionsTask, Task.Delay(IbGatewayPositionsTimeout));
                if (completedTask != wrapper.PositionsTask)
                    Console.WriteLine("fail to fetch IB Gateway positions: timeout");

                foreach (var holding in wrapper.Holdings.Where(it => it.holdingType == HoldingType.UST))
                    holding.displayText = await FetchIbGatewayBondDisplayText(holding.code) ?? holding.displayText;

                return wrapper.Holdings;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch IB Gateway positions: {e.Message}");
                return new List<Holding>();
            }
            finally
            {
                try
                {
                    if (client.IsConnected())
                    {
                        client.cancelPositions();
                        client.eDisconnect(false);
                    }
                }
                catch
                {
                }

                signal.issueSignal();
                if (readerThread is not null && readerThread.IsAlive)
                    readerThread.Join(1000);
            }
        }
        public async Task<Currency?> FetchCurrencyToRmb(CurrencyType currencyType)
        {
            if (currencyType == CurrencyType.RMB)
                return new Currency(1, CurrencyType.RMB);

            var fromCurrency = currencyType.ToString();
            var url = $"https://www.google.com/finance/quote/{fromCurrency}-CNY?hl=en";
            try
            {
                var html = await HttpGetString(url);
                var rate = String.IsNullOrWhiteSpace(html) ? null : ParseGoogleFinanceExchangeRate(html, fromCurrency);
                if (rate is null)
                    return null;

                Console.WriteLine($"{fromCurrency}/CNY:{rate.Value}");
                return new Currency(rate.Value, CurrencyType.RMB);
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch currency exchange rate {fromCurrency}/CNY: {e}");
            }
            return null;
        }
        private static decimal? ParseGoogleFinanceExchangeRate(string html, string fromCurrency)
        {
            var match = Regex.Match(html, $@"""{Regex.Escape(fromCurrency)} / CNY""\s*,\s*3\s*,\s*null\s*,\s*\[(?<rate>\d+(?:\.\d+)?)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                return null;

            return decimal.TryParse(match.Groups["rate"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var rate) ? rate : null;
        }
        // Google Finance 会同时返回常规交易价格和盘前/盘后价格；如有扩展时段价格，优先使用扩展时段价格。
        public async Task<decimal> FetchGoogleFinanceStock(string code, string exchange = "NASDAQ")
        {
            var symbol = code.Trim().ToUpperInvariant();
            var market = exchange.Trim().ToUpperInvariant();
            var url = $"https://www.google.com/finance/quote/{Uri.EscapeDataString(symbol)}:{market}?hl=en";
            try
            {
                var html = await HttpGetString(url);
                if (String.IsNullOrWhiteSpace(html))
                    return -1;

                var price = ParseGoogleFinanceStockPrice(html, symbol, market);
                if (price is null)
                    return -1;

                Console.WriteLine($"{symbol}:{market}:{price.Value}");
                return price.Value;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {symbol}:{market}: {e}");
            }
            return -1;
        }
        private static decimal? ParseGoogleFinanceStockPrice(string html, string symbol, string exchange)
        {
            var entityMatch = Regex.Match(
                html,
                $@"\[""[^""]+"",\[""{Regex.Escape(symbol)}"",""{Regex.Escape(exchange)}""\],""[^""]+"",\d+,""[A-Z]{{3}}"",\[-?\d+(?:\.\d+)?",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!entityMatch.Success)
                return null;

            var symbolMarker = $@"""{symbol}:{exchange}""";
            var entityEnd = html.IndexOf(symbolMarker, entityMatch.Index, StringComparison.OrdinalIgnoreCase);
            var entityLength = entityEnd > entityMatch.Index ? entityEnd - entityMatch.Index : Math.Min(1200, html.Length - entityMatch.Index);
            var entity = html.Substring(entityMatch.Index, entityLength);
            var quoteMatches = Regex.Matches(entity, @"\[(?<price>-?\d+(?:\.\d+)?),-?\d+(?:\.\d+)?,-?\d+(?:\.\d+)?,\d+,\d+,\d+\]");
            if (quoteMatches.Count == 0)
                return null;

            var match = quoteMatches[quoteMatches.Count - 1];
            return decimal.TryParse(match.Groups["price"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) ? price : null;
        }
        // 返回小于0表示错误。新浪行情字段 3 是当前价，停牌或未开盘时可能为 0，此时回退到昨收。
        public async Task<decimal> FetchShanghaiStock(string code)
        {
            var symbol = $"sh{code.Trim().ToLowerInvariant().Replace("sh", "")}";
            var url = $"https://hq.sinajs.cn/list={symbol}";
            try
            {
                var text = await HttpGetString(url);
                if (String.IsNullOrWhiteSpace(text))
                    return -1;

                var match = Regex.Match(text, "=\"(?<data>[^\"]*)\"");
                if (!match.Success)
                    return -1;

                var fields = match.Groups["data"].Value.Split(',');
                if (fields.Length < 4)
                    return -1;

                var priceText = fields[3];
                if (!Decimal.TryParse(priceText, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) || price <= 0)
                    Decimal.TryParse(fields[2], NumberStyles.Number, CultureInfo.InvariantCulture, out price);

                if (price <= 0)
                    return -1;

                Console.WriteLine($"{fields[0]}({symbol}):{price}");
                return price;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {code}: {e}");
            }
            return -1;
        }
        public async Task<decimal> FetchCNFund(string code = "021282")
        {
            var url = $"http://fundgz.1234567.com.cn/js/{code}.js";
            try
            {
                var doc = await HttpGetJson(url);
                if (doc is null)
                    return -1;
                var priceText = doc["gsz"]?.ToString();
                if (String.IsNullOrWhiteSpace(priceText) || !Decimal.TryParse(priceText, NumberStyles.Number, CultureInfo.InvariantCulture, out var price) || price <= 0)
                    price = decimal.Parse(doc["dwjz"]!.ToString()!, CultureInfo.InvariantCulture);
                var name = doc["name"]!.ToString()!;
                var date = doc["gztime"]!.ToString();
                Console.WriteLine($"{name}({date}):{price}");
                return price;
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch stock {code}: {e}");
            }
            return 0;
        }
        public async Task<JObject?> HttpGetJson(string url)
        {
            try
            {
                using (HttpClient client = new())
                using (HttpResponseMessage response = await client.GetAsync(url))
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                        return null;
                    var text = await response.Content.ReadAsStringAsync();
                    // 只选最外层的大括号里面的东西
                    text = text.Substring(text.IndexOf('{'), text.LastIndexOf('}') - text.IndexOf('{')+1);
                    return (JObject?)JsonConvert.DeserializeObject(text);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch {url}: {e}");
            }
            return null;
        }
        public async Task<string?> HttpGetString(string url)
        {
            try
            {
                using (HttpClient client = new())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36");
                    client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.9");
                    using (HttpResponseMessage response = await client.GetAsync(url))
                    {
                        if (response.StatusCode != HttpStatusCode.OK)
                            return null;
                        return await response.Content.ReadAsStringAsync();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"fail to fetch {url}: {e}");
            }
            return null;
        }
        private sealed class IbGatewayPriceWrapper : DefaultEWrapper
        {
            private const int TickTypeBid = 1;
            private const int TickTypeAsk = 2;
            private const int TickTypeLast = 4;
            private const int TickTypeClose = 9;
            private const int TickTypeMarkPrice = 37;
            private const int TickTypeDelayedBid = 66;
            private const int TickTypeDelayedAsk = 67;
            private const int TickTypeDelayedLast = 68;
            private const int TickTypeDelayedClose = 75;

            private readonly int tickerId;
            private readonly string desc;
            private readonly TaskCompletionSource<bool> readyResult = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly TaskCompletionSource<decimal> priceResult = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private decimal? bidPrice;
            private decimal? askPrice;
            private decimal? fallbackPrice;

            public IbGatewayPriceWrapper(int tickerId, string desc)
            {
                this.tickerId = tickerId;
                this.desc = desc;
            }

            public Task<bool> ReadyTask => readyResult.Task;
            public Task<decimal> PriceTask => priceResult.Task;
            public bool IsCompleted => priceResult.Task.IsCompleted;
            public decimal? FallbackPrice
            {
                get
                {
                    if (fallbackPrice.HasValue)
                        return fallbackPrice;
                    if (bidPrice.HasValue && askPrice.HasValue)
                        return (bidPrice.Value + askPrice.Value) / 2;
                    return bidPrice ?? askPrice;
                }
            }

            public override void tickPrice(int tickerId, int field, double price, TickAttrib attribs)
            {
                if (tickerId != this.tickerId || price <= 0 || Double.IsNaN(price))
                    return;

                var value = (decimal)price;
                switch (field)
                {
                    case TickTypeLast:
                    case TickTypeDelayedLast:
                        priceResult.TrySetResult(value);
                        break;
                    case TickTypeBid:
                    case TickTypeDelayedBid:
                        bidPrice = value;
                        break;
                    case TickTypeAsk:
                    case TickTypeDelayedAsk:
                        askPrice = value;
                        break;
                    case TickTypeMarkPrice:
                    case TickTypeClose:
                    case TickTypeDelayedClose:
                        fallbackPrice = value;
                        break;
                }
            }

            public override void nextValidId(int orderId)
            {
                readyResult.TrySetResult(true);
            }

            public override void managedAccounts(string accountsList)
            {
                readyResult.TrySetResult(true);
            }

            public override void connectAck()
            {
            }

            public override void error(Exception e)
            {
                if (IsIbGatewayDisconnectMessage(e.Message))
                    return;

                Console.WriteLine($"IB Gateway error {desc}: {e.Message}");
            }

            public override void error(string str)
            {
                if (IsIbGatewayDisconnectMessage(str))
                    return;

                Console.WriteLine($"IB Gateway error {desc}: {str}");
            }

            public override void error(int id, int errorCode, string errorMsg)
            {
                if (id < 0 && IsIbGatewayStatusCode(errorCode))
                    return;

                Console.WriteLine($"IB Gateway error {desc}: {errorCode} {errorMsg}");
                if (id == tickerId && IsFatalMarketDataError(errorCode))
                    priceResult.TrySetResult(-1);
            }

            public override void connectionClosed()
            {
                priceResult.TrySetResult(-1);
            }

            private static bool IsFatalMarketDataError(int errorCode)
            {
                return errorCode is 200 or 321 or 354 or 420 or 502 or 503 or 504;
            }
        }

        private sealed class IbGatewayContractDetailsWrapper : DefaultEWrapper
        {
            private readonly int requestId;
            private readonly string fallbackText;
            private readonly TaskCompletionSource<bool> readyResult = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly TaskCompletionSource<string?> displayTextResult = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private string? displayText;

            public IbGatewayContractDetailsWrapper(int requestId, string fallbackText)
            {
                this.requestId = requestId;
                this.fallbackText = fallbackText;
            }

            public Task<bool> ReadyTask => readyResult.Task;
            public Task<string?> DisplayTextTask => displayTextResult.Task;
            public bool IsCompleted => displayTextResult.Task.IsCompleted;

            public override void nextValidId(int orderId)
            {
                readyResult.TrySetResult(true);
            }

            public override void managedAccounts(string accountsList)
            {
                readyResult.TrySetResult(true);
            }

            public override void contractDetails(int reqId, ContractDetails contractDetails)
            {
                HandleContractDetails(reqId, contractDetails);
            }

            public override void bondContractDetails(int reqId, ContractDetails contractDetails)
            {
                HandleContractDetails(reqId, contractDetails);
            }

            private void HandleContractDetails(int reqId, ContractDetails contractDetails)
            {
                if (reqId != requestId)
                    return;

                displayText = BuildDisplayText(contractDetails) ?? displayText;
            }

            public override void contractDetailsEnd(int reqId)
            {
                if (reqId == requestId)
                    displayTextResult.TrySetResult(displayText ?? fallbackText);
            }

            public override void error(Exception e)
            {
                if (IsIbGatewayDisconnectMessage(e.Message))
                    return;

                Console.WriteLine($"IB Gateway contract details error {fallbackText}: {e.Message}");
            }

            public override void error(string str)
            {
                if (IsIbGatewayDisconnectMessage(str))
                    return;

                Console.WriteLine($"IB Gateway contract details error {fallbackText}: {str}");
            }

            public override void error(int id, int errorCode, string errorMsg)
            {
                if (id < 0 && IsIbGatewayStatusCode(errorCode))
                    return;

                Console.WriteLine($"IB Gateway contract details error {fallbackText}: {errorCode} {errorMsg}");
                if (id == requestId && errorCode is 200 or 321)
                    displayTextResult.TrySetResult(displayText ?? fallbackText);
            }

            public override void connectionClosed()
            {
                displayTextResult.TrySetResult(displayText ?? fallbackText);
            }

            private static string? BuildDisplayText(ContractDetails contractDetails)
            {
                if (String.Equals(contractDetails.Contract?.SecType, "BOND", StringComparison.OrdinalIgnoreCase) ||
                    !String.IsNullOrWhiteSpace(contractDetails.Maturity) ||
                    contractDetails.Coupon > 0)
                {
                    var parts = new List<string>();
                    AddIfNotBlank(parts, FirstNonBlank(
                        contractDetails.LongName,
                        contractDetails.Contract?.Symbol,
                        contractDetails.MarketName,
                        contractDetails.Cusip));
                    AddIfNotBlank(parts, contractDetails.Coupon > 0
                        ? $"{contractDetails.Coupon.ToString("0.###", CultureInfo.InvariantCulture)}%"
                        : null);
                    AddIfNotBlank(parts, FormatIbDate(contractDetails.Maturity));
                    AddIfNotBlank(parts, contractDetails.DescAppend);
                    AddIfNotBlank(parts, contractDetails.Cusip);

                    if (parts.Count > 0)
                        return String.Join(" ", parts.Distinct(StringComparer.OrdinalIgnoreCase));
                }

                return FirstNonBlank(
                    contractDetails.LongName,
                    contractDetails.Contract?.LocalSymbol,
                    contractDetails.Contract?.Symbol);
            }

            private static void AddIfNotBlank(List<string> parts, string? value)
            {
                if (!String.IsNullOrWhiteSpace(value))
                    parts.Add(value.Trim());
            }

            private static string? FormatIbDate(string? value)
            {
                if (String.IsNullOrWhiteSpace(value))
                    return null;

                value = value.Trim();
                if (value.Length == 8 &&
                    DateTime.TryParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                {
                    return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                }

                return value;
            }

            private static string? FirstNonBlank(params string?[] values)
            {
                foreach (var value in values)
                {
                    if (!String.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }

                return null;
            }
        }

        private sealed class IbGatewayPositionsWrapper : DefaultEWrapper
        {
            private readonly Account account;
            private readonly string accountFilterText;
            private readonly TaskCompletionSource<bool> readyResult = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly TaskCompletionSource<bool> positionsResult = new(TaskCreationOptions.RunContinuationsAsynchronously);

            public IbGatewayPositionsWrapper(Account account)
            {
                this.account = account;
                accountFilterText = $"{account.name} {account.desc}";
            }

            public Task<bool> ReadyTask => readyResult.Task;
            public Task<bool> PositionsTask => positionsResult.Task;
            public bool IsCompleted => positionsResult.Task.IsCompleted;
            public List<Holding> Holdings { get; } = new();

            public override void nextValidId(int orderId)
            {
                readyResult.TrySetResult(true);
            }

            public override void managedAccounts(string accountsList)
            {
                readyResult.TrySetResult(true);
            }

            public override void position(string ibAccount, Contract contract, double pos, double avgCost)
            {
                if (pos == 0 || !ShouldUsePosition(ibAccount))
                    return;

                var holding = CreateHolding(contract, pos, avgCost);
                if (holding is not null)
                    Holdings.Add(holding);
            }

            public override void positionEnd()
            {
                positionsResult.TrySetResult(true);
            }

            public override void error(Exception e)
            {
                if (IsIbGatewayDisconnectMessage(e.Message))
                    return;

                Console.WriteLine($"IB Gateway positions error: {e.Message}");
            }

            public override void error(string str)
            {
                if (IsIbGatewayDisconnectMessage(str))
                    return;

                Console.WriteLine($"IB Gateway positions error: {str}");
            }

            public override void error(int id, int errorCode, string errorMsg)
            {
                if (id < 0 && IsIbGatewayStatusCode(errorCode))
                    return;

                Console.WriteLine($"IB Gateway positions error: {errorCode} {errorMsg}");
            }

            public override void connectionClosed()
            {
                positionsResult.TrySetResult(true);
            }

            private bool ShouldUsePosition(string ibAccount)
            {
                if (accountFilterText.Contains(ibAccount, StringComparison.OrdinalIgnoreCase))
                    return true;

                var last4 = ibAccount.Length >= 4 ? ibAccount[^4..] : ibAccount;
                if (accountFilterText.Contains(last4, StringComparison.OrdinalIgnoreCase))
                    return true;

                return !Regex.IsMatch(accountFilterText, @"\d{4,}");
            }

            private Holding? CreateHolding(Contract contract, double pos, double avgCost)
            {
                var holdingType = GetHoldingType(contract);
                if (holdingType is null)
                    return null;

                var code = GetStockCode(contract, holdingType.Value);
                if (String.IsNullOrWhiteSpace(code))
                    return null;

                var currencyType = Enum.TryParse<CurrencyType>(contract.Currency, out var parsedCurrency)
                    ? parsedCurrency
                    : CurrencyType.USD;
                var exchange = !String.IsNullOrWhiteSpace(contract.PrimaryExch)
                    ? contract.PrimaryExch
                    : contract.Exchange;
                var displayText = holdingType.Value == HoldingType.UST
                    ? FirstNonBlank(contract.LocalSymbol, contract.Symbol, code)
                    : code;

                return new Holding
                {
                    Account = account,
                    _account_Id = account.Id,
                    code = code,
                    holdingType = holdingType.Value,
                    quantity = ToHoldingQuantity(pos),
                    desc = $"IB Gateway {contract.SecType} {exchange} avgCost={avgCost.ToString(CultureInfo.InvariantCulture)}",
                    displayText = displayText,
                    _currentPrice_t = currencyType
                };
            }

            private static int ToHoldingQuantity(double quantity)
            {
                var rounded = Math.Round(quantity);
                if (Math.Abs(quantity - rounded) > 0.0000001)
                    throw new InvalidOperationException($"IB Gateway holding quantity is not an integer: {quantity.ToString(CultureInfo.InvariantCulture)}");

                return checked((int)rounded);
            }

            private static HoldingType? GetHoldingType(Contract contract)
            {
                if (String.Equals(contract.SecType, "BOND", StringComparison.OrdinalIgnoreCase))
                    return HoldingType.UST;

                if (!String.Equals(contract.SecType, "STK", StringComparison.OrdinalIgnoreCase))
                    return null;

                var exchange = $"{contract.PrimaryExch} {contract.Exchange}";
                if (exchange.Contains("ARCA", StringComparison.OrdinalIgnoreCase))
                    return HoldingType.ARCA;

                return HoldingType.NASDAQ;
            }

            private static string GetStockCode(Contract contract, HoldingType holdingType)
            {
                if (holdingType == HoldingType.UST)
                {
                    if (contract.ConId > 0)
                        return contract.ConId.ToString(CultureInfo.InvariantCulture);
                    if (!String.IsNullOrWhiteSpace(contract.SecId))
                        return contract.SecId.Trim().ToUpperInvariant();
                }

                if (!String.IsNullOrWhiteSpace(contract.LocalSymbol))
                    return contract.LocalSymbol.Trim().ToUpperInvariant();
                return contract.Symbol?.Trim().ToUpperInvariant() ?? "";
            }

            private static string FirstNonBlank(params string?[] values)
            {
                foreach (var value in values)
                {
                    if (!String.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }

                return "";
            }
        }

        private static bool IsIbGatewayStatusCode(int errorCode)
        {
            return errorCode is 2104 or 2105 or 2106 or 2119 or 2157 or 2158;
        }

        private static bool IsIbGatewayDisconnectMessage(string? message)
        {
            if (String.IsNullOrWhiteSpace(message))
                return false;

            return message.Contains("Unable to read data from the transport connection", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("Unable to read beyond the end of the stream", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("你的主机中的软件中止", StringComparison.OrdinalIgnoreCase);
        }
    }
}
