using System.ComponentModel;
using System.Globalization;

namespace MyBook
{
    public sealed class LiveHoldingsViewModel : INotifyPropertyChanged
    {
        readonly List<Holding> holdings;
        readonly IReadOnlyDictionary<int, string> accountNames;
        readonly IReadOnlyDictionary<CurrencyType, decimal> exchangeRates;
        IReadOnlyDictionary<(string Code, HoldingType HoldingType), MarketPrice> prices =
            new Dictionary<(string, HoldingType), MarketPrice>();
        bool showAccounts;

        internal LiveHoldingsViewModel(List<Holding> holdings, IReadOnlyDictionary<int, string> accountNames,
            IReadOnlyDictionary<CurrencyType, decimal> exchangeRates)
        {
            this.holdings = holdings;
            this.accountNames = accountNames;
            this.exchangeRates = exchangeRates;
            RebuildRows();
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public List<LiveHoldingRow> Rows { get; private set; } = [];
        public int ModeIndex
        {
            get => ShowAccounts ? 1 : 0;
            set { if (value is 0 or 1) ShowAccounts = value == 1; }
        }
        public bool ShowAccounts
        {
            get => showAccounts;
            set
            {
                if (showAccounts == value) return;
                showAccounts = value;
                RebuildRows();
                PropertyChanged?.Invoke(this, new(nameof(ShowAccounts)));
                PropertyChanged?.Invoke(this, new(nameof(ModeIndex)));
            }
        }

        internal void UpdatePrices(IReadOnlyDictionary<(string Code, HoldingType HoldingType), MarketPrice> quotes)
        {
            // The existing UI timer polls each second; unchanged quotes must not reset selection or sorting.
            if (prices.Count == quotes.Count && quotes.All(pair =>
                    prices.TryGetValue(pair.Key, out var previous) && previous == pair.Value))
            {
                RefreshQuoteAges();
                return;
            }
            prices = quotes;
            RebuildRows();
        }

        void RebuildRows()
        {
            Rows = holdings.GroupBy(holding => new
                {
                    AccountId = ShowAccounts ? holding._account_Id : (int?)null,
                    Code = holding.holdingType == HoldingType.Crypto ? KrakenPubUtil.GetBaseAsset(holding.code) : holding.code,
                    holding.holdingType,
                    Currency = holding.currentPrice.t
                })
                .Select(group =>
                {
                    prices.TryGetValue((group.Key.Code, group.Key.holdingType), out var quote);
                    if (quote is not null && (quote.Price <= 0 || quote.Currency != group.Key.Currency)) quote = null;
                    var quantity = group.Sum(holding => holding.quantity);
                    var bookValue = group.Sum(holding => holding.totalPrice.v);
                    // Weight original unit prices, not rounded book totals; a net-zero position has no average unit price.
                    decimal? bookPrice = quantity == 0 ? null : group.Sum(holding => holding.quantity * holding.currentPrice.v) / quantity;
                    // Sum the same per-holding valuations in both modes, preserving monetary rounding.
                    decimal? value = quote is null ? null : group.Sum(holding => Holding.CalculateTotalValue(holding.quantity, quote.Price, holding.holdingType));
                    decimal? valueRmb = group.Key.Currency == CurrencyType.RMB ? value
                        : exchangeRates.TryGetValue(group.Key.Currency, out var rate) && rate > 0 ? value * rate : null;
                    return new LiveHoldingRow(
                        group.Key.AccountId is int accountId ? accountNames[accountId] : "所有账户",
                        group.Key.Code, group.Key.holdingType, quantity, group.Key.Currency,
                        quote?.Price, value, bookPrice, bookValue, valueRmb, quote?.FetchedAt);
                })
                .OrderByDescending(row => row.ValueRmb)
                .ThenBy(row => row.Code, StringComparer.Ordinal)
                .ThenBy(row => row.AccountName, StringComparer.Ordinal)
                .ThenBy(row => row.HoldingType)
                .ThenBy(row => row.Currency)
                .ToList();
            RefreshQuoteAges();
            PropertyChanged?.Invoke(this, new(nameof(Rows)));
        }

        void RefreshQuoteAges()
        {
            var now = DateTimeOffset.Now;
            foreach (var row in Rows) row.RefreshQuoteAge(now);
        }
    }

    public sealed record LiveHoldingRow(string AccountName, string Code, HoldingType HoldingType,
        decimal Quantity, CurrencyType Currency, decimal? Price, decimal? Value,
        decimal? BookPrice, decimal BookValue, decimal? ValueRmb, DateTimeOffset? FetchedAt) : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        string Symbol => Currency switch
        {
            CurrencyType.GBP => "£",
            CurrencyType.EUR => "€",
            _ => CurrencySummaryViewModel.FormatCurrencySymbol(Currency)
        };
        public string QuantityText => Quantity.ToString("0.##", CultureInfo.CurrentCulture);
        public string PriceText => Price is decimal price ? $"{Symbol}{price:N2}" : "—";
        public string ValueText => Value is decimal value ? $"{Symbol}{value:N0}" : "—";
        public decimal? PriceChange => Price - BookPrice;
        public decimal? ValueChange => Value - BookValue;
        public string PriceChangeText => FormatChange(PriceChange, "N2");
        public string ValueChangeText => FormatChange(ValueChange, "N0");
        public string PriceWithChangeText => PriceText + PriceChangeText;
        public string ValueWithChangeText => ValueText + ValueChangeText;
        public string PriceChangeColor => ChangeColor(PriceChange);
        public string ValueChangeColor => ChangeColor(ValueChange);
        public string FetchedAtText { get; private set; } = "—";

        internal void RefreshQuoteAge(DateTimeOffset now)
        {
            var text = FetchedAt is DateTimeOffset fetched ? $"~{Math.Max(0, (long)(now - fetched).TotalMinutes)}min" : "—";
            if (FetchedAtText == text) return;
            FetchedAtText = text;
            PropertyChanged?.Invoke(this, new(nameof(FetchedAtText)));
        }

        string FormatChange(decimal? change, string format) => change is decimal value
            ? $" ({(value > 0 ? "↑" : value < 0 ? "↓" : "")}{Math.Abs(value).ToString(format, CultureInfo.CurrentCulture)})" : "";

        static string ChangeColor(decimal? change) => change > 0 ? "#047857" : change < 0 ? "#B91C1C" : "#64748B";
    }
}
