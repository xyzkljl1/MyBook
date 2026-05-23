using Microsoft.Extensions.Configuration;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace MyBook
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        readonly Fetcher fetcher = new();
        Forms.NotifyIcon? trayIcon;
        Drawing.Icon? trayIconImage;
        bool isExitRequested;

        public MainWindow()
        {
#if DEBUG
            AllocConsole();
#endif
            InitializeComponent();
            InitializeTrayIcon();
            fetcher.RunSchedule();
            _ = LoadDashboardAsync();
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e)
        {
            await LoadDashboardAsync();
        }

        private async Task LoadDashboardAsync()
        {
            try
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var data = database.GetDashboardData(DateTime.Today);
                if (data.MissingExchangeRateCurrencies.Count > 0)
                {
                    var stock = new StockUtil(config, database);
                    await stock.FetchExchangeRates(data.MissingExchangeRateCurrencies);
                    data = database.GetDashboardData(DateTime.Today);
                    if (data.MissingExchangeRateCurrencies.Count > 0)
                    {
                        throw new InvalidOperationException(
                            $"缺少汇率：{String.Join("、", data.MissingExchangeRateCurrencies)}");
                    }
                }

                DataContext = DashboardViewModel.From(data);
            }
            catch (Exception e)
            {
                DataContext = DashboardViewModel.FromError(e.Message);
            }
        }

        private void InitializeTrayIcon()
        {
            var menu = new Forms.ContextMenuStrip();
            menu.Items.Add("退出", null, (_, _) => ExitFromTray());
            trayIcon = new Forms.NotifyIcon
            {
                Text = "MyBook",
                Icon = trayIconImage = LoadAppIcon(),
                ContextMenuStrip = menu,
                Visible = true
            };
            trayIcon.MouseClick += (_, e) =>
            {
                if (e.Button == Forms.MouseButtons.Left)
                    RestoreFromTray();
            };
        }

        private static Drawing.Icon LoadAppIcon()
        {
            var iconResource = Application.GetResourceStream(new Uri("pack://application:,,,/Assets/AppIcon.ico"));
            if (iconResource?.Stream is null)
                return Drawing.SystemIcons.Application;

            using var stream = iconResource.Stream;
            using var icon = new Drawing.Icon(stream);
            return (Drawing.Icon)icon.Clone();
        }

        private void RestoreFromTray()
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        }

        private void ExitFromTray()
        {
            isExitRequested = true;
            if (trayIcon is not null)
                trayIcon.Visible = false;
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!isExitRequested)
            {
                e.Cancel = true;
                Hide();
                return;
            }

            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            trayIcon?.Dispose();
            trayIconImage?.Dispose();
            fetcher.Dispose();
            base.OnClosed(e);
        }
    }

    public class DashboardViewModel : INotifyPropertyChanged
    {
        bool showSingleCurrencyMonthly;
        double selectedReasonMonthIndex;
        bool showInvestmentByHolding;
        InvestmentAccountStatisticsViewModel? selectedInvestmentAccount;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string SnapshotTimeText { get; set; } = "";
        public string ReasonTabHeader { get; set; } = "分类";
        public TotalAssetsViewModel TotalAssets { get; set; } = new();
        public List<CurrencySummaryViewModel> CurrencySummaries { get; set; } = [];
        public List<MonthlyFlowSeriesViewModel> MonthlySeries { get; set; } = [];
        public List<MonthlyFlowSeriesViewModel> RmbMonthlySeries { get; set; } = [];
        public List<ReasonFlowSeriesViewModel> ReasonMonthSeries { get; set; } = [];
        public List<InvestmentAccountStatisticsViewModel> InvestmentAccounts { get; set; } = [];
        public IEnumerable<MonthlyFlowSeriesViewModel> VisibleMonthlySeries => ShowSingleCurrencyMonthly ? MonthlySeries : RmbMonthlySeries;
        public IEnumerable<InvestmentStatisticsPeriodViewModel> VisibleInvestmentPeriods => SelectedInvestmentAccount is null
            ? Enumerable.Empty<InvestmentStatisticsPeriodViewModel>()
            : ShowInvestmentByHolding
                ? SelectedInvestmentAccount.ByHoldingPeriods
                : SelectedInvestmentAccount.ByReasonPeriods;
        public double ReasonMonthMaximum => Math.Max(0, ReasonMonthSeries.Count - 1);
        public ReasonFlowSeriesViewModel SelectedReasonSeries => ReasonMonthSeries.Count == 0
            ? new ReasonFlowSeriesViewModel()
            : ReasonMonthSeries[Math.Clamp((int)Math.Round(selectedReasonMonthIndex), 0, ReasonMonthSeries.Count - 1)];
        public string SelectedReasonMonthLabel => SelectedReasonSeries.MonthLabel;

        public bool ShowSingleCurrencyMonthly
        {
            get => showSingleCurrencyMonthly;
            set
            {
                if (showSingleCurrencyMonthly == value)
                    return;

                showSingleCurrencyMonthly = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ShowSingleCurrencyMonthly)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VisibleMonthlySeries)));
            }
        }

        public double SelectedReasonMonthIndex
        {
            get => selectedReasonMonthIndex;
            set
            {
                var rounded = Math.Round(value);
                var normalized = ReasonMonthSeries.Count == 0
                    ? 0
                    : Math.Clamp(rounded, 0, ReasonMonthSeries.Count - 1);
                if (selectedReasonMonthIndex == normalized)
                    return;

                selectedReasonMonthIndex = normalized;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedReasonMonthIndex)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedReasonSeries)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedReasonMonthLabel)));
            }
        }

        public bool ShowInvestmentByReason
        {
            get => !showInvestmentByHolding;
            set
            {
                if (value)
                    ShowInvestmentByHolding = false;
            }
        }

        public bool ShowInvestmentByHolding
        {
            get => showInvestmentByHolding;
            set
            {
                if (showInvestmentByHolding == value)
                    return;

                showInvestmentByHolding = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ShowInvestmentByReason)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ShowInvestmentByHolding)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VisibleInvestmentPeriods)));
            }
        }

        public InvestmentAccountStatisticsViewModel? SelectedInvestmentAccount
        {
            get => selectedInvestmentAccount;
            set
            {
                if (ReferenceEquals(selectedInvestmentAccount, value))
                    return;

                selectedInvestmentAccount = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedInvestmentAccount)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VisibleInvestmentPeriods)));
            }
        }

        public static DashboardViewModel From(DashboardData data)
        {
            var reasonSeries = data.RmbReasonFlowSeriesByMonth
                .Select(ReasonFlowSeriesViewModel.From)
                .ToList();
            List<InvestmentAccountStatistics> investmentAccounts = data.InvestmentAccounts.Count == 0
                ?
                [
                    new InvestmentAccountStatistics
                    {
                        DisplayName = "所有账户",
                        ByReason = data.InvestmentByReason,
                        ByHolding = data.InvestmentByHolding
                    }
                ]
                : data.InvestmentAccounts;
            var viewModel = new DashboardViewModel
            {
                SnapshotTimeText = $"更新于 {DateTime.Now:yyyy-MM-dd HH:mm}",
                ReasonTabHeader = "分类",
                TotalAssets = TotalAssetsViewModel.From(data.TotalAssetsRmb),
                CurrencySummaries = data.CurrencySummaries.Select(CurrencySummaryViewModel.From).ToList(),
                MonthlySeries = data.MonthlyFlowSeries
                    .Select(MonthlyFlowSeriesViewModel.From)
                    .ToList(),
                RmbMonthlySeries = [MonthlyFlowSeriesViewModel.From(data.RmbMonthlyFlowSeries)],
                ReasonMonthSeries = reasonSeries,
                InvestmentAccounts = investmentAccounts
                    .Select(InvestmentAccountStatisticsViewModel.From)
                    .ToList()
            };
            viewModel.SelectedInvestmentAccount = viewModel.InvestmentAccounts.FirstOrDefault();
            viewModel.SelectedReasonMonthIndex = Math.Clamp(data.DefaultReasonMonthIndex, 0, Math.Max(0, reasonSeries.Count - 1));
            return viewModel;
        }

        public static DashboardViewModel FromError(string message)
        {
            return new DashboardViewModel
            {
                SnapshotTimeText = $"数据读取失败：{message}"
            };
        }
    }

    public class CurrencySummaryViewModel
    {
        public string Currency { get; set; } = "";
        public string NetText { get; set; } = "";
        public string NetExactText { get; set; } = "";
        public string AssetsText { get; set; } = "";
        public string AssetsExactText { get; set; } = "";
        public string LiabilitiesText { get; set; } = "";
        public string LiabilitiesExactText { get; set; } = "";

        public static CurrencySummaryViewModel From(CurrencyBalanceSummary summary)
        {
            var symbol = FormatCurrencySymbol(summary.Currency);
            var net = FormatSummaryMoney(summary.Net, symbol);
            var shouldShowBreakdown = summary.Assets != 0 && summary.Liabilities != 0;
            var assets = shouldShowBreakdown ? FormatSignedMoney(summary.Assets, "+", symbol) : new MoneyText("", "");
            var liabilities = shouldShowBreakdown ? FormatSignedMoney(summary.Liabilities, "-", symbol) : new MoneyText("", "");
            return new CurrencySummaryViewModel
            {
                Currency = symbol,
                NetText = net.DisplayText,
                NetExactText = net.ExactText,
                AssetsText = assets.DisplayText,
                AssetsExactText = assets.ExactText,
                LiabilitiesText = liabilities.DisplayText,
                LiabilitiesExactText = liabilities.ExactText
            };
        }

        private static MoneyText FormatSummaryMoney(decimal value, string prefix)
        {
            return MoneyText.From(value, prefix);
        }

        public static string FormatCurrencySymbol(CurrencyType currency)
        {
            return currency switch
            {
                CurrencyType.RMB => "¥",
                CurrencyType.USD => "$",
                CurrencyType.JPY => "JP¥",
                CurrencyType.SGD => "S$",
                CurrencyType.HKD => "HK$",
                _ => currency.ToString()
            };
        }

        private static MoneyText FormatSignedMoney(decimal value, string sign, string prefix)
        {
            if (value == 0)
                return new MoneyText("", "");

            var text = MoneyText.From(Math.Abs(value), prefix);
            return new MoneyText(
                $"{sign}{text.DisplayText}",
                String.IsNullOrWhiteSpace(text.ExactText) ? "" : $"{sign}{text.ExactText}");
        }
    }

    public class TotalAssetsViewModel
    {
        public string ValueText { get; set; } = "";
        public string ExactText { get; set; } = "";

        public static TotalAssetsViewModel From(decimal value)
        {
            var text = MoneyText.From(value, "¥");
            return new TotalAssetsViewModel
            {
                ValueText = text.DisplayText,
                ExactText = text.ExactText
            };
        }
    }

    public readonly record struct MoneyText(string DisplayText, string ExactText)
    {
        private const decimal AbbreviationThreshold = 100000m;

        public static MoneyText From(decimal value, string prefix = "")
        {
            var sign = value < 0 ? "-" : "";
            var absoluteValue = Math.Abs(value);
            var exact = $"{sign}{prefix}{absoluteValue:N2}";
            if (absoluteValue <= AbbreviationThreshold)
                return new MoneyText(exact, "");

            var abbreviated = $"{sign}{prefix}{Math.Round(absoluteValue / 1000, 0):0}k";
            return new MoneyText(abbreviated, exact);
        }
    }

    public class MonthlyFlowSeriesViewModel
    {
        public string Currency { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string TotalIncomeText { get; set; } = "";
        public string TotalExpenseText { get; set; } = "";
        public string NetChangeText { get; set; } = "";
        public List<MonthlyFlowPointViewModel> Points { get; set; } = [];

        public static MonthlyFlowSeriesViewModel From(MonthlyFlowSeries series)
        {
            var symbol = CurrencySummaryViewModel.FormatCurrencySymbol(series.Currency);
            return new MonthlyFlowSeriesViewModel
            {
                Currency = series.Currency.ToString(),
                DisplayName = series.DisplayName,
                TotalIncomeText = $"收 {symbol}{series.TotalIncome:N2}",
                TotalExpenseText = $"支 {symbol}{series.TotalExpense:N2}",
                NetChangeText = $"净变动 {MoneyText.From(series.NetChange, symbol).DisplayText}",
                Points = series.Points.Select(MonthlyFlowPointViewModel.From).ToList()
            };
        }
    }

    public class MonthlyFlowPointViewModel
    {
        public string MonthLabel { get; set; } = "";
        public decimal Income { get; set; }
        public decimal Expense { get; set; }
        public List<MonthlyFlowSegmentViewModel> IncomeSegments { get; set; } = [];
        public List<MonthlyFlowSegmentViewModel> ExpenseSegments { get; set; } = [];

        public static MonthlyFlowPointViewModel From(MonthlyFlowPoint point)
        {
            return new MonthlyFlowPointViewModel
            {
                MonthLabel = point.MonthLabel,
                Income = point.Income,
                Expense = point.Expense,
                IncomeSegments = point.IncomeSegments.Select(MonthlyFlowSegmentViewModel.From).ToList(),
                ExpenseSegments = point.ExpenseSegments.Select(MonthlyFlowSegmentViewModel.From).ToList()
            };
        }
    }

    public class MonthlyFlowSegmentViewModel
    {
        public string Currency { get; set; } = "";
        public decimal Value { get; set; }

        public static MonthlyFlowSegmentViewModel From(MonthlyFlowSegment segment)
        {
            return new MonthlyFlowSegmentViewModel
            {
                Currency = segment.Currency.ToString(),
                Value = segment.Value
            };
        }
    }

    public class ReasonFlowSeriesViewModel
    {
        public string Currency { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string MonthLabel { get; set; } = "";
        public string TotalIncomeText { get; set; } = "";
        public string TotalExpenseText { get; set; } = "";
        public string TotalFlowText { get; set; } = "";
        public List<ReasonFlowItemViewModel> Items { get; set; } = [];

        public static ReasonFlowSeriesViewModel From(ReasonFlowSeries series)
        {
            var items = series.Items.Select(ReasonFlowItemViewModel.From).ToList();
            var maxTotal = Math.Max(1, items.Count == 0 ? 1 : items.Max(item => item.Total));
            foreach (var item in items)
            {
                item.BarPercent = 100 * (double)(item.Total / maxTotal);
            }

            return new ReasonFlowSeriesViewModel
            {
                Currency = series.Currency.ToString(),
                DisplayName = series.DisplayName,
                MonthLabel = series.MonthLabel,
                TotalIncomeText = $"+¥{series.TotalIncome:N2}",
                TotalExpenseText = $"-¥{series.TotalExpense:N2}",
                TotalFlowText = $"¥{series.TotalIncome + series.TotalExpense:N2}",
                Items = items
            };
        }
    }

    public class ReasonFlowItemViewModel
    {
        public string Reason { get; set; } = "";
        public bool IsIncome { get; set; }
        public decimal Total { get; set; }
        public string KindText => IsIncome ? "收" : "支";
        public string TotalText { get; set; } = "";
        public double BarPercent { get; set; }
        public string CurrencyDetails { get; set; } = "";

        public static ReasonFlowItemViewModel From(ReasonFlowItem item)
        {
            return new ReasonFlowItemViewModel
            {
                Reason = item.Reason,
                IsIncome = item.IsIncome,
                Total = item.Total,
                TotalText = $"¥{item.Total:N0}",
                CurrencyDetails = item.CurrencyDetails
            };
        }
    }

    public class InvestmentStatisticsPeriodViewModel
    {
        public string Title { get; set; } = "";
        public List<InvestmentStatisticsItemViewModel> Items { get; set; } = [];
        public string TotalText { get; set; } = "";

        public static InvestmentStatisticsPeriodViewModel From(InvestmentStatisticsPeriod period)
        {
            return new InvestmentStatisticsPeriodViewModel
            {
                Title = period.Title,
                Items = period.Items.Select(InvestmentStatisticsItemViewModel.From).ToList(),
                TotalText = $"¥{period.Total:N2}"
            };
        }
    }

    public class InvestmentAccountStatisticsViewModel
    {
        public string DisplayName { get; set; } = "";
        public List<InvestmentStatisticsPeriodViewModel> ByReasonPeriods { get; set; } = [];
        public List<InvestmentStatisticsPeriodViewModel> ByHoldingPeriods { get; set; } = [];

        public static InvestmentAccountStatisticsViewModel From(InvestmentAccountStatistics account)
        {
            return new InvestmentAccountStatisticsViewModel
            {
                DisplayName = account.DisplayName,
                ByReasonPeriods = account.ByReason.Periods
                    .Select(InvestmentStatisticsPeriodViewModel.From)
                    .ToList(),
                ByHoldingPeriods = account.ByHolding.Periods
                    .Select(InvestmentStatisticsPeriodViewModel.From)
                    .ToList()
            };
        }
    }

    public class InvestmentStatisticsItemViewModel
    {
        public string Name { get; set; } = "";
        public string TotalText { get; set; } = "";

        public static InvestmentStatisticsItemViewModel From(InvestmentStatisticsItem item)
        {
            return new InvestmentStatisticsItemViewModel
            {
                Name = item.Name,
                TotalText = $"¥{item.Total:N2}"
            };
        }
    }

    public class MonthlyFlowChart : FrameworkElement
    {
        public static readonly DependencyProperty ItemsProperty = DependencyProperty.Register(
            nameof(Items),
            typeof(IEnumerable),
            typeof(MonthlyFlowChart),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));
        public static readonly DependencyProperty FlowKindProperty = DependencyProperty.Register(
            nameof(FlowKind),
            typeof(string),
            typeof(MonthlyFlowChart),
            new FrameworkPropertyMetadata("Income", FrameworkPropertyMetadataOptions.AffectsRender));

        public IEnumerable? Items
        {
            get => (IEnumerable?)GetValue(ItemsProperty);
            set => SetValue(ItemsProperty, value);
        }

        public string FlowKind
        {
            get => (string)GetValue(FlowKindProperty);
            set => SetValue(FlowKindProperty, value);
        }

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);
            var points = Items?.Cast<MonthlyFlowPointViewModel>().ToList() ?? [];
            if (points.Count == 0)
            {
                DrawEmpty(dc);
                return;
            }

            var width = ActualWidth;
            var height = ActualHeight;
            var left = 58.0;
            var top = 46.0;
            var right = 58.0;
            var bottom = 28.0;
            var chartWidth = Math.Max(1, width - left - right);
            var chartHeight = Math.Max(1, height - top - bottom);
            var chartRight = left + chartWidth;
            var useExpense = String.Equals(FlowKind, "Expense", StringComparison.OrdinalIgnoreCase);
            var maxValue = Math.Max(1, points.Max(point => GetFlowSegments(point, useExpense).Sum(segment => segment.Value)));
            var axisStep = CalculateAxisStep(maxValue);
            var axisMax = axisStep * Math.Max(1, (int)Math.Ceiling(maxValue / axisStep));
            var gridPen = new Pen(new SolidColorBrush(Color.FromRgb(226, 232, 240)), 1);

            for (decimal value = 0; value <= axisMax; value += axisStep)
            {
                var y = top + chartHeight - chartHeight * (double)(value / axisMax);
                dc.DrawLine(gridPen, new Point(left, y), new Point(chartRight, y));
                dc.DrawLine(gridPen, new Point(left - 4, y), new Point(left, y));
                dc.DrawLine(gridPen, new Point(chartRight, y), new Point(chartRight + 4, y));
                var label = value.ToString("N0", CultureInfo.InvariantCulture);
                DrawRightAlignedText(dc, label, 11, "#94A3B8", left - 8, y - 8);
                DrawText(dc, label, 11, "#94A3B8", chartRight + 8, y - 8);
            }

            dc.DrawLine(gridPen, new Point(left, top), new Point(left, top + chartHeight));
            dc.DrawLine(gridPen, new Point(chartRight, top), new Point(chartRight, top + chartHeight));
            dc.DrawLine(gridPen, new Point(left, top + chartHeight), new Point(chartRight, top + chartHeight));
            DrawBars(dc, points, axisMax, left, top, chartWidth, chartHeight, useExpense);
            DrawLegend(dc, width);

            for (var i = 0; i < points.Count; i++)
            {
                var groupWidth = chartWidth / points.Count;
                var x = left + groupWidth * i + groupWidth / 2;
                DrawText(dc, points[i].MonthLabel, 11, "#64748B", x - 15, top + chartHeight + 13);
            }
        }

        private static void DrawBars(
            DrawingContext dc,
            List<MonthlyFlowPointViewModel> points,
            decimal axisMax,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            bool useExpense)
        {
            var groupWidth = chartWidth / points.Count;
            var barWidth = Math.Min(24, Math.Max(8, groupWidth * 0.28));
            for (var i = 0; i < points.Count; i++)
            {
                var segments = GetFlowSegments(points[i], useExpense);
                var total = segments.Sum(segment => segment.Value);
                var center = left + groupWidth * i + groupWidth / 2;
                DrawStackedBar(dc, segments, center - barWidth / 2, top, chartHeight, barWidth, axisMax);
                DrawBarValueLabel(dc, center, top, chartHeight, total, axisMax, useExpense);
            }
        }

        private static List<MonthlyFlowSegmentViewModel> GetFlowSegments(MonthlyFlowPointViewModel point, bool useExpense)
        {
            return useExpense ? point.ExpenseSegments : point.IncomeSegments;
        }

        private static void DrawStackedBar(
            DrawingContext dc,
            List<MonthlyFlowSegmentViewModel> segments,
            double x,
            double top,
            double chartHeight,
            double width,
            decimal axisMax)
        {
            var y = top + chartHeight;
            foreach (var segment in segments.Where(segment => segment.Value > 0))
            {
                var height = Math.Max(1, chartHeight * (double)(segment.Value / axisMax));
                y -= height;
                dc.DrawRoundedRectangle(
                    GetCurrencyBrush(segment.Currency),
                    null,
                    new Rect(x, y, width, height),
                    2,
                    2);
            }
        }

        private static void DrawBarValueLabel(
            DrawingContext dc,
            double centerX,
            double top,
            double chartHeight,
            decimal total,
            decimal axisMax,
            bool useExpense)
        {
            if (total <= 0)
                return;

            var color = useExpense ? "#BE123C" : "#047857";
            var text = CreateText(total.ToString("N0", CultureInfo.InvariantCulture), 10, ParseBrush(color));
            var barTop = top + chartHeight - chartHeight * (double)(total / axisMax);
            var textTop = Math.Max(2, barTop - text.Height - 4);
            var textLeft = centerX - text.Width / 2;
            var background = new SolidColorBrush(Color.FromArgb(230, 255, 255, 255));
            dc.DrawRoundedRectangle(
                background,
                null,
                new Rect(textLeft - 3, textTop - 1, text.Width + 6, text.Height + 2),
                4,
                4);
            dc.DrawText(text, new Point(textLeft, textTop));
        }

        private static void DrawLegend(DrawingContext dc, double width)
        {
            var currencies = new[] { "RMB", "USD", "HKD", "JPY", "SGD" };
            var x = Math.Max(80, width - 330);
            foreach (var currency in currencies)
            {
                DrawLegendItem(dc, x, 4, GetCurrencyBrush(currency), currency);
                x += 60;
            }
        }

        private static void DrawLegendItem(DrawingContext dc, double x, double y, Brush brush, string text)
        {
            dc.DrawEllipse(brush, null, new Point(x, y + 8), 4, 4);
            DrawText(dc, text, 12, "#334155", x + 10, y);
        }

        private void DrawEmpty(DrawingContext dc)
        {
            DrawText(dc, "暂无数据", 14, "#94A3B8", ActualWidth / 2 - 30, ActualHeight / 2 - 10);
        }

        private static void DrawText(DrawingContext dc, string text, double size, string color, double x, double y)
        {
            dc.DrawText(CreateText(text, size, ParseBrush(color)), new Point(x, y));
        }

        private static void DrawCenteredText(DrawingContext dc, string text, double size, string color, double centerX, double y)
        {
            var formatted = CreateText(text, size, ParseBrush(color));
            dc.DrawText(formatted, new Point(centerX - formatted.Width / 2, y));
        }

        private static void DrawRightAlignedText(DrawingContext dc, string text, double size, string color, double rightX, double y)
        {
            var formatted = CreateText(text, size, ParseBrush(color));
            dc.DrawText(formatted, new Point(rightX - formatted.Width, y));
        }

        private static FormattedText CreateText(string text, double size, Brush brush)
        {
            return new FormattedText(
                text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Microsoft YaHei UI"),
                size,
                brush,
                1.25);
        }

        private static SolidColorBrush ParseBrush(string color)
        {
            return (SolidColorBrush)new BrushConverter().ConvertFromString(color)!;
        }

        private static decimal CalculateAxisStep(decimal maxValue)
        {
            var rawStep = Math.Max(500, maxValue / 4);
            return Math.Ceiling(rawStep / 500) * 500;
        }

        private static Brush GetCurrencyBrush(string currency)
        {
            return currency switch
            {
                "RMB" => ParseBrush("#0F766E"),
                "USD" => ParseBrush("#2563EB"),
                "HKD" => ParseBrush("#7C3AED"),
                "JPY" => ParseBrush("#F59E0B"),
                "SGD" => ParseBrush("#DB2777"),
                _ => ParseBrush("#64748B")
            };
        }
    }

    public class ReasonFlowChart : FrameworkElement
    {
        public static readonly DependencyProperty ItemsProperty = DependencyProperty.Register(
            nameof(Items),
            typeof(IEnumerable),
            typeof(ReasonFlowChart),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));

        public IEnumerable? Items
        {
            get => (IEnumerable?)GetValue(ItemsProperty);
            set => SetValue(ItemsProperty, value);
        }

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);
            var items = Items?.Cast<ReasonFlowItemViewModel>().ToList() ?? [];
            if (items.Count == 0)
            {
                DrawText(dc, "暂无数据", 14, "#94A3B8", ActualWidth / 2 - 30, ActualHeight / 2 - 10);
                return;
            }

            var visibleItems = items.Take(Math.Max(1, (int)((ActualHeight - 18) / 44))).ToList();
            var labelWidth = 178.0;
            var valueWidth = 210.0;
            var barLeft = labelWidth;
            var barWidth = Math.Max(1, ActualWidth - labelWidth - valueWidth - 22);
            var maxValue = Math.Max(1, visibleItems.Max(item => item.Total));
            var y = 10.0;
            foreach (var item in visibleItems)
            {
                var color = item.IsIncome ? "#047857" : "#BE123C";
                var kind = item.IsIncome ? "收" : "支";
                DrawText(dc, kind, 12, color, 0, y + 3);
                DrawText(dc, TrimText(item.Reason, 16), 12, "#334155", 24, y + 3);
                DrawBar(dc, barLeft, y + 8, barWidth, item.Total, maxValue, color);
                DrawText(dc, $"{item.Total:N0}", 12, color, barLeft + barWidth + 12, y + 2);
                if (!String.IsNullOrWhiteSpace(item.CurrencyDetails))
                    DrawText(dc, TrimText(item.CurrencyDetails, 30), 10, "#94A3B8", barLeft + barWidth + 12, y + 20);
                y += 44;
            }
        }

        private static void DrawBar(DrawingContext dc, double x, double y, double maxWidth, decimal value, decimal maxValue, string color)
        {
            var width = maxWidth * (double)(value / maxValue);
            var background = new SolidColorBrush(Color.FromRgb(241, 245, 249));
            dc.DrawRoundedRectangle(background, null, new Rect(x, y, maxWidth, 7), 3, 3);
            if (width > 0)
                dc.DrawRoundedRectangle(ParseBrush(color), null, new Rect(x, y, width, 7), 3, 3);
        }

        private static string TrimText(string text, int maxLength)
        {
            return text.Length <= maxLength ? text : text[..(maxLength - 1)] + "…";
        }

        private static void DrawText(DrawingContext dc, string text, double size, string color, double x, double y)
        {
            dc.DrawText(
                new FormattedText(
                    text,
                    CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Microsoft YaHei UI"),
                    size,
                    ParseBrush(color),
                    1.25),
                new Point(x, y));
        }

        private static SolidColorBrush ParseBrush(string color)
        {
            return (SolidColorBrush)new BrushConverter().ConvertFromString(color)!;
        }
    }
}
