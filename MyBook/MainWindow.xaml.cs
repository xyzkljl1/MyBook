using Microsoft.Extensions.Configuration;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
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
        private const string DefaultDetailAccountName = "ICBC";

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        readonly Fetcher fetcher = new();
        Forms.NotifyIcon? trayIcon;
        Drawing.Icon? trayIconImage;
        bool isExitRequested;
        bool isLoadingRecordDetails;

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
            var viewModel = DataContext as DashboardViewModel;
            await LoadDashboardAsync(viewModel?.DetailStartDate, viewModel?.DetailEndDate, viewModel?.SelectedDetailAccountName);
        }

        private void AddRecordDetail_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not DashboardViewModel viewModel)
                return;

            var accountName = viewModel.SelectedDetailAccountName
                ?? viewModel.DetailAccountNames.FirstOrDefault()
                ?? "";
            viewModel.RecordDetails.Insert(0, RecordDetailRowViewModel.CreateNew(accountName));
            viewModel.SetDetailStatus("新增记录尚未保存");
        }

        private async void DetailDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isLoadingRecordDetails)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            await LoadRecordDetailsAsync(viewModel);
        }

        private async void DetailAccount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isLoadingRecordDetails)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            await LoadRecordDetailsAsync(viewModel);
        }

        private async void DiscardRecordDetails_Click(object sender, RoutedEventArgs e)
        {
            await LoadRecordDetailsAsync();
            if (DataContext is DashboardViewModel viewModel)
                viewModel.SetDetailStatus("已放弃未保存修改");
        }

        private async void SaveRecordDetails_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not DashboardViewModel viewModel)
                return;

            RecordDetailsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            RecordDetailsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var changes = viewModel.GetPendingRecordChanges()
                .Concat(viewModel.GetPendingBalanceCorrections())
                .ToList();
            if (changes.Count == 0)
            {
                viewModel.SetDetailStatus("没有需要保存的修改");
                return;
            }

            if (!ConfirmRecordDetailChanges(changes))
            {
                viewModel.SetDetailStatus("保存已取消");
                return;
            }

            try
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                database.SaveRecordDetails(changes.Select(change => change.Edit));
                await LoadDashboardAsync(viewModel.DetailStartDate, viewModel.DetailEndDate, viewModel.SelectedDetailAccountName);
            }
            catch (Exception ex)
            {
                viewModel.SetDetailStatus($"保存失败：{ex.Message}");
            }
        }

        private bool ConfirmRecordDetailChanges(IReadOnlyList<RecordDetailPendingChange> changes)
        {
            var text = $"将保存以下 {changes.Count} 项修改，并作为同一个手动导入批次写入数据库：{Environment.NewLine}{Environment.NewLine}"
                + String.Join(
                    Environment.NewLine + Environment.NewLine,
                    changes.Select((change, index) => $"{index + 1}. {change.Description}"));
            var confirmed = false;
            var content = new Grid { Margin = new Thickness(16) };
            content.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            content.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var changeText = new TextBox
            {
                Text = text,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                BorderBrush = Brushes.Transparent,
                Background = Brushes.Transparent,
                FontSize = 13
            };
            content.Children.Add(changeText);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Grid.SetRow(buttons, 1);
            var cancelButton = new Button
            {
                Content = "取消",
                Width = 82,
                Height = 30,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var saveButton = new Button
            {
                Content = "保存",
                Width = 82,
                Height = 30,
                Background = new SolidColorBrush(Color.FromRgb(15, 118, 110)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };
            buttons.Children.Add(cancelButton);
            buttons.Children.Add(saveButton);
            content.Children.Add(buttons);

            var dialog = new Window
            {
                Title = "确认保存明细修改",
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Width = 760,
                Height = 520,
                MinWidth = 560,
                MinHeight = 360,
                Content = content
            };
            cancelButton.Click += (_, _) => dialog.Close();
            saveButton.Click += (_, _) =>
            {
                confirmed = true;
                dialog.Close();
            };
            dialog.ShowDialog();
            return confirmed;
        }

        private async Task LoadDashboardAsync(DateTime? detailStartDate = null, DateTime? detailEndDate = null, string? detailAccountName = null)
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

                var viewModel = DashboardViewModel.From(data);
                if (detailStartDate.HasValue)
                    viewModel.DetailStartDate = detailStartDate.Value;
                if (detailEndDate.HasValue)
                    viewModel.DetailEndDate = detailEndDate.Value;
                viewModel.SelectedDetailAccountName = detailAccountName;
                DataContext = viewModel;
                await LoadRecordDetailsAsync(viewModel);
            }
            catch (Exception e)
            {
                DataContext = DashboardViewModel.FromError(e.Message);
            }
        }

        private async Task LoadRecordDetailsAsync(DashboardViewModel? viewModel = null)
        {
            viewModel ??= DataContext as DashboardViewModel;
            if (viewModel is null)
                return;

            if (isLoadingRecordDetails)
                return;

            isLoadingRecordDetails = true;
            try
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var accountNames = database.GetAccountNames();
                viewModel.SetDetailAccountNames(accountNames, DefaultDetailAccountName);
                var details = database.GetRecordDetails(viewModel.DetailStartDate, viewModel.DetailEndDate, viewModel.SelectedDetailAccountName)
                    .Select(RecordDetailRowViewModel.From)
                    .ToList();
                LoadDetailAccountBalances(viewModel, database);
                viewModel.SetRecordDetails(details);
                viewModel.SetDetailStatus($"共 {details.Count} 条");
            }
            catch (Exception e)
            {
                viewModel.SetDetailStatus($"读取失败：{e.Message}");
            }
            finally
            {
                isLoadingRecordDetails = false;
            }
        }

        private static void LoadDetailAccountBalances(DashboardViewModel viewModel, DatabaseUtil database)
        {
            if (String.IsNullOrWhiteSpace(viewModel.SelectedDetailAccountName))
            {
                viewModel.SetDetailAccountBalances([]);
                return;
            }

            var balances = database.GetAccountBalanceDetails(viewModel.SelectedDetailAccountName)
                .Select(AccountBalanceRowViewModel.From)
                .ToList();
            viewModel.SetDetailAccountBalances(balances);
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
        double selectedAssetSummaryOffset = -1;
        double selectedReasonMonthIndex;
        DateTime detailStartDate = DateTime.Today.AddDays(-30);
        DateTime detailEndDate = DateTime.Today;
        bool showInvestmentByHolding;
        string? selectedDetailAccountName;
        MonthlyFlowAccountStatisticsViewModel? selectedMonthlyAccount;
        InvestmentAccountStatisticsViewModel? selectedInvestmentAccount;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string SnapshotTimeText { get; set; } = "";
        public string SelectedAssetDateText { get; set; } = "";
        public string ReasonTabHeader { get; set; } = "分类";
        public TotalAssetsViewModel TotalAssets { get; set; } = new();
        public List<CurrencySummaryViewModel> CurrencySummaries { get; set; } = [];
        public List<AssetSummaryViewModel> AssetSummaries { get; set; } = [];
        public List<MonthlyFlowSeriesViewModel> MonthlySeries { get; set; } = [];
        public List<MonthlyFlowSeriesViewModel> RmbMonthlySeries { get; set; } = [];
        public List<MonthlyFlowAccountStatisticsViewModel> MonthlyAccounts { get; set; } = [];
        public List<AccountNetFlowStatisticsViewModel> AccountNetFlows { get; set; } = [];
        public List<ReasonFlowSeriesViewModel> ReasonMonthSeries { get; set; } = [];
        public List<InvestmentAccountStatisticsViewModel> InvestmentAccounts { get; set; } = [];
        public ObservableCollection<RecordDetailRowViewModel> RecordDetails { get; } = [];
        public ObservableCollection<AccountBalanceRowViewModel> DetailAccountBalances { get; } = [];
        public List<string> DetailAccountNames { get; private set; } = [];
        public List<CurrencyType> DetailCurrencyTypes { get; } = Enum.GetValues<CurrencyType>().ToList();
        public string DetailStatusText { get; private set; } = "";
        public IEnumerable<MonthlyFlowSeriesViewModel> VisibleMonthlySeries
        {
            get
            {
                if (SelectedMonthlyAccount is null)
                    return ShowSingleCurrencyMonthly ? MonthlySeries : RmbMonthlySeries;

                return ShowSingleCurrencyMonthly
                    ? SelectedMonthlyAccount.MonthlySeries
                    : SelectedMonthlyAccount.RmbMonthlySeries;
            }
        }
        public IEnumerable<InvestmentStatisticsPeriodViewModel> VisibleInvestmentPeriods => SelectedInvestmentAccount is null
            ? Enumerable.Empty<InvestmentStatisticsPeriodViewModel>()
            : ShowInvestmentByHolding
                ? SelectedInvestmentAccount.ByHoldingPeriods
                : SelectedInvestmentAccount.ByReasonPeriods;
        public DoubleCollection AssetSummaryTicks => new(AssetSummaries.Select(summary => summary.DayOffset));
        public double AssetSummaryMaximum => AssetSummaries.Count == 0 ? 0 : AssetSummaries[^1].DayOffset;
        public double ReasonMonthMaximum => Math.Max(0, ReasonMonthSeries.Count - 1);
        public ReasonFlowSeriesViewModel SelectedReasonSeries => ReasonMonthSeries.Count == 0
            ? new ReasonFlowSeriesViewModel()
            : ReasonMonthSeries[Math.Clamp((int)Math.Round(selectedReasonMonthIndex), 0, ReasonMonthSeries.Count - 1)];
        public string SelectedReasonMonthLabel => SelectedReasonSeries.MonthLabel;

        public DateTime DetailStartDate
        {
            get => detailStartDate;
            set
            {
                value = value.Date;
                if (detailStartDate == value)
                    return;

                detailStartDate = value;
                OnPropertyChanged(nameof(DetailStartDate));
                if (detailEndDate < detailStartDate)
                {
                    detailEndDate = detailStartDate;
                    OnPropertyChanged(nameof(DetailEndDate));
                }
            }
        }

        public DateTime DetailEndDate
        {
            get => detailEndDate;
            set
            {
                value = value.Date < detailStartDate ? detailStartDate : value.Date;
                if (detailEndDate == value)
                    return;

                detailEndDate = value;
                OnPropertyChanged(nameof(DetailEndDate));
            }
        }

        public string? SelectedDetailAccountName
        {
            get => selectedDetailAccountName;
            set
            {
                if (selectedDetailAccountName == value)
                    return;

                selectedDetailAccountName = value;
                OnPropertyChanged(nameof(SelectedDetailAccountName));
            }
        }

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

        public MonthlyFlowAccountStatisticsViewModel? SelectedMonthlyAccount
        {
            get => selectedMonthlyAccount;
            set
            {
                if (ReferenceEquals(selectedMonthlyAccount, value))
                    return;

                selectedMonthlyAccount = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedMonthlyAccount)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(VisibleMonthlySeries)));
            }
        }

        public double SelectedAssetSummaryOffset
        {
            get => selectedAssetSummaryOffset;
            set
            {
                var normalized = AssetSummaries.Count == 0
                    ? 0
                    : AssetSummaries
                        .OrderBy(summary => Math.Abs(summary.DayOffset - value))
                        .First()
                        .DayOffset;
                if (selectedAssetSummaryOffset == normalized)
                    return;

                selectedAssetSummaryOffset = normalized;
                ApplySelectedAssetSummary();
                OnPropertyChanged(nameof(SelectedAssetSummaryOffset));
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
            List<MonthlyFlowAccountStatistics> monthlyAccounts = data.MonthlyAccounts.Count == 0
                ?
                [
                    new MonthlyFlowAccountStatistics
                    {
                        DisplayName = "所有账户",
                        MonthlyFlowSeries = data.MonthlyFlowSeries,
                        RmbMonthlyFlowSeries = data.RmbMonthlyFlowSeries
                    }
                ]
                : data.MonthlyAccounts;
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
                AssetSummaries = BuildAssetSummaryViewModels(data),
                MonthlySeries = data.MonthlyFlowSeries
                    .Select(MonthlyFlowSeriesViewModel.From)
                    .ToList(),
                RmbMonthlySeries = [MonthlyFlowSeriesViewModel.From(data.RmbMonthlyFlowSeries)],
                MonthlyAccounts = monthlyAccounts
                    .Select(MonthlyFlowAccountStatisticsViewModel.From)
                    .ToList(),
                AccountNetFlows = data.AccountNetFlows
                    .Select(AccountNetFlowStatisticsViewModel.From)
                    .ToList(),
                ReasonMonthSeries = reasonSeries,
                InvestmentAccounts = investmentAccounts
                    .Select(InvestmentAccountStatisticsViewModel.From)
                    .ToList()
            };
            viewModel.SelectedMonthlyAccount = viewModel.MonthlyAccounts.FirstOrDefault();
            viewModel.SelectedInvestmentAccount = viewModel.InvestmentAccounts.FirstOrDefault();
            viewModel.SelectedAssetSummaryOffset = viewModel.AssetSummaries.Count == 0 ? 0 : viewModel.AssetSummaries[^1].DayOffset;
            viewModel.SelectedReasonMonthIndex = Math.Clamp(data.DefaultReasonMonthIndex, 0, Math.Max(0, reasonSeries.Count - 1));
            return viewModel;
        }

        private static List<AssetSummaryViewModel> BuildAssetSummaryViewModels(DashboardData data)
        {
            if (data.AssetSummaryPoints.Count > 0)
            {
                var firstDate = data.AssetSummaryPoints.Min(point => point.Date);
                return data.AssetSummaryPoints
                    .Select(point => AssetSummaryViewModel.From(point, firstDate))
                    .ToList();
            }

            return
            [
                AssetSummaryViewModel.FromCurrent(data.TotalAssetsRmb, data.CurrencySummaries)
            ];
        }

        public static DashboardViewModel FromError(string message)
        {
            return new DashboardViewModel
            {
                SnapshotTimeText = $"数据读取失败：{message}"
            };
        }

        private void ApplySelectedAssetSummary()
        {
            if (AssetSummaries.Count == 0)
                return;

            var selected = AssetSummaries
                .OrderBy(summary => Math.Abs(summary.DayOffset - selectedAssetSummaryOffset))
                .First();
            SnapshotTimeText = selected.StatusText;
            SelectedAssetDateText = selected.DateLabel;
            TotalAssets = selected.TotalAssets;
            CurrencySummaries = selected.CurrencySummaries;
            OnPropertyChanged(nameof(SnapshotTimeText));
            OnPropertyChanged(nameof(SelectedAssetDateText));
            OnPropertyChanged(nameof(TotalAssets));
            OnPropertyChanged(nameof(CurrencySummaries));
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void SetDetailAccountNames(IEnumerable<string> accountNames, string defaultAccountName)
        {
            DetailAccountNames = accountNames.ToList();
            if (String.IsNullOrWhiteSpace(SelectedDetailAccountName) ||
                !DetailAccountNames.Contains(SelectedDetailAccountName, StringComparer.OrdinalIgnoreCase))
            {
                SelectedDetailAccountName = DetailAccountNames.Contains(defaultAccountName, StringComparer.OrdinalIgnoreCase)
                    ? defaultAccountName
                    : DetailAccountNames.FirstOrDefault();
            }

            OnPropertyChanged(nameof(DetailAccountNames));
        }

        public void SetRecordDetails(IEnumerable<RecordDetailRowViewModel> records)
        {
            RecordDetails.Clear();
            foreach (var record in records)
                RecordDetails.Add(record);
            OnPropertyChanged(nameof(RecordDetails));
        }

        public void SetDetailStatus(string status)
        {
            DetailStatusText = status;
            OnPropertyChanged(nameof(DetailStatusText));
        }

        public void SetDetailAccountBalances(IEnumerable<AccountBalanceRowViewModel> balances)
        {
            DetailAccountBalances.Clear();
            foreach (var balance in balances)
                DetailAccountBalances.Add(balance);
            OnPropertyChanged(nameof(DetailAccountBalances));
        }

        public List<RecordDetailPendingChange> GetPendingRecordChanges()
        {
            return RecordDetails
                .Select(record => record.GetPendingChange())
                .Where(change => change is not null)
                .Select(change => change!)
                .ToList();
        }

        public List<RecordDetailPendingChange> GetPendingBalanceCorrections()
        {
            return DetailAccountBalances
                .Select(balance => balance.GetPendingCorrection(SelectedDetailAccountName))
                .Where(change => change is not null)
                .Select(change => change!)
                .ToList();
        }
    }

    public class AssetSummaryViewModel
    {
        public double DayOffset { get; set; }
        public string DateLabel { get; set; } = "";
        public string StatusText { get; set; } = "";
        public TotalAssetsViewModel TotalAssets { get; set; } = new();
        public List<CurrencySummaryViewModel> CurrencySummaries { get; set; } = [];

        public static AssetSummaryViewModel From(AssetSummaryPoint point, DateTime firstDate)
        {
            var dateLabel = point.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return new AssetSummaryViewModel
            {
                DayOffset = (point.Date - firstDate).TotalDays,
                DateLabel = dateLabel,
                StatusText = BuildStatusText(point, dateLabel),
                TotalAssets = TotalAssetsViewModel.From(point.HasData ? point.TotalAssetsRmb : 0),
                CurrencySummaries = point.HasData
                    ? point.CurrencySummaries.Select(CurrencySummaryViewModel.From).ToList()
                    : []
            };
        }

        public static AssetSummaryViewModel FromCurrent(
            decimal totalAssetsRmb,
            List<CurrencyBalanceSummary> currencySummaries)
        {
            return new AssetSummaryViewModel
            {
                DateLabel = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                StatusText = $"更新于 {DateTime.Now:yyyy-MM-dd HH:mm}",
                TotalAssets = TotalAssetsViewModel.From(totalAssetsRmb),
                CurrencySummaries = currencySummaries.Select(CurrencySummaryViewModel.From).ToList()
            };
        }

        private static string BuildStatusText(AssetSummaryPoint point, string dateLabel)
        {
            if (!point.HasData)
                return $"{dateLabel} 缺少快照";

            if (point.IsToday)
                return $"更新于 {DateTime.Now:yyyy-MM-dd HH:mm}";

            var snapshotTime = point.SnapshotTime ?? point.Date;
            return $"快照 {snapshotTime:yyyy-MM-dd HH:mm}";
        }
    }

    public class RecordDetailRowViewModel
    {
        RecordDetailSnapshot? original;

        public int Id { get; set; }
        public DateTime Date { get; set; }
        public string AccountName { get; set; } = "";
        public decimal Amount { get; set; }
        public CurrencyType Currency { get; set; }
        public string Reason { get; set; } = "";
        public string DestAccount { get; set; } = "";
        public int HoldingQuantity { get; set; }
        public bool IsInternal { get; set; }
        public int? MatchedRecordId { get; set; }
        public string MatchedRecordReason { get; set; } = "";
        public bool IsRefundMatched { get; set; }
        public string Source { get; set; } = "";
        public StatementImportProvider StatementProvider { get; set; }
        public string StatementProviderText { get; set; } = "";
        public string StatementKey { get; set; } = "";
        public bool HasBackup { get; set; }
        public bool IsDateReadOnly => StatementProvider != StatementImportProvider.Manual;
        public bool IsCurrencyEditable => StatementProvider == StatementImportProvider.Manual;
        public bool IsHoldingQuantityReadOnly => original is not null && original.Value.HoldingQuantity == 0;

        public static RecordDetailRowViewModel From(RecordDetailData data)
        {
            var row = new RecordDetailRowViewModel
            {
                Id = data.Id,
                Date = data.Date,
                AccountName = data.AccountName,
                Amount = data.Amount,
                Currency = data.Currency,
                Reason = data.Reason,
                DestAccount = data.DestAccount,
                HoldingQuantity = data.HoldingQuantity,
                IsInternal = data.IsInternal,
                MatchedRecordId = data.MatchedRecordId,
                MatchedRecordReason = data.MatchedRecordReason,
                IsRefundMatched = data.IsRefundMatched,
                Source = data.Source,
                StatementProvider = data.StatementProvider,
                StatementProviderText = data.StatementProvider.ToString(),
                StatementKey = data.StatementKey,
                HasBackup = data.HasBackup
            };
            row.original = row.ToSnapshot();
            return row;
        }

        public static RecordDetailRowViewModel CreateNew(string accountName)
        {
            return new RecordDetailRowViewModel
            {
                Date = DateTime.Today,
                AccountName = accountName,
                Currency = CurrencyType.RMB,
                StatementProvider = StatementImportProvider.Manual,
                StatementProviderText = StatementImportProvider.Manual.ToString()
            };
        }

        public RecordDetailPendingChange? GetPendingChange()
        {
            if (original is null)
                return new RecordDetailPendingChange(ToEdit(), $"新增 {FormatSummary()}");

            var current = ToSnapshot();
            if (current == original)
                return null;

            return new RecordDetailPendingChange(ToEdit(), $"修改 #{Id} {FormatSummary()}{Environment.NewLine}  {BuildFieldChanges(original.Value, current)}");
        }

        public RecordDetailEdit ToEdit()
        {
            return new RecordDetailEdit
            {
                Id = Id,
                Date = Date,
                AccountName = AccountName ?? "",
                Amount = Amount,
                Currency = Currency,
                Reason = Reason ?? "",
                DestAccount = DestAccount ?? "",
                HoldingQuantity = HoldingQuantity,
                IsInternal = IsInternal,
                IsRefundMatched = IsRefundMatched,
                Source = Source ?? ""
            };
        }

        private RecordDetailSnapshot ToSnapshot()
        {
            return new RecordDetailSnapshot(
                Date,
                AccountName ?? "",
                Amount,
                Currency,
                Reason ?? "",
                DestAccount ?? "",
                HoldingQuantity,
                IsInternal,
                IsRefundMatched,
                Source ?? "");
        }

        private string FormatSummary()
        {
            var target = String.IsNullOrWhiteSpace(DestAccount) ? "" : $" / {DestAccount}";
            return $"{Date:yyyy-MM-dd HH:mm} {AccountName} {Amount:0.##} {Currency} {Reason}{target}";
        }

        private static string BuildFieldChanges(RecordDetailSnapshot original, RecordDetailSnapshot current)
        {
            var changes = new List<string>();
            AddChange(changes, "日期", original.Date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture), current.Date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture));
            AddChange(changes, "账户", original.AccountName, current.AccountName);
            AddChange(changes, "金额", original.Amount.ToString("0.##", CultureInfo.InvariantCulture), current.Amount.ToString("0.##", CultureInfo.InvariantCulture));
            AddChange(changes, "币种", original.Currency.ToString(), current.Currency.ToString());
            AddChange(changes, "原因", original.Reason, current.Reason);
            AddChange(changes, "对方", original.DestAccount, current.DestAccount);
            AddChange(changes, "数量", original.HoldingQuantity.ToString(CultureInfo.InvariantCulture), current.HoldingQuantity.ToString(CultureInfo.InvariantCulture));
            AddChange(changes, "内部", original.IsInternal ? "是" : "否", current.IsInternal ? "是" : "否");
            AddChange(changes, "已抵消", original.IsRefundMatched ? "是" : "否", current.IsRefundMatched ? "是" : "否");
            AddChange(changes, "来源", original.Source, current.Source);
            return String.Join("；", changes);
        }

        private static void AddChange(List<string> changes, string name, string oldValue, string newValue)
        {
            if (oldValue == newValue)
                return;

            changes.Add($"{name}: {oldValue} -> {newValue}");
        }
    }

    public class AccountBalanceRowViewModel
    {
        decimal originalAmount;

        public CurrencyType Currency { get; set; }
        public decimal Amount { get; set; }

        public static AccountBalanceRowViewModel From(AccountBalanceDetailData data)
        {
            var row = new AccountBalanceRowViewModel
            {
                Currency = data.Currency,
                Amount = data.Amount
            };
            row.originalAmount = row.Amount;
            return row;
        }

        public RecordDetailPendingChange? GetPendingCorrection(string? accountName)
        {
            if (String.IsNullOrWhiteSpace(accountName))
                return null;

            var delta = Amount - originalAmount;
            if (delta == 0)
                return null;

            var edit = new RecordDetailEdit
            {
                AccountName = accountName,
                Amount = delta,
                Currency = Currency,
                Date = DateTime.Now,
                Reason = "手动校正"
            };
            var description = $"余额校正 {accountName} {Currency}: {originalAmount:0.############} -> {Amount:0.############}，生成记录 {FormatSigned(delta)} {Currency}，原因 手动校正";
            return new RecordDetailPendingChange(edit, description);
        }

        private static string FormatSigned(decimal value)
        {
            return value.ToString("+0.############;-0.############;0", CultureInfo.InvariantCulture);
        }
    }

    public record RecordDetailPendingChange(RecordDetailEdit Edit, string Description);

    readonly record struct RecordDetailSnapshot(
        DateTime Date,
        string AccountName,
        decimal Amount,
        CurrencyType Currency,
        string Reason,
        string DestAccount,
        int HoldingQuantity,
        bool IsInternal,
        bool IsRefundMatched,
        string Source);

    public class CurrencySummaryViewModel
    {
        public string Currency { get; set; } = "";
        public string NetText { get; set; } = "";
        public string NetLineText { get; set; } = "";
        public string NetExactText { get; set; } = "";
        public string IncomeText { get; set; } = "";
        public string ExpenseText { get; set; } = "";
        public string AssetsText { get; set; } = "";
        public string LiabilitiesText { get; set; } = "";
        public string BreakdownSeparatorText { get; set; } = "";

        public static CurrencySummaryViewModel From(CurrencyBalanceSummary summary)
        {
            var net = FormatSummaryMoney(summary.Net, summary.Currency);
            return new CurrencySummaryViewModel
            {
                Currency = summary.Currency.ToString(),
                NetText = net.DisplayText,
                NetLineText = net.DisplayText,
                NetExactText = net.ExactText,
                IncomeText = FormatFlowAmount(summary.TotalIncome, "+"),
                ExpenseText = FormatFlowAmount(summary.TotalExpense, "-"),
                AssetsText = FormatBreakdownAmount(summary.Assets, summary.Liabilities),
                LiabilitiesText = FormatBreakdownAmount(summary.Liabilities, summary.Assets),
                BreakdownSeparatorText = summary.Assets != 0 && summary.Liabilities != 0 ? "-" : ""
            };
        }

        private static MoneyText FormatSummaryMoney(decimal value, CurrencyType currency)
        {
            var text = MoneyText.From(value);
            return new MoneyText($"{text.DisplayText}  {currency}", text.ExactText);
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

        private static string FormatBreakdownAmount(decimal value, decimal pairedValue)
        {
            if (value == 0 || pairedValue == 0)
                return "";

            return MoneyText.FormatAmount(Math.Abs(value));
        }

        private static string FormatCompactAmount(decimal value)
        {
            var sign = value < 0 ? "-" : "";
            var absoluteValue = Math.Abs(value);
            if (absoluteValue >= 1000)
                return $"{sign}{Math.Round(absoluteValue / 1000, 0):0}k";

            return $"{sign}{MoneyText.FormatAmount(absoluteValue)}";
        }

        private static string FormatFlowAmount(decimal value, string sign)
        {
            return value == 0 ? "" : $"{sign}{FormatCompactAmount(Math.Abs(value))}";
        }
    }

    public class TotalAssetsViewModel
    {
        public string ValueText { get; set; } = "";
        public string ExactText { get; set; } = "";

        public static TotalAssetsViewModel From(decimal value)
        {
            var text = MoneyText.From(value, "¥ ");
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

        public string ExactTextOrDisplay => String.IsNullOrWhiteSpace(ExactText) ? DisplayText : ExactText;

        public static MoneyText From(decimal value, string prefix = "")
        {
            var sign = value < 0 ? "-" : "";
            var absoluteValue = Math.Abs(value);
            var exact = $"{sign}{prefix}{FormatAmount(absoluteValue)}";
            if (absoluteValue <= AbbreviationThreshold)
                return new MoneyText(exact, "");

            var abbreviated = $"{sign}{prefix}{Math.Round(absoluteValue / 1000, 0):0}k";
            return new MoneyText(abbreviated, exact);
        }

        public static string FormatAmount(decimal value)
        {
            var rounded = Decimal.Round(value, 2, MidpointRounding.ToEven);
            return rounded == Decimal.Truncate(rounded)
                ? rounded.ToString("N0", CultureInfo.InvariantCulture)
                : rounded.ToString("N2", CultureInfo.InvariantCulture);
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

    public class MonthlyFlowAccountStatisticsViewModel
    {
        public string DisplayName { get; set; } = "";
        public List<MonthlyFlowSeriesViewModel> MonthlySeries { get; set; } = [];
        public List<MonthlyFlowSeriesViewModel> RmbMonthlySeries { get; set; } = [];

        public static MonthlyFlowAccountStatisticsViewModel From(MonthlyFlowAccountStatistics account)
        {
            return new MonthlyFlowAccountStatisticsViewModel
            {
                DisplayName = account.DisplayName,
                MonthlySeries = account.MonthlyFlowSeries
                    .Select(MonthlyFlowSeriesViewModel.From)
                    .ToList(),
                RmbMonthlySeries = [MonthlyFlowSeriesViewModel.From(account.RmbMonthlyFlowSeries)]
            };
        }
    }

    public class AccountNetFlowStatisticsViewModel
    {
        public string DisplayName { get; set; } = "";
        public decimal NetRmb { get; set; }
        public string NetRmbText { get; set; } = "";
        public string CurrencyDetails { get; set; } = "";
        public string RecordCountText { get; set; } = "";
        public bool IsInflow => NetRmb > 0;
        public bool IsOutflow => NetRmb < 0;

        public static AccountNetFlowStatisticsViewModel From(AccountNetFlowStatistics statistics)
        {
            var missingRateText = statistics.HasMissingExchangeRate ? " + 未折算" : "";
            return new AccountNetFlowStatisticsViewModel
            {
                DisplayName = statistics.DisplayName,
                NetRmb = statistics.NetRmb,
                NetRmbText = $"{FormatSignedMoney(statistics.NetRmb, "¥")}{missingRateText}",
                CurrencyDetails = String.Join("、", statistics.CurrencyTotals.Select(FormatCurrencyTotal)),
                RecordCountText = $"{statistics.RecordCount} 笔"
            };
        }

        private static string FormatCurrencyTotal(AccountNetFlowCurrencyTotal total)
        {
            var symbol = CurrencySummaryViewModel.FormatCurrencySymbol(total.Currency);
            return $"{total.Currency} {FormatSignedMoneyExact(total.Amount, symbol)}";
        }

        private static string FormatSignedMoney(decimal value, string symbol)
        {
            var sign = value > 0 ? "+" : value < 0 ? "-" : "";
            return $"{sign}{symbol}{MoneyText.FormatAmount(Math.Abs(value))}";
        }

        private static string FormatSignedMoneyExact(decimal value, string symbol)
        {
            var sign = value > 0 ? "+" : value < 0 ? "-" : "";
            var amount = Math.Abs(value).ToString("#,0.##################", CultureInfo.InvariantCulture);
            return $"{sign}{symbol}{amount}";
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
