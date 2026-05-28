using Microsoft.Extensions.Configuration;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        bool isLoadingRecordDetails;
        bool isLoadingAllocatedExpenses;
        bool isAdjustingAllocatedExpenseRange;

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

        private async void AllocatedExpenseDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isLoadingAllocatedExpenses || isAdjustingAllocatedExpenseRange)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            await LoadAllocatedExpensesAsync(viewModel);
        }

        private async void AllocatedExpenseUnit_Changed(object sender, RoutedEventArgs e)
        {
            if (isLoadingAllocatedExpenses || isAdjustingAllocatedExpenseRange)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            isAdjustingAllocatedExpenseRange = true;
            viewModel.ResetAllocatedExpenseRange();
            isAdjustingAllocatedExpenseRange = false;
            await LoadAllocatedExpensesAsync(viewModel);
        }

        private async void DashboardTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(sender, e.OriginalSource) || isLoadingAllocatedExpenses)
                return;
            if (sender is not TabControl tabControl || !ReferenceEquals(tabControl.SelectedItem, AllocatedExpenseTab))
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            await LoadAllocatedExpensesAsync(viewModel);
        }

        private async void AllocatedExpenseChart_WindowShiftRequested(object sender, AllocatedExpenseWindowShiftEventArgs e)
        {
            if (e.UnitOffset == 0 || isLoadingAllocatedExpenses)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            isAdjustingAllocatedExpenseRange = true;
            viewModel.MoveAllocatedExpenseWindow(e.UnitOffset);
            isAdjustingAllocatedExpenseRange = false;
            await LoadAllocatedExpensesAsync(viewModel);
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
                var previousViewModel = DataContext as DashboardViewModel;
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var data = database.GetDashboardData(DateTime.Today);
                if (data.MissingExchangeRateCurrencies.Count > 0)
                {
                    using var pubWeb = new PubWebUtil(config, database);
                    await pubWeb.FetchExchangeRates(data.MissingExchangeRateCurrencies);
                    data = database.GetDashboardData(DateTime.Today);
                    if (data.MissingExchangeRateCurrencies.Count > 0)
                    {
                        throw new InvalidOperationException(
                            $"缺少汇率：{String.Join("、", data.MissingExchangeRateCurrencies)}");
                    }
                }

                var viewModel = DashboardViewModel.From(data);
                if (previousViewModel is not null)
                    viewModel.CopyAllocatedExpenseSettingsFrom(previousViewModel);
                if (detailStartDate.HasValue)
                    viewModel.DetailStartDate = detailStartDate.Value;
                if (detailEndDate.HasValue)
                    viewModel.DetailEndDate = detailEndDate.Value;
                viewModel.SelectedDetailAccountName = detailAccountName;
                DataContext = viewModel;
                await LoadRecordDetailsAsync(viewModel);
                await LoadAllocatedExpensesAsync(viewModel);
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
                var accounts = database.GetAllAccounts();
                var accountNames = accounts
                    .Select(account => account.name)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var defaultDetailAccountName = accounts
                    .Where(account => account.usage != AccountUsage.Undetermined)
                    .OrderBy(account => account.Id)
                    .Select(account => account.name)
                    .FirstOrDefault()
                    ?? accountNames.FirstOrDefault()
                    ?? "";
                viewModel.SetDetailAccountNames(accountNames, defaultDetailAccountName);
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

        private async Task LoadAllocatedExpensesAsync(DashboardViewModel? viewModel = null)
        {
            viewModel ??= DataContext as DashboardViewModel;
            if (viewModel is null || isLoadingAllocatedExpenses)
                return;

            isLoadingAllocatedExpenses = true;
            try
            {
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                database.ProcessAllocatedExpenseDirtyRecords();
                if (viewModel.ShowAllocatedExpenseMonthly)
                {
                    var firstMonth = new DateTime(
                        viewModel.AllocatedExpenseStartDate.Year,
                        viewModel.AllocatedExpenseStartDate.Month,
                        1);
                    var endMonthExclusive = new DateTime(
                        viewModel.AllocatedExpenseEndDate.Year,
                        viewModel.AllocatedExpenseEndDate.Month,
                        1).AddMonths(1);
                    var items = database.GetAllocatedExpenseMonthly(firstMonth, endMonthExclusive);
                    var details = database.GetAllocatedExpenseRecordDetails(firstMonth, endMonthExclusive);
                    viewModel.SetAllocatedExpenseBuckets(AllocatedExpenseBucketViewModel.FromMonthly(items, details, firstMonth, endMonthExclusive));
                    viewModel.SetAllocatedExpenseStatus($"共 {viewModel.AllocatedExpenseBuckets.Count} 个月");
                }
                else
                {
                    var items = database.GetAllocatedExpenseDaily(
                        viewModel.AllocatedExpenseStartDate,
                        viewModel.AllocatedExpenseEndDate.AddDays(1));
                    var details = database.GetAllocatedExpenseRecordDetails(
                        viewModel.AllocatedExpenseStartDate,
                        viewModel.AllocatedExpenseEndDate.AddDays(1));
                    viewModel.SetAllocatedExpenseBuckets(AllocatedExpenseBucketViewModel.FromDaily(
                        items,
                        details,
                        viewModel.AllocatedExpenseStartDate,
                        viewModel.AllocatedExpenseEndDate));
                    viewModel.SetAllocatedExpenseStatus($"共 {viewModel.AllocatedExpenseBuckets.Count} 天");
                }
            }
            catch (Exception e)
            {
                viewModel.SetAllocatedExpenseBuckets([]);
                viewModel.SetAllocatedExpenseStatus($"均摊支出读取失败：{e.Message}");
            }
            finally
            {
                isLoadingAllocatedExpenses = false;
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
        bool showAllocatedExpenseLineChart;
        bool showAllocatedExpenseMonthly;
        double selectedAssetSummaryOffset = -1;
        double selectedReasonMonthIndex;
        DateTime detailStartDate = DateTime.Today.AddDays(-30);
        DateTime detailEndDate = DateTime.Today;
        DateTime allocatedExpenseStartDate = DateTime.Today.AddDays(-14);
        DateTime allocatedExpenseEndDate = DateTime.Today;
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
        public ObservableCollection<AllocatedExpenseBucketViewModel> AllocatedExpenseBuckets { get; } = [];
        public List<string> DetailAccountNames { get; private set; } = [];
        public List<CurrencyType> DetailCurrencyTypes { get; } = Enum.GetValues<CurrencyType>().ToList();
        public string DetailStatusText { get; private set; } = "";
        public string AllocatedExpenseStatusText { get; private set; } = "";
        public string AllocatedExpenseTotalText => $"¥{AllocatedExpenseBuckets.Sum(bucket => bucket.TotalRmb):N2}";
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

        public DateTime AllocatedExpenseStartDate
        {
            get => allocatedExpenseStartDate;
            set
            {
                value = value.Date;
                if (allocatedExpenseStartDate == value)
                    return;

                allocatedExpenseStartDate = value;
                OnPropertyChanged(nameof(AllocatedExpenseStartDate));
                if (allocatedExpenseEndDate < allocatedExpenseStartDate)
                {
                    allocatedExpenseEndDate = allocatedExpenseStartDate;
                    OnPropertyChanged(nameof(AllocatedExpenseEndDate));
                }
            }
        }

        public DateTime AllocatedExpenseEndDate
        {
            get => allocatedExpenseEndDate;
            set
            {
                value = value.Date < allocatedExpenseStartDate ? allocatedExpenseStartDate : value.Date;
                if (allocatedExpenseEndDate == value)
                    return;

                allocatedExpenseEndDate = value;
                OnPropertyChanged(nameof(AllocatedExpenseEndDate));
            }
        }

        public bool ShowAllocatedExpenseLineChart
        {
            get => showAllocatedExpenseLineChart;
            set
            {
                if (showAllocatedExpenseLineChart == value)
                    return;

                showAllocatedExpenseLineChart = value;
                OnPropertyChanged(nameof(ShowAllocatedExpenseLineChart));
            }
        }

        public bool ShowAllocatedExpenseMonthly
        {
            get => showAllocatedExpenseMonthly;
            set
            {
                if (showAllocatedExpenseMonthly == value)
                    return;

                showAllocatedExpenseMonthly = value;
                OnPropertyChanged(nameof(ShowAllocatedExpenseMonthly));
            }
        }

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

        public void CopyAllocatedExpenseSettingsFrom(DashboardViewModel source)
        {
            showAllocatedExpenseLineChart = source.ShowAllocatedExpenseLineChart;
            showAllocatedExpenseMonthly = source.ShowAllocatedExpenseMonthly;
            allocatedExpenseStartDate = source.AllocatedExpenseStartDate;
            allocatedExpenseEndDate = source.AllocatedExpenseEndDate;
            OnPropertyChanged(nameof(ShowAllocatedExpenseLineChart));
            OnPropertyChanged(nameof(ShowAllocatedExpenseMonthly));
            OnPropertyChanged(nameof(AllocatedExpenseStartDate));
            OnPropertyChanged(nameof(AllocatedExpenseEndDate));
        }

        public void ResetAllocatedExpenseRange()
        {
            if (ShowAllocatedExpenseMonthly)
            {
                allocatedExpenseStartDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-14);
                allocatedExpenseEndDate = DateTime.Today;
            }
            else
            {
                allocatedExpenseStartDate = DateTime.Today.AddDays(-14);
                allocatedExpenseEndDate = DateTime.Today;
            }

            OnPropertyChanged(nameof(AllocatedExpenseStartDate));
            OnPropertyChanged(nameof(AllocatedExpenseEndDate));
        }

        public void MoveAllocatedExpenseWindow(int unitOffset)
        {
            if (unitOffset == 0)
                return;

            if (ShowAllocatedExpenseMonthly)
            {
                allocatedExpenseStartDate = allocatedExpenseStartDate.AddMonths(unitOffset);
                allocatedExpenseEndDate = allocatedExpenseEndDate.AddMonths(unitOffset);
            }
            else
            {
                allocatedExpenseStartDate = allocatedExpenseStartDate.AddDays(unitOffset);
                allocatedExpenseEndDate = allocatedExpenseEndDate.AddDays(unitOffset);
            }

            OnPropertyChanged(nameof(AllocatedExpenseStartDate));
            OnPropertyChanged(nameof(AllocatedExpenseEndDate));
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

        public void SetAllocatedExpenseBuckets(IEnumerable<AllocatedExpenseBucketViewModel> buckets)
        {
            AllocatedExpenseBuckets.Clear();
            foreach (var bucket in buckets)
                AllocatedExpenseBuckets.Add(bucket);
            OnPropertyChanged(nameof(AllocatedExpenseBuckets));
            OnPropertyChanged(nameof(AllocatedExpenseTotalText));
        }

        public void SetAllocatedExpenseStatus(string status)
        {
            AllocatedExpenseStatusText = status;
            OnPropertyChanged(nameof(AllocatedExpenseStatusText));
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
        public int? ExpenseAllocationDays { get; set; }
        public int? ExpenseAllocationSkipDays { get; set; }
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
                ExpenseAllocationDays = data.ExpenseAllocationDays,
                ExpenseAllocationSkipDays = data.ExpenseAllocationSkipDays,
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
                ExpenseAllocationDays = ExpenseAllocationDays,
                ExpenseAllocationSkipDays = ExpenseAllocationSkipDays,
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
                ExpenseAllocationDays,
                ExpenseAllocationSkipDays,
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
            AddChange(changes, "均摊天数", FormatNullableInt(original.ExpenseAllocationDays), FormatNullableInt(current.ExpenseAllocationDays));
            AddChange(changes, "均摊跳过", FormatNullableInt(original.ExpenseAllocationSkipDays), FormatNullableInt(current.ExpenseAllocationSkipDays));
            AddChange(changes, "来源", original.Source, current.Source);
            return String.Join("；", changes);
        }

        private static string FormatNullableInt(int? value)
        {
            return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : "";
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
        int? ExpenseAllocationDays,
        int? ExpenseAllocationSkipDays,
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

    public class AllocatedExpenseBucketViewModel
    {
        public DateTime PeriodStart { get; set; }
        public string PeriodLabel { get; set; } = "";
        public decimal TotalRmb { get; set; }
        public string TotalText { get; set; } = "";
        public string DetailsText { get; set; } = "";
        public List<AllocatedExpenseSegmentViewModel> Segments { get; set; } = [];

        public static List<AllocatedExpenseBucketViewModel> FromDaily(
            IEnumerable<AllocatedExpenseDailyData> items,
            IEnumerable<AllocatedExpenseRecordData> details,
            DateTime startInclusive,
            DateTime endInclusive)
        {
            var groups = items
                .GroupBy(item => item.Date.Date)
                .ToDictionary(group => group.Key, group => group.ToList());
            var detailList = details.ToList();
            var result = new List<AllocatedExpenseBucketViewModel>();
            for (var date = startInclusive.Date; date <= endInclusive.Date; date = date.AddDays(1))
            {
                groups.TryGetValue(date, out var dayItems);
                var periodDetails = detailList
                    .Where(detail => detail.AllocationStart <= date && detail.AllocationEnd >= date)
                    .ToList();
                result.Add(FromItems(date, date.ToString("MM-dd", CultureInfo.InvariantCulture), dayItems ?? [], periodDetails));
            }

            return result;
        }

        public static List<AllocatedExpenseBucketViewModel> FromMonthly(
            IEnumerable<AllocatedExpenseMonthlyData> items,
            IEnumerable<AllocatedExpenseRecordData> details,
            DateTime firstMonthInclusive,
            DateTime endMonthExclusive)
        {
            var groups = items
                .GroupBy(item => new DateTime(item.Month.Year, item.Month.Month, 1))
                .ToDictionary(group => group.Key, group => group.ToList());
            var detailList = details.ToList();
            var result = new List<AllocatedExpenseBucketViewModel>();
            for (var month = firstMonthInclusive.Date; month < endMonthExclusive.Date; month = month.AddMonths(1))
            {
                groups.TryGetValue(month, out var monthItems);
                var nextMonth = month.AddMonths(1);
                var periodDetails = detailList
                    .Where(detail => detail.AllocationStart < nextMonth && detail.AllocationEnd >= month)
                    .ToList();
                result.Add(FromItems(month, month.ToString("yy-MM", CultureInfo.InvariantCulture), monthItems ?? [], periodDetails));
            }

            return result;
        }

        private static AllocatedExpenseBucketViewModel FromItems<T>(
            DateTime periodStart,
            string label,
            List<T> items,
            List<AllocatedExpenseRecordData> detailRecords)
            where T : class
        {
            var segments = items
                .Select(item => item switch
                {
                    AllocatedExpenseDailyData daily => AllocatedExpenseSegmentViewModel.From(daily.Reason, daily.Currency, daily.Amount, daily.RmbAmount),
                    AllocatedExpenseMonthlyData monthly => AllocatedExpenseSegmentViewModel.From(monthly.Reason, monthly.Currency, monthly.Amount, monthly.RmbAmount),
                    _ => null
                })
                .Where(segment => segment is not null)
                .Select(segment => segment!)
                .OrderBy(segment => segment.Reason, StringComparer.Ordinal)
                .ThenBy(segment => segment.Currency)
                .ToList();
            var total = Currency.RoundMoney(segments.Sum(segment => segment.RmbAmount ?? 0));
            return new AllocatedExpenseBucketViewModel
            {
                PeriodStart = periodStart,
                PeriodLabel = label,
                TotalRmb = total,
                TotalText = $"¥{total:N2}",
                DetailsText = BuildDetailsText(detailRecords),
                Segments = segments
            };
        }

        private static string BuildDetailsText(List<AllocatedExpenseRecordData> records)
        {
            if (records.Count == 0)
                return "";

            return String.Join("; ", records
                .OrderBy(record => record.Date)
                .ThenBy(record => record.AllocationStart)
                .ThenBy(record => record.Reason, StringComparer.Ordinal)
                .ThenBy(record => record.Currency)
                .ThenBy(record => record.RecordId)
                .Select(record =>
                {
                    var symbol = CurrencySummaryViewModel.FormatCurrencySymbol(record.Currency);
                    var period = $"{record.AllocationStart:yyyy-MM-dd}~{record.AllocationEnd:yyyy-MM-dd}";
                    var amount = $"{symbol}{MoneyText.FormatAmount(record.Amount)}";
                    return $"{record.Date:yyyy-MM-dd} {period} {amount} {record.Reason}";
                }));
        }
    }

    public class AllocatedExpenseSegmentViewModel
    {
        public string Reason { get; set; } = "";
        public CurrencyType Currency { get; set; }
        public decimal Amount { get; set; }
        public decimal? RmbAmount { get; set; }
        public decimal ChartValue => RmbAmount ?? 0;

        public static AllocatedExpenseSegmentViewModel From(
            string reason,
            CurrencyType currency,
            decimal amount,
            decimal? rmbAmount)
        {
            return new AllocatedExpenseSegmentViewModel
            {
                Reason = reason,
                Currency = currency,
                Amount = amount,
                RmbAmount = rmbAmount
            };
        }
    }

    public class AllocatedExpenseWindowShiftEventArgs : EventArgs
    {
        public int UnitOffset { get; set; }
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
        public string CurrentBalanceText { get; set; } = "";
        public string CurrencyDetails { get; set; } = "";
        public bool IsInflow => NetRmb > 0;
        public bool IsOutflow => NetRmb < 0;

        public static AccountNetFlowStatisticsViewModel From(AccountNetFlowStatistics statistics)
        {
            var missingRateText = statistics.HasMissingExchangeRate ? " + 未折算" : "";
            var missingBalanceRateText = statistics.HasMissingCurrentBalanceExchangeRate ? " + 未折算" : "";
            return new AccountNetFlowStatisticsViewModel
            {
                DisplayName = statistics.DisplayName,
                NetRmb = statistics.NetRmb,
                NetRmbText = $"{FormatSignedMoney(statistics.NetRmb, "¥")}{missingRateText}",
                CurrentBalanceText = $"{FormatMoney(statistics.CurrentBalanceRmb, "¥")}{missingBalanceRateText}",
                CurrencyDetails = String.Join("、", statistics.CurrencyTotals.Select(FormatCurrencyTotal)),
            };
        }

        private static string FormatCurrencyTotal(AccountNetFlowCurrencyTotal total)
        {
            var symbol = CurrencySummaryViewModel.FormatCurrencySymbol(total.Currency);
            return $"{total.Currency} {FormatSignedMoneyRounded(total.Amount, symbol)}";
        }

        private static string FormatMoney(decimal value, string symbol)
        {
            var sign = value < 0 ? "-" : "";
            return $"{sign}{symbol}{MoneyText.FormatAmount(Math.Abs(value))}";
        }

        private static string FormatSignedMoney(decimal value, string symbol)
        {
            var sign = value > 0 ? "+" : value < 0 ? "-" : "";
            return $"{sign}{symbol}{MoneyText.FormatAmount(Math.Abs(value))}";
        }

        private static string FormatSignedMoneyRounded(decimal value, string symbol)
        {
            var rounded = Decimal.Round(value, 0, MidpointRounding.ToEven);
            var sign = rounded > 0 ? "+" : rounded < 0 ? "-" : "";
            var amount = Math.Abs(rounded).ToString("#,0", CultureInfo.InvariantCulture);
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

    public class AllocatedExpenseChart : FrameworkElement
    {
        const double DragUnitMinimumPixels = 32;
        const double ChartLeftPadding = 64;
        const double ChartRightPadding = 36;

        public static readonly DependencyProperty ItemsProperty = DependencyProperty.Register(
            nameof(Items),
            typeof(IEnumerable),
            typeof(AllocatedExpenseChart),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender, OnItemsChanged));
        public static readonly DependencyProperty IsLineChartProperty = DependencyProperty.Register(
            nameof(IsLineChart),
            typeof(bool),
            typeof(AllocatedExpenseChart),
            new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

        static readonly string[] SegmentColors =
        [
            "#0F766E",
            "#2563EB",
            "#7C3AED",
            "#DB2777",
            "#F59E0B",
            "#0891B2",
            "#65A30D",
            "#DC2626",
            "#4F46E5"
        ];

        bool isDraggingWindow;
        Point dragStartPoint;

        public event EventHandler<AllocatedExpenseWindowShiftEventArgs>? WindowShiftRequested;

        public IEnumerable? Items
        {
            get => (IEnumerable?)GetValue(ItemsProperty);
            set => SetValue(ItemsProperty, value);
        }

        public bool IsLineChart
        {
            get => (bool)GetValue(IsLineChartProperty);
            set => SetValue(IsLineChartProperty, value);
        }

        private static void OnItemsChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            var chart = (AllocatedExpenseChart)dependencyObject;
            if (e.OldValue is INotifyCollectionChanged oldCollection)
                oldCollection.CollectionChanged -= chart.Items_CollectionChanged;
            if (e.NewValue is INotifyCollectionChanged newCollection)
                newCollection.CollectionChanged += chart.Items_CollectionChanged;
            chart.InvalidateVisual();
        }

        private void Items_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            InvalidateVisual();
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            isDraggingWindow = true;
            dragStartPoint = e.GetPosition(this);
            CaptureMouse();
            e.Handled = true;
        }

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonUp(e);
            if (!isDraggingWindow)
                return;

            isDraggingWindow = false;
            ReleaseMouseCapture();
            var unitOffset = CalculateDragUnitOffset(e.GetPosition(this));
            if (unitOffset != 0)
                WindowShiftRequested?.Invoke(this, new AllocatedExpenseWindowShiftEventArgs { UnitOffset = unitOffset });
            e.Handled = true;
        }

        protected override void OnLostMouseCapture(MouseEventArgs e)
        {
            base.OnLostMouseCapture(e);
            isDraggingWindow = false;
        }

        private int CalculateDragUnitOffset(Point endPoint)
        {
            var deltaX = endPoint.X - dragStartPoint.X;
            var buckets = Items?.Cast<AllocatedExpenseBucketViewModel>().ToList() ?? [];
            if (buckets.Count == 0)
                return 0;

            var chartWidth = Math.Max(1, ActualWidth - ChartLeftPadding - ChartRightPadding);
            var pixelsPerUnit = Math.Max(DragUnitMinimumPixels, chartWidth / buckets.Count);
            if (Math.Abs(deltaX) < pixelsPerUnit * 0.5)
                return 0;

            return (int)Math.Round(-deltaX / pixelsPerUnit, MidpointRounding.AwayFromZero);
        }

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);
            var buckets = Items?.Cast<AllocatedExpenseBucketViewModel>().ToList() ?? [];
            if (buckets.Count == 0 || buckets.All(bucket => bucket.TotalRmb == 0))
            {
                DrawText(dc, "暂无均摊支出", 14, "#94A3B8", ActualWidth / 2 - 42, ActualHeight / 2 - 10);
                return;
            }

            var width = ActualWidth;
            var height = ActualHeight;
            var left = 64.0;
            var top = 44.0;
            var right = 36.0;
            var bottom = 34.0;
            var chartWidth = Math.Max(1, width - left - right);
            var chartHeight = Math.Max(1, height - top - bottom);
            var axisMax = CalculateAxisMax(buckets.Max(bucket => bucket.TotalRmb));
            DrawAxes(dc, left, top, chartWidth, chartHeight, axisMax);
            DrawLegend(dc, buckets, width);

            if (IsLineChart)
                DrawLineChart(dc, buckets, left, top, chartWidth, chartHeight, axisMax);
            else
                DrawStackedBars(dc, buckets, left, top, chartWidth, chartHeight, axisMax);

            DrawLabels(dc, buckets, left, top + chartHeight, chartWidth);
        }

        private static void DrawAxes(
            DrawingContext dc,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            decimal axisMax)
        {
            var gridPen = new Pen(ParseBrush("#E2E8F0"), 1);
            var chartRight = left + chartWidth;
            var axisStep = axisMax / 4;
            for (var i = 0; i <= 4; i++)
            {
                var value = axisStep * i;
                var y = top + chartHeight - chartHeight * i / 4.0;
                dc.DrawLine(gridPen, new Point(left, y), new Point(chartRight, y));
                DrawRightAlignedText(dc, value.ToString("N0", CultureInfo.InvariantCulture), 11, "#94A3B8", left - 8, y - 8);
            }

            dc.DrawLine(gridPen, new Point(left, top), new Point(left, top + chartHeight));
            dc.DrawLine(gridPen, new Point(left, top + chartHeight), new Point(chartRight, top + chartHeight));
        }

        private static void DrawStackedBars(
            DrawingContext dc,
            List<AllocatedExpenseBucketViewModel> buckets,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            decimal axisMax)
        {
            var groupWidth = chartWidth / buckets.Count;
            var barWidth = Math.Min(28, Math.Max(8, groupWidth * 0.42));
            for (var i = 0; i < buckets.Count; i++)
            {
                var x = left + groupWidth * i + (groupWidth - barWidth) / 2;
                var y = top + chartHeight;
                foreach (var segment in buckets[i].Segments.Where(segment => segment.ChartValue > 0))
                {
                    var height = Math.Max(1, chartHeight * (double)(segment.ChartValue / axisMax));
                    y -= height;
                    dc.DrawRoundedRectangle(
                        GetReasonBrush(segment.Reason),
                        null,
                        new Rect(x, y, barWidth, height),
                        2,
                        2);
                }

                DrawValueLabel(dc, buckets[i].TotalRmb, left + groupWidth * i + groupWidth / 2, y - 17);
            }
        }

        private static void DrawLineChart(
            DrawingContext dc,
            List<AllocatedExpenseBucketViewModel> buckets,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            decimal axisMax)
        {
            var pen = new Pen(ParseBrush("#0F766E"), 2);
            var fill = ParseBrush("#0F766E");
            var points = buckets.Select((bucket, index) =>
            {
                var x = buckets.Count == 1
                    ? left + chartWidth / 2
                    : left + chartWidth * index / (buckets.Count - 1);
                var y = top + chartHeight - chartHeight * (double)(bucket.TotalRmb / axisMax);
                return new Point(x, y);
            }).ToList();

            for (var i = 1; i < points.Count; i++)
                dc.DrawLine(pen, points[i - 1], points[i]);

            for (var i = 0; i < points.Count; i++)
            {
                dc.DrawEllipse(fill, null, points[i], 3.5, 3.5);
                if (buckets[i].TotalRmb > 0)
                    DrawValueLabel(dc, buckets[i].TotalRmb, points[i].X, points[i].Y - 20);
            }
        }

        private static void DrawLabels(
            DrawingContext dc,
            List<AllocatedExpenseBucketViewModel> buckets,
            double left,
            double baselineY,
            double chartWidth)
        {
            var groupWidth = chartWidth / buckets.Count;
            for (var i = 0; i < buckets.Count; i++)
            {
                if (buckets.Count > 18 && i % 2 == 1)
                    continue;

                var x = left + groupWidth * i + groupWidth / 2;
                var text = CreateText(buckets[i].PeriodLabel, 11, ParseBrush("#64748B"));
                dc.DrawText(text, new Point(x - text.Width / 2, baselineY + 12));
            }
        }

        private static void DrawLegend(DrawingContext dc, List<AllocatedExpenseBucketViewModel> buckets, double width)
        {
            var reasons = buckets
                .SelectMany(bucket => bucket.Segments)
                .GroupBy(segment => segment.Reason)
                .Select(group => new { Reason = group.Key, Total = group.Sum(segment => segment.ChartValue) })
                .Where(item => item.Total > 0)
                .OrderByDescending(item => item.Total)
                .Take(5)
                .ToList();
            var x = Math.Max(70, width - 520);
            foreach (var item in reasons)
            {
                dc.DrawEllipse(GetReasonBrush(item.Reason), null, new Point(x, 13), 4, 4);
                DrawText(dc, TrimText(item.Reason, 10), 12, "#334155", x + 10, 5);
                x += 96;
            }
        }

        private static void DrawValueLabel(DrawingContext dc, decimal value, double centerX, double y)
        {
            if (value <= 0)
                return;

            var text = CreateText(value.ToString("N0", CultureInfo.InvariantCulture), 10, ParseBrush("#0F172A"));
            var x = centerX - text.Width / 2;
            dc.DrawRoundedRectangle(
                new SolidColorBrush(Color.FromArgb(230, 255, 255, 255)),
                null,
                new Rect(x - 3, y - 1, text.Width + 6, text.Height + 2),
                4,
                4);
            dc.DrawText(text, new Point(x, y));
        }

        private static decimal CalculateAxisMax(decimal maxValue)
        {
            if (maxValue <= 0)
                return 1;

            var rawStep = (double)maxValue / 4.0;
            var magnitude = Math.Pow(10, Math.Floor(Math.Log10(Math.Max(1, rawStep))));
            var step = (decimal)(Math.Ceiling(rawStep / magnitude) * magnitude);
            return Math.Max(1, step * 4);
        }

        private static Brush GetReasonBrush(string reason)
        {
            var hash = 0;
            foreach (var c in reason)
                hash = unchecked(hash * 31 + c);
            var index = Math.Abs(hash) % SegmentColors.Length;
            return ParseBrush(SegmentColors[index]);
        }

        private static string TrimText(string text, int maxLength)
        {
            return text.Length <= maxLength ? text : text[..Math.Max(1, maxLength - 1)] + "…";
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
