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
using System.Windows.Threading;
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
        readonly DispatcherTimer importStatusTimer = new() { Interval = TimeSpan.FromSeconds(1) };
        Forms.NotifyIcon? trayIcon;
        Drawing.Icon? trayIconImage;
        bool isExitRequested;
        bool isLoadingRecordDetails;
        bool isAdjustingDetailSelection;
        bool isLoadingAllocatedExpenses;
        bool isAdjustingAllocatedExpenseRange;

        public MainWindow()
        {
#if DEBUG
            AllocConsole();
#endif
            InitializeComponent();
            InitializeTrayIcon();
            importStatusTimer.Tick += ImportStatusTimer_Tick;
            fetcher.RunSchedule();
            RefreshImportRuntimeStatus();
            importStatusTimer.Start();
            _ = LoadDashboardAsync();
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as DashboardViewModel;
            if (viewModel is not null && !ConfirmDiscardRecordDetailChanges(viewModel))
                return;

            await LoadDashboardAsync(viewModel?.DetailStartDate, viewModel?.DetailEndDate, viewModel?.SelectedDetailAccountFilterKey);
        }

        private void AddRecordDetail_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not DashboardViewModel viewModel)
                return;

            var accountName = viewModel.GetDefaultNewRecordAccountName();
            viewModel.RecordDetails.Insert(0, RecordDetailRowViewModel.CreateNew(accountName));
            viewModel.SetDetailStatus("新增记录尚未保存");
        }

        private async void DetailDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isLoadingRecordDetails || isAdjustingDetailSelection)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            if (!ConfirmDiscardRecordDetailChanges(viewModel))
            {
                RestoreLoadedDetailSelection(viewModel);
                return;
            }

            await LoadRecordDetailsAsync(viewModel);
        }

        private async void DetailAccount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isLoadingRecordDetails || isAdjustingDetailSelection)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;

            if (!ConfirmDiscardRecordDetailChanges(viewModel))
            {
                RestoreLoadedDetailSelection(viewModel);
                return;
            }

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

            var movedWithinLoadedWindow = false;
            isAdjustingAllocatedExpenseRange = true;
            try
            {
                movedWithinLoadedWindow = viewModel.TryMoveAllocatedExpenseWindowWithinLoaded(e.UnitOffset);
                if (!movedWithinLoadedWindow)
                    viewModel.MoveAllocatedExpenseWindow(e.UnitOffset);
            }
            finally
            {
                isAdjustingAllocatedExpenseRange = false;
            }
            if (!movedWithinLoadedWindow)
                await LoadAllocatedExpensesAsync(viewModel);
        }

        private async void AllocatedExpenseChart_PeriodDoubleClicked(object sender, AllocatedExpensePeriodDoubleClickEventArgs e)
        {
            if (DataContext is not DashboardViewModel viewModel)
                return;
            if (!ConfirmDiscardRecordDetailChanges(viewModel))
                return;

            isAdjustingDetailSelection = true;
            try
            {
                viewModel.DetailStartDate = e.StartDate;
                viewModel.DetailEndDate = e.EndDate;
            }
            finally
            {
                isAdjustingDetailSelection = false;
            }

            DetailTab.IsSelected = true;
            await LoadRecordDetailsAsync(viewModel);
        }

        private async void MonthlyFlowChart_WindowShiftRequested(object sender, AllocatedExpenseWindowShiftEventArgs e)
        {
            if (e.UnitOffset == 0)
                return;
            if (DataContext is not DashboardViewModel viewModel)
                return;
            if (!ConfirmDiscardRecordDetailChanges(viewModel))
                return;

            await LoadDashboardAsync(
                viewModel.DetailStartDate,
                viewModel.DetailEndDate,
                viewModel.SelectedDetailAccountFilterKey,
                viewModel.MonthlyFlowStartMonth.AddMonths(e.UnitOffset));
        }

        private async void MonthlyFlowChart_PeriodDoubleClicked(object sender, AllocatedExpensePeriodDoubleClickEventArgs e)
        {
            if (DataContext is not DashboardViewModel viewModel)
                return;
            if (!ConfirmDiscardRecordDetailChanges(viewModel))
                return;

            isAdjustingDetailSelection = true;
            try
            {
                viewModel.DetailStartDate = e.StartDate;
                viewModel.DetailEndDate = e.EndDate;
            }
            finally
            {
                isAdjustingDetailSelection = false;
            }

            DetailTab.IsSelected = true;
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

            CommitRecordDetailsGridEdits();
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
                await LoadDashboardAsync(viewModel.DetailStartDate, viewModel.DetailEndDate, viewModel.SelectedDetailAccountFilterKey);
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

        private bool ConfirmDiscardRecordDetailChanges(DashboardViewModel viewModel)
        {
            CommitRecordDetailsGridEdits();
            if (!viewModel.HasPendingDetailChanges())
                return true;

            var result = MessageBox.Show(
                this,
                "详情页有未保存的修改，继续操作会丢弃这些修改。是否继续？",
                "确认丢弃详情修改",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            return result == MessageBoxResult.Yes;
        }

        private void CommitRecordDetailsGridEdits()
        {
            RecordDetailsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            RecordDetailsGrid.CommitEdit(DataGridEditingUnit.Row, true);
        }

        private void RestoreLoadedDetailSelection(DashboardViewModel viewModel)
        {
            isAdjustingDetailSelection = true;
            try
            {
                viewModel.RestoreLoadedDetailSelection();
            }
            finally
            {
                isAdjustingDetailSelection = false;
            }
        }

        private async Task LoadDashboardAsync(
            DateTime? detailStartDate = null,
            DateTime? detailEndDate = null,
            string? detailAccountFilterKey = null,
            DateTime? monthlyFlowStartMonth = null)
        {
            try
            {
                var previousViewModel = DataContext as DashboardViewModel;
                var config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
                var database = new DatabaseUtil(config);
                var effectiveMonthlyFlowStartMonth = NormalizeMonth(
                    monthlyFlowStartMonth ??
                    previousViewModel?.MonthlyFlowStartMonth ??
                    DateTime.Today.AddMonths(-11));
                var data = database.GetDashboardData(DateTime.Today, effectiveMonthlyFlowStartMonth);
                if (data.MissingExchangeRateCurrencies.Count > 0)
                {
                    using var pubWeb = new PubWebUtil(config, database);
                    await pubWeb.FetchExchangeRates(data.MissingExchangeRateCurrencies);
                    data = database.GetDashboardData(DateTime.Today, effectiveMonthlyFlowStartMonth);
                    if (data.MissingExchangeRateCurrencies.Count > 0)
                    {
                        throw new InvalidOperationException(
                            $"缺少汇率：{String.Join("、", data.MissingExchangeRateCurrencies)}");
                    }
                }

                var viewModel = DashboardViewModel.From(data);
                if (previousViewModel is not null)
                    viewModel.CopyDashboardSettingsFrom(previousViewModel);
                if (detailStartDate.HasValue)
                    viewModel.DetailStartDate = detailStartDate.Value;
                if (detailEndDate.HasValue)
                    viewModel.DetailEndDate = detailEndDate.Value;
                DataContext = viewModel;
                RefreshImportRuntimeStatus(viewModel);
                await LoadRecordDetailsAsync(viewModel, detailAccountFilterKey);
                await LoadAllocatedExpensesAsync(viewModel);
            }
            catch (Exception e)
            {
                DataContext = DashboardViewModel.FromError(e.Message);
                RefreshImportRuntimeStatus();
            }
        }

        private static DateTime NormalizeMonth(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        private async Task LoadRecordDetailsAsync(DashboardViewModel? viewModel = null, string? preferredDetailAccountFilterKey = null)
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
                var defaultDetailAccountName = accounts
                    .Where(account => account.usage != AccountUsage.Undetermined)
                    .OrderBy(account => account.Id)
                    .Select(account => account.name)
                    .FirstOrDefault()
                    ?? accounts
                        .OrderBy(account => account.name, StringComparer.OrdinalIgnoreCase)
                        .Select(account => account.name)
                        .FirstOrDefault()
                    ?? "";
                viewModel.SetDetailAccountFilters(
                    accounts,
                    preferredDetailAccountFilterKey ?? viewModel.SelectedDetailAccountFilterKey,
                    defaultDetailAccountName);
                var details = database.GetRecordDetails(viewModel.DetailStartDate, viewModel.DetailEndDate, viewModel.SelectedDetailAccountNames)
                    .Select(RecordDetailRowViewModel.From)
                    .ToList();
                LoadDetailAccountBalances(viewModel, database);
                viewModel.SetRecordDetails(details);
                viewModel.MarkLoadedDetailSelection();
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
                    var visibleMonthCount = CountMonths(firstMonth, endMonthExclusive);
                    var loadedFirstMonth = firstMonth.AddMonths(-visibleMonthCount);
                    var loadedEndMonthExclusive = endMonthExclusive.AddMonths(visibleMonthCount);
                    var items = database.GetAllocatedExpenseMonthly(loadedFirstMonth, loadedEndMonthExclusive);
                    var details = database.GetAllocatedExpenseRecordDetails(loadedFirstMonth, loadedEndMonthExclusive);
                    viewModel.SetLoadedAllocatedExpenseBuckets(AllocatedExpenseBucketViewModel.FromMonthly(items, details, loadedFirstMonth, loadedEndMonthExclusive));
                    viewModel.SetAllocatedExpenseStatus(viewModel.BuildAllocatedExpenseStatus());
                }
                else
                {
                    var startDate = viewModel.AllocatedExpenseStartDate.Date;
                    var endDate = viewModel.AllocatedExpenseEndDate.Date;
                    var visibleDayCount = Math.Max(1, (endDate - startDate).Days + 1);
                    var loadedStartDate = startDate.AddDays(-visibleDayCount);
                    var loadedEndDate = endDate.AddDays(visibleDayCount);
                    var items = database.GetAllocatedExpenseDaily(
                        loadedStartDate,
                        loadedEndDate.AddDays(1));
                    var details = database.GetAllocatedExpenseRecordDetails(
                        loadedStartDate,
                        loadedEndDate.AddDays(1));
                    viewModel.SetLoadedAllocatedExpenseBuckets(AllocatedExpenseBucketViewModel.FromDaily(
                        items,
                        details,
                        loadedStartDate,
                        loadedEndDate));
                    viewModel.SetAllocatedExpenseStatus(viewModel.BuildAllocatedExpenseStatus());
                }
            }
            catch (Exception e)
            {
                viewModel.ClearAllocatedExpenseBuckets();
                viewModel.SetAllocatedExpenseStatus($"均摊支出读取失败：{e.Message}");
            }
            finally
            {
                isLoadingAllocatedExpenses = false;
            }
        }

        private static int CountMonths(DateTime startInclusive, DateTime endExclusive)
        {
            return Math.Max(
                1,
                (endExclusive.Year - startInclusive.Year) * 12 + endExclusive.Month - startInclusive.Month);
        }

        private void ImportStatusTimer_Tick(object? sender, EventArgs e)
        {
            RefreshImportRuntimeStatus();
        }

        private void RefreshImportRuntimeStatus(DashboardViewModel? viewModel = null)
        {
            viewModel ??= DataContext as DashboardViewModel;
            viewModel?.SetImportRuntimeStatus(fetcher.GetRuntimeStatus());
        }

        private static void LoadDetailAccountBalances(DashboardViewModel viewModel, DatabaseUtil database)
        {
            if (String.IsNullOrWhiteSpace(viewModel.SelectedDetailAccountName))
            {
                viewModel.SetDetailAccountBalances([], null);
                return;
            }

            var balances = database.GetAccountBalanceDetails(viewModel.SelectedDetailAccountName)
                .Select(AccountBalanceRowViewModel.From)
                .ToList();
            viewModel.SetDetailAccountBalances(balances, viewModel.SelectedDetailAccountName);
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

            if (DataContext is DashboardViewModel viewModel && !ConfirmDiscardRecordDetailChanges(viewModel))
            {
                e.Cancel = true;
                isExitRequested = false;
                if (trayIcon is not null)
                    trayIcon.Visible = true;
                return;
            }

            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            importStatusTimer.Stop();
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
        DateTime loadedDetailStartDate = DateTime.Today.AddDays(-30);
        DateTime loadedDetailEndDate = DateTime.Today;
        DateTime monthlyFlowStartMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-11);
        DateTime allocatedExpenseStartDate = DateTime.Today.AddDays(-14);
        DateTime allocatedExpenseEndDate = DateTime.Today;
        bool showInvestmentByHolding;
        DetailAccountFilterViewModel? selectedDetailAccountFilter;
        string? loadedDetailAccountFilterKey;
        string? loadedDetailBalanceAccountName;
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
        public List<StatementImportSummaryViewModel> LatestStatementImports { get; set; } = [];
        public string ImportRuntimeText { get; private set; } = "导入未启用";
        public ObservableCollection<RecordDetailRowViewModel> RecordDetails { get; } = [];
        public ObservableCollection<AccountBalanceRowViewModel> DetailAccountBalances { get; } = [];
        public ObservableCollection<AllocatedExpenseBucketViewModel> AllocatedExpenseBuckets { get; } = [];
        public ObservableCollection<AllocatedExpenseBucketViewModel> LoadedAllocatedExpenseBuckets { get; } = [];
        public List<DetailAccountFilterViewModel> DetailAccountFilters { get; private set; } = [];
        public List<CurrencyType> DetailCurrencyTypes { get; } = Enum.GetValues<CurrencyType>().ToList();
        public string DetailStatusText { get; private set; } = "";
        public string AllocatedExpenseStatusText { get; private set; } = "";
        public string AllocatedExpenseTotalText => $"¥{AllocatedExpenseBuckets.Sum(bucket => bucket.TotalRmb):N2}";
        public DateTime MonthlyFlowStartMonth
        {
            get => monthlyFlowStartMonth;
            set
            {
                value = new DateTime(value.Year, value.Month, 1);
                if (monthlyFlowStartMonth == value)
                    return;

                monthlyFlowStartMonth = value;
                OnPropertyChanged(nameof(MonthlyFlowStartMonth));
                OnPropertyChanged(nameof(MonthlyFlowRangeText));
            }
        }
        public string MonthlyFlowRangeText => $"{MonthlyFlowStartMonth:yyyy-MM} - {MonthlyFlowStartMonth.AddMonths(11):yyyy-MM}";
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

        public DetailAccountFilterViewModel? SelectedDetailAccountFilter
        {
            get => selectedDetailAccountFilter;
            set
            {
                if (ReferenceEquals(selectedDetailAccountFilter, value))
                    return;

                selectedDetailAccountFilter = value;
                OnPropertyChanged(nameof(SelectedDetailAccountFilter));
                OnPropertyChanged(nameof(SelectedDetailAccountName));
                OnPropertyChanged(nameof(SelectedDetailAccountNames));
                OnPropertyChanged(nameof(SelectedDetailAccountFilterKey));
            }
        }

        public string? SelectedDetailAccountName => SelectedDetailAccountFilter?.Kind == DetailAccountFilterKind.Account
            ? SelectedDetailAccountFilter.AccountNames.FirstOrDefault()
            : null;

        public IReadOnlyList<string>? SelectedDetailAccountNames => SelectedDetailAccountFilter?.Kind == DetailAccountFilterKind.All
            ? null
            : SelectedDetailAccountFilter?.AccountNames;

        public string? SelectedDetailAccountFilterKey => SelectedDetailAccountFilter?.Key;

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
                SnapshotTimeText = "最新",
                ReasonTabHeader = "分类",
                TotalAssets = TotalAssetsViewModel.From(data.TotalAssetsRmb),
                MonthlyFlowStartMonth = data.MonthlyFlowStartMonth,
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
                    .ToList(),
                LatestStatementImports = data.LatestStatementImports
                    .Select(StatementImportSummaryViewModel.From)
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

        public void CopyDashboardSettingsFrom(DashboardViewModel source)
        {
            showSingleCurrencyMonthly = source.ShowSingleCurrencyMonthly;
            selectedMonthlyAccount = MonthlyAccounts.FirstOrDefault(account =>
                String.Equals(account.DisplayName, source.SelectedMonthlyAccount?.DisplayName, StringComparison.Ordinal)) ??
                MonthlyAccounts.FirstOrDefault();
            showAllocatedExpenseLineChart = source.ShowAllocatedExpenseLineChart;
            showAllocatedExpenseMonthly = source.ShowAllocatedExpenseMonthly;
            allocatedExpenseStartDate = source.AllocatedExpenseStartDate;
            allocatedExpenseEndDate = source.AllocatedExpenseEndDate;
            OnPropertyChanged(nameof(ShowSingleCurrencyMonthly));
            OnPropertyChanged(nameof(SelectedMonthlyAccount));
            OnPropertyChanged(nameof(VisibleMonthlySeries));
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

        public bool TryMoveAllocatedExpenseWindowWithinLoaded(int unitOffset)
        {
            if (unitOffset == 0)
                return true;

            var targetStartDate = ShowAllocatedExpenseMonthly
                ? allocatedExpenseStartDate.AddMonths(unitOffset)
                : allocatedExpenseStartDate.AddDays(unitOffset);
            var targetEndDate = ShowAllocatedExpenseMonthly
                ? allocatedExpenseEndDate.AddMonths(unitOffset)
                : allocatedExpenseEndDate.AddDays(unitOffset);
            if (!IsAllocatedExpenseWindowWithinLoaded(targetStartDate, targetEndDate))
                return false;

            MoveAllocatedExpenseWindow(unitOffset);
            RefreshVisibleAllocatedExpenseBucketsFromLoaded();
            SetAllocatedExpenseStatus(BuildAllocatedExpenseStatus());
            return true;
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

        public void SetImportRuntimeStatus(FetchRuntimeStatus status)
        {
            var now = DateTime.Now;
            var lastText = status.LastFetchTime.HasValue
                ? status.LastFetchTime.Value.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture)
                : "无";
            var failurePrefix = status.HasImportFailureMarker ? "导入失败；" : "";
            if (!String.IsNullOrWhiteSpace(status.CurrentTaskName))
            {
                ImportRuntimeText = $"{failurePrefix}导入：{status.CurrentTaskName}{FormatElapsedSuffix(status.CurrentTaskStartedAt, now)}；上次：{lastText}";
            }
            else if (status.IsScheduledFetchEnabled && status.NextFetchTime.HasValue)
            {
                ImportRuntimeText = $"{failurePrefix}上次：{lastText}；下次：{status.NextFetchTime.Value:MM-dd HH:mm}";
            }
            else
            {
                ImportRuntimeText = status.IsScheduledFetchEnabled
                    ? $"{failurePrefix}上次：{lastText}；下次：待定"
                    : $"{failurePrefix}上次：{lastText}；导入未启用";
            }
            OnPropertyChanged(nameof(ImportRuntimeText));
        }

        private static string FormatElapsedSuffix(DateTime? startedAt, DateTime now)
        {
            if (!startedAt.HasValue)
                return "";

            var elapsed = now - startedAt.Value;
            if (elapsed < TimeSpan.Zero)
                elapsed = TimeSpan.Zero;
            var totalHours = (int)elapsed.TotalHours;
            return $" {totalHours:D2}:{elapsed.Minutes:D2}:{elapsed.Seconds:D2}";
        }

        public void SetDetailAccountFilters(IEnumerable<Account> accounts, string? preferredFilterKey, string defaultAccountName)
        {
            var accountList = accounts
                .OrderBy(account => account.name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var filters = new List<DetailAccountFilterViewModel>();
            if (accountList.Count > 0)
                filters.Add(DetailAccountFilterViewModel.All(accountList.Select(account => account.name)));
            filters.AddRange(accountList
                .GroupBy(account => GetAccountType(account.name), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .Select(group => DetailAccountFilterViewModel.AccountType(
                    group.Key,
                    group.Select(account => account.name).OrderBy(name => name, StringComparer.OrdinalIgnoreCase))));
            filters.AddRange(accountList
                .Select(account => DetailAccountFilterViewModel.Account(account.name)));

            DetailAccountFilters = filters;
            OnPropertyChanged(nameof(DetailAccountFilters));
            SelectedDetailAccountFilter =
                filters.FirstOrDefault(filter => String.Equals(filter.Key, preferredFilterKey, StringComparison.Ordinal)) ??
                filters.FirstOrDefault(filter => filter.Kind == DetailAccountFilterKind.All) ??
                filters.FirstOrDefault(filter => filter.Kind == DetailAccountFilterKind.Account &&
                    String.Equals(filter.AccountNames.FirstOrDefault(), defaultAccountName, StringComparison.OrdinalIgnoreCase)) ??
                filters.FirstOrDefault();
        }

        public void MarkLoadedDetailSelection()
        {
            loadedDetailStartDate = DetailStartDate;
            loadedDetailEndDate = DetailEndDate;
            loadedDetailAccountFilterKey = SelectedDetailAccountFilterKey;
        }

        public void RestoreLoadedDetailSelection()
        {
            DetailStartDate = loadedDetailStartDate;
            DetailEndDate = loadedDetailEndDate;
            SelectDetailAccountFilter(loadedDetailAccountFilterKey);
        }

        public bool HasPendingDetailChanges()
        {
            return GetPendingRecordChanges().Count > 0 || GetPendingBalanceCorrections().Count > 0;
        }

        private void SelectDetailAccountFilter(string? key)
        {
            SelectedDetailAccountFilter =
                DetailAccountFilters.FirstOrDefault(filter => String.Equals(filter.Key, key, StringComparison.Ordinal)) ??
                DetailAccountFilters.FirstOrDefault();
        }

        public string GetDefaultNewRecordAccountName()
        {
            if (SelectedDetailAccountFilter?.Kind == DetailAccountFilterKind.Account)
                return SelectedDetailAccountFilter.AccountNames.FirstOrDefault() ?? "";
            if (SelectedDetailAccountFilter?.Kind == DetailAccountFilterKind.AccountType)
                return SelectedDetailAccountFilter.AccountNames.FirstOrDefault() ?? "";

            return DetailAccountFilters
                .Where(filter => filter.Kind == DetailAccountFilterKind.Account)
                .Select(filter => filter.AccountNames.FirstOrDefault())
                .FirstOrDefault(accountName => !String.IsNullOrWhiteSpace(accountName))
                ?? "";
        }

        private static string GetAccountType(string accountName)
        {
            var normalized = accountName.Trim();
            var separatorIndex = normalized.IndexOf('_');
            return separatorIndex > 0 ? normalized[..separatorIndex] : normalized;
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

        public void SetLoadedAllocatedExpenseBuckets(IEnumerable<AllocatedExpenseBucketViewModel> buckets)
        {
            LoadedAllocatedExpenseBuckets.Clear();
            foreach (var bucket in buckets)
                LoadedAllocatedExpenseBuckets.Add(bucket);
            OnPropertyChanged(nameof(LoadedAllocatedExpenseBuckets));
            RefreshVisibleAllocatedExpenseBucketsFromLoaded();
        }

        public void ClearAllocatedExpenseBuckets()
        {
            LoadedAllocatedExpenseBuckets.Clear();
            AllocatedExpenseBuckets.Clear();
            OnPropertyChanged(nameof(LoadedAllocatedExpenseBuckets));
            OnPropertyChanged(nameof(AllocatedExpenseBuckets));
            OnPropertyChanged(nameof(AllocatedExpenseTotalText));
        }

        public void RefreshVisibleAllocatedExpenseBucketsFromLoaded()
        {
            AllocatedExpenseBuckets.Clear();
            foreach (var bucket in GetVisibleAllocatedExpenseBucketsFromLoaded())
                AllocatedExpenseBuckets.Add(bucket);
            OnPropertyChanged(nameof(AllocatedExpenseBuckets));
            OnPropertyChanged(nameof(AllocatedExpenseTotalText));
        }

        public string BuildAllocatedExpenseStatus()
        {
            return ShowAllocatedExpenseMonthly
                ? $"共 {AllocatedExpenseBuckets.Count} 个月"
                : $"共 {AllocatedExpenseBuckets.Count} 天";
        }

        public void SetAllocatedExpenseStatus(string status)
        {
            AllocatedExpenseStatusText = status;
            OnPropertyChanged(nameof(AllocatedExpenseStatusText));
        }

        private bool IsAllocatedExpenseWindowWithinLoaded(DateTime startDate, DateTime endDate)
        {
            if (LoadedAllocatedExpenseBuckets.Count == 0)
                return false;

            var loadedStart = LoadedAllocatedExpenseBuckets[0].PeriodStart.Date;
            var loadedEnd = LoadedAllocatedExpenseBuckets[LoadedAllocatedExpenseBuckets.Count - 1].PeriodStart.Date;
            var targetStart = ShowAllocatedExpenseMonthly
                ? GetMonthStart(startDate)
                : startDate.Date;
            var targetEnd = ShowAllocatedExpenseMonthly
                ? GetMonthStart(endDate)
                : endDate.Date;
            return loadedStart <= targetStart && loadedEnd >= targetEnd;
        }

        private List<AllocatedExpenseBucketViewModel> GetVisibleAllocatedExpenseBucketsFromLoaded()
        {
            if (ShowAllocatedExpenseMonthly)
            {
                var firstMonth = GetMonthStart(AllocatedExpenseStartDate);
                var endMonthExclusive = GetMonthStart(AllocatedExpenseEndDate).AddMonths(1);
                return LoadedAllocatedExpenseBuckets
                    .Where(bucket => bucket.PeriodStart >= firstMonth && bucket.PeriodStart < endMonthExclusive)
                    .ToList();
            }

            return LoadedAllocatedExpenseBuckets
                .Where(bucket => bucket.PeriodStart >= AllocatedExpenseStartDate.Date && bucket.PeriodStart <= AllocatedExpenseEndDate.Date)
                .ToList();
        }

        private static DateTime GetMonthStart(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        public void SetDetailAccountBalances(IEnumerable<AccountBalanceRowViewModel> balances, string? accountName)
        {
            loadedDetailBalanceAccountName = accountName;
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
            if (String.IsNullOrWhiteSpace(loadedDetailBalanceAccountName))
                return [];

            return DetailAccountBalances
                .Select(balance => balance.GetPendingCorrection(loadedDetailBalanceAccountName))
                .Where(change => change is not null)
                .Select(change => change!)
                .ToList();
        }
    }

    public class StatementImportSummaryViewModel
    {
        public StatementImportProvider Provider { get; set; }
        public string ProviderText => Provider.ToString();
        public int? Id { get; set; }
        public DateTime? Time { get; set; }
        public string TimeText => Time.HasValue
            ? Time.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : "无";
        public string StatementKey { get; set; } = "";
        public string StatementKeyText => String.IsNullOrWhiteSpace(StatementKey) ? "无" : StatementKey;
        public bool HasImport => Id.HasValue;

        public static StatementImportSummaryViewModel From(StatementImportSummaryData data)
        {
            return new StatementImportSummaryViewModel
            {
                Provider = data.Provider,
                Id = data.Id,
                Time = data.Time,
                StatementKey = data.StatementKey
            };
        }
    }

    public enum DetailAccountFilterKind
    {
        All,
        AccountType,
        Account
    }

    public class DetailAccountFilterViewModel
    {
        public DetailAccountFilterKind Kind { get; init; }
        public string Key { get; init; } = "";
        public string DisplayName { get; init; } = "";
        public IReadOnlyList<string> AccountNames { get; init; } = [];

        public static DetailAccountFilterViewModel All(IEnumerable<string> accountNames)
        {
            return new DetailAccountFilterViewModel
            {
                Kind = DetailAccountFilterKind.All,
                Key = "all",
                DisplayName = "所有账户",
                AccountNames = accountNames.ToList()
            };
        }

        public static DetailAccountFilterViewModel AccountType(string accountType, IEnumerable<string> accountNames)
        {
            return new DetailAccountFilterViewModel
            {
                Kind = DetailAccountFilterKind.AccountType,
                Key = $"type:{accountType}",
                DisplayName = $"类型：{BuildAccountTypeDisplayName(accountType)}",
                AccountNames = accountNames.ToList()
            };
        }

        public static DetailAccountFilterViewModel Account(string accountName)
        {
            return new DetailAccountFilterViewModel
            {
                Kind = DetailAccountFilterKind.Account,
                Key = $"account:{accountName}",
                DisplayName = accountName,
                AccountNames = [accountName]
            };
        }

        private static string BuildAccountTypeDisplayName(string accountType)
        {
            return String.IsNullOrWhiteSpace(accountType)
                ? "未分类账户"
                : $"{accountType} 账户";
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
                StatusText = "最新",
                TotalAssets = TotalAssetsViewModel.From(totalAssetsRmb),
                CurrencySummaries = currencySummaries.Select(CurrencySummaryViewModel.From).ToList()
            };
        }

        private static string BuildStatusText(AssetSummaryPoint point, string dateLabel)
        {
            if (!point.HasData)
                return $"{dateLabel} 缺少快照";

            if (point.IsToday)
                return "最新";

            var snapshotTime = point.SnapshotTime ?? point.Date;
            return $"基于 {snapshotTime:yyyy-MM-dd} 快照计算";
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
            var text = MoneyText.FromWan(value, "¥ ");
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
        private const decimal WanAbbreviationThreshold = 10000m;

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

        public static MoneyText FromWan(decimal value, string prefix = "")
        {
            var sign = value < 0 ? "-" : "";
            var absoluteValue = Math.Abs(value);
            var exact = $"{sign}{prefix}{FormatAmount(absoluteValue)}";
            if (absoluteValue < WanAbbreviationThreshold)
                return new MoneyText(exact, "");

            var abbreviatedValue = Decimal.Round(absoluteValue / WanAbbreviationThreshold, 2, MidpointRounding.ToEven);
            var abbreviatedText = $"{sign}{prefix}{FormatAmount(abbreviatedValue)}万";
            return new MoneyText(abbreviatedText, abbreviatedText == exact ? "" : exact);
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

    public class AllocatedExpensePeriodDoubleClickEventArgs : EventArgs
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public static AllocatedExpensePeriodDoubleClickEventArgs From(DateTime periodStart, bool isMonthly)
        {
            var start = isMonthly
                ? new DateTime(periodStart.Year, periodStart.Month, 1)
                : periodStart.Date;
            var end = isMonthly
                ? start.AddMonths(1).AddDays(-1)
                : start;
            return new AllocatedExpensePeriodDoubleClickEventArgs
            {
                StartDate = start,
                EndDate = end
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
        public bool IsGroup { get; set; }
        public int TreeLevel { get; set; }
        public double IndentWidth => TreeLevel * 18;
        public string TreeGlyph => IsGroup ? "▾" : "└";
        public FontWeight RowFontWeight => IsGroup ? FontWeights.SemiBold : FontWeights.Normal;
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
                IsGroup = statistics.IsGroup,
                TreeLevel = statistics.TreeLevel,
                NetRmb = statistics.NetRmb,
                NetRmbText = $"{FormatSignedMoney(statistics.NetRmb, "¥")}{missingRateText}",
                CurrentBalanceText = $"{FormatMoney(statistics.CurrentBalanceRmb, "¥")}{missingBalanceRateText}",
                CurrencyDetails = String.Join("、", statistics.CurrentBalanceCurrencyTotals.Select(FormatCurrencyTotal)),
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
        public DateTime Month { get; set; }
        public string MonthLabel { get; set; } = "";
        public decimal Income { get; set; }
        public decimal Expense { get; set; }
        public List<MonthlyFlowSegmentViewModel> IncomeSegments { get; set; } = [];
        public List<MonthlyFlowSegmentViewModel> ExpenseSegments { get; set; } = [];

        public static MonthlyFlowPointViewModel From(MonthlyFlowPoint point)
        {
            return new MonthlyFlowPointViewModel
            {
                Month = point.Month,
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
        const double ChartTopPadding = 44;
        const double ChartBottomPadding = 34;
        const decimal FixedAxisMax = 1000m;

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
        public static readonly DependencyProperty IsMonthlyProperty = DependencyProperty.Register(
            nameof(IsMonthly),
            typeof(bool),
            typeof(AllocatedExpenseChart),
            new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));
        public static readonly DependencyProperty VisibleStartDateProperty = DependencyProperty.Register(
            nameof(VisibleStartDate),
            typeof(DateTime),
            typeof(AllocatedExpenseChart),
            new FrameworkPropertyMetadata(DateTime.Today, FrameworkPropertyMetadataOptions.AffectsRender));
        public static readonly DependencyProperty VisibleEndDateProperty = DependencyProperty.Register(
            nameof(VisibleEndDate),
            typeof(DateTime),
            typeof(AllocatedExpenseChart),
            new FrameworkPropertyMetadata(DateTime.Today, FrameworkPropertyMetadataOptions.AffectsRender));

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
        Point dragCurrentPoint;
        AllocatedExpenseBucketViewModel? hoveredBucket;

        public event EventHandler<AllocatedExpenseWindowShiftEventArgs>? WindowShiftRequested;
        public event EventHandler<AllocatedExpensePeriodDoubleClickEventArgs>? PeriodDoubleClicked;

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

        public bool IsMonthly
        {
            get => (bool)GetValue(IsMonthlyProperty);
            set => SetValue(IsMonthlyProperty, value);
        }

        public DateTime VisibleStartDate
        {
            get => (DateTime)GetValue(VisibleStartDateProperty);
            set => SetValue(VisibleStartDateProperty, value);
        }

        public DateTime VisibleEndDate
        {
            get => (DateTime)GetValue(VisibleEndDateProperty);
            set => SetValue(VisibleEndDateProperty, value);
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
            var position = e.GetPosition(this);
            if (e.ClickCount >= 2)
            {
                if (TryGetBucketAtPoint(position, out var bucket))
                {
                    PeriodDoubleClicked?.Invoke(
                        this,
                        AllocatedExpensePeriodDoubleClickEventArgs.From(bucket.PeriodStart, IsMonthly));
                    e.Handled = true;
                }
                return;
            }

            if (!IsPointInDragArea(position))
                return;

            isDraggingWindow = true;
            dragStartPoint = position;
            dragCurrentPoint = dragStartPoint;
            CaptureMouse();
            e.Handled = true;
        }

        protected override HitTestResult HitTestCore(PointHitTestParameters hitTestParameters)
        {
            return IsPointInInteractiveArea(hitTestParameters.HitPoint)
                ? new PointHitTestResult(this, hitTestParameters.HitPoint)
                : null!;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var position = e.GetPosition(this);
            Cursor = isDraggingWindow || IsPointInDragArea(position)
                ? Cursors.Hand
                : null;
            UpdateHoveredBucket(position);
            if (!isDraggingWindow)
                return;

            dragCurrentPoint = position;
            InvalidateVisual();
            e.Handled = true;
        }

        protected override void OnMouseLeave(MouseEventArgs e)
        {
            base.OnMouseLeave(e);
            if (!isDraggingWindow)
                Cursor = null;
            SetHoveredBucket(null);
        }

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonUp(e);
            if (!isDraggingWindow)
                return;

            var endPoint = e.GetPosition(this);
            var unitOffset = CalculateDragUnitOffset(endPoint);
            isDraggingWindow = false;
            ReleaseMouseCapture();
            InvalidateVisual();
            if (unitOffset != 0)
                WindowShiftRequested?.Invoke(this, new AllocatedExpenseWindowShiftEventArgs { UnitOffset = unitOffset });
            e.Handled = true;
        }

        protected override void OnLostMouseCapture(MouseEventArgs e)
        {
            base.OnLostMouseCapture(e);
            isDraggingWindow = false;
            InvalidateVisual();
        }

        private int CalculateDragUnitOffset(Point endPoint)
        {
            var deltaX = endPoint.X - dragStartPoint.X;
            var visibleUnitCount = GetVisibleUnitCount();
            var chartWidth = Math.Max(1, ActualWidth - ChartLeftPadding - ChartRightPadding);
            var pixelsPerUnit = Math.Max(DragUnitMinimumPixels, chartWidth / visibleUnitCount);
            if (Math.Abs(deltaX) < pixelsPerUnit * 0.5)
                return 0;

            return (int)Math.Round(-deltaX / pixelsPerUnit, MidpointRounding.AwayFromZero);
        }

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);
            var width = ActualWidth;
            var height = ActualHeight;
            var left = ChartLeftPadding;
            var top = ChartTopPadding;
            var right = ChartRightPadding;
            var bottom = ChartBottomPadding;
            var chartWidth = Math.Max(1, width - left - right);
            var chartHeight = Math.Max(1, height - top - bottom);

            var loadedBuckets = Items?.Cast<AllocatedExpenseBucketViewModel>().ToList() ?? [];
            var placements = BuildBucketPlacements(loadedBuckets, left, chartWidth);
            var viewportPlacements = placements
                .Where(placement => IsInViewport(placement, left, chartWidth))
                .ToList();
            if (viewportPlacements.Count == 0 || viewportPlacements.All(placement => placement.Bucket.TotalRmb == 0))
            {
                DrawText(dc, "暂无均摊支出", 14, "#94A3B8", ActualWidth / 2 - 42, ActualHeight / 2 - 10);
                return;
            }

            var axisMax = IsMonthly
                ? CalculateAxisMax(viewportPlacements.Max(placement => placement.Bucket.TotalRmb))
                : FixedAxisMax;
            DrawAxes(dc, left, top, chartWidth, chartHeight, axisMax);
            DrawLegend(dc, viewportPlacements.Select(placement => placement.Bucket).ToList(), width);

            if (IsLineChart)
                DrawLineChart(dc, placements, left, top, chartWidth, chartHeight, axisMax);
            else
                DrawStackedBars(dc, placements, left, top, chartWidth, chartHeight, axisMax);

            DrawLabels(dc, placements, left, top + chartHeight, chartWidth, hoveredBucket);
        }

        private List<AllocatedExpenseBucketPlacement> BuildBucketPlacements(
            List<AllocatedExpenseBucketViewModel> loadedBuckets,
            double left,
            double chartWidth)
        {
            if (loadedBuckets.Count == 0)
                return [];

            var visibleStart = GetVisiblePeriodStart();
            var visibleStartIndex = loadedBuckets.FindIndex(bucket => bucket.PeriodStart.Date == visibleStart);
            if (visibleStartIndex < 0)
                visibleStartIndex = loadedBuckets.FindIndex(bucket => bucket.PeriodStart.Date > visibleStart);
            if (visibleStartIndex < 0)
                visibleStartIndex = loadedBuckets.Count;

            var visibleUnitCount = GetVisibleUnitCount();
            var groupWidth = chartWidth / visibleUnitCount;
            var dragPixelOffset = GetDragPixelOffset();
            return loadedBuckets
                .Select((bucket, index) =>
                {
                    var slot = index - visibleStartIndex;
                    var centerX = left + (slot + 0.5) * groupWidth + dragPixelOffset;
                    return new AllocatedExpenseBucketPlacement(bucket, centerX, groupWidth);
                })
                .Where(placement => placement.CenterX + placement.GroupWidth >= left && placement.CenterX - placement.GroupWidth <= left + chartWidth)
                .ToList();
        }

        private bool TryGetBucketAtPoint(Point point, out AllocatedExpenseBucketViewModel bucket)
        {
            var left = ChartLeftPadding;
            var right = ChartRightPadding;
            var chartWidth = Math.Max(1, ActualWidth - left - right);
            var loadedBuckets = Items?.Cast<AllocatedExpenseBucketViewModel>().ToList() ?? [];
            var placements = BuildBucketPlacements(loadedBuckets, left, chartWidth);
            var placement = placements
                .Where(placement => IsInViewport(placement, left, chartWidth))
                .Where(placement =>
                    point.X >= placement.CenterX - placement.GroupWidth / 2 &&
                    point.X <= placement.CenterX + placement.GroupWidth / 2)
                .OrderBy(placement => Math.Abs(point.X - placement.CenterX))
                .FirstOrDefault();
            if (placement.Bucket is null)
            {
                bucket = new AllocatedExpenseBucketViewModel();
                return false;
            }

            bucket = placement.Bucket;
            return true;
        }

        private void UpdateHoveredBucket(Point point)
        {
            if (!IsPointInDateLabelArea(point))
            {
                SetHoveredBucket(null);
                return;
            }

            SetHoveredBucket(TryGetBucketAtPoint(point, out var bucket) ? bucket : null);
        }

        private void SetHoveredBucket(AllocatedExpenseBucketViewModel? bucket)
        {
            if (ReferenceEquals(hoveredBucket, bucket))
                return;

            hoveredBucket = bucket;
            InvalidateVisual();
        }

        private bool IsPointInInteractiveArea(Point point)
        {
            return IsPointInDragArea(point) || IsPointInDateLabelArea(point);
        }

        private bool IsPointInDragArea(Point point)
        {
            var left = ChartLeftPadding;
            var top = ChartTopPadding;
            var right = Math.Max(left, ActualWidth - ChartRightPadding);
            var bottom = Math.Max(top, ActualHeight - ChartBottomPadding);
            return point.X > left && point.X < right && point.Y > top && point.Y < bottom;
        }

        private bool IsPointInDateLabelArea(Point point)
        {
            var left = ChartLeftPadding;
            var right = Math.Max(left, ActualWidth - ChartRightPadding);
            var baselineY = Math.Max(ChartTopPadding, ActualHeight - ChartBottomPadding);
            return point.X > left && point.X < right && point.Y > baselineY && point.Y < ActualHeight;
        }

        private int GetVisibleUnitCount()
        {
            if (IsMonthly)
            {
                var start = GetMonthStart(VisibleStartDate);
                var endExclusive = GetMonthStart(VisibleEndDate).AddMonths(1);
                return Math.Max(1, (endExclusive.Year - start.Year) * 12 + endExclusive.Month - start.Month);
            }

            return Math.Max(1, (VisibleEndDate.Date - VisibleStartDate.Date).Days + 1);
        }

        private DateTime GetVisiblePeriodStart()
        {
            return IsMonthly ? GetMonthStart(VisibleStartDate) : VisibleStartDate.Date;
        }

        private double GetDragPixelOffset()
        {
            return isDraggingWindow ? dragCurrentPoint.X - dragStartPoint.X : 0;
        }

        private static DateTime GetMonthStart(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        private static bool IsInViewport(AllocatedExpenseBucketPlacement placement, double left, double chartWidth)
        {
            return placement.CenterX + placement.GroupWidth / 2 >= left &&
                placement.CenterX - placement.GroupWidth / 2 <= left + chartWidth;
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
            List<AllocatedExpenseBucketPlacement> placements,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            decimal axisMax)
        {
            if (placements.Count == 0)
                return;

            var groupWidth = placements[0].GroupWidth;
            var barWidth = Math.Min(28, Math.Max(8, groupWidth * 0.42));
            foreach (var placement in placements)
            {
                var x = placement.CenterX - barWidth / 2;
                var y = top + chartHeight;
                foreach (var segment in placement.Bucket.Segments.Where(segment => segment.ChartValue > 0))
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

                if (IsInViewport(placement, left, chartWidth))
                    DrawValueLabel(dc, placement.Bucket.TotalRmb, placement.CenterX, y - 17);
            }
        }

        private static void DrawLineChart(
            DrawingContext dc,
            List<AllocatedExpenseBucketPlacement> placements,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            decimal axisMax)
        {
            var pen = new Pen(ParseBrush("#0F766E"), 2);
            var fill = ParseBrush("#0F766E");
            var points = placements.Select(placement =>
            {
                var y = top + chartHeight - chartHeight * (double)(placement.Bucket.TotalRmb / axisMax);
                return new { Placement = placement, Point = new Point(placement.CenterX, y) };
            }).ToList();

            for (var i = 1; i < points.Count; i++)
                dc.DrawLine(pen, points[i - 1].Point, points[i].Point);

            for (var i = 0; i < points.Count; i++)
            {
                dc.DrawEllipse(fill, null, points[i].Point, 3.5, 3.5);
                if (points[i].Placement.Bucket.TotalRmb > 0 && IsInViewport(points[i].Placement, left, chartWidth))
                    DrawValueLabel(dc, points[i].Placement.Bucket.TotalRmb, points[i].Point.X, points[i].Point.Y - 20);
            }
        }

        private static void DrawLabels(
            DrawingContext dc,
            List<AllocatedExpenseBucketPlacement> placements,
            double left,
            double baselineY,
            double chartWidth,
            AllocatedExpenseBucketViewModel? hoveredBucket)
        {
            var visibleCount = placements.Count == 0
                ? 0
                : Math.Max(1, (int)Math.Round(chartWidth / placements[0].GroupWidth, MidpointRounding.AwayFromZero));
            for (var i = 0; i < placements.Count; i++)
            {
                if (visibleCount > 18 && i % 2 == 1)
                    continue;
                if (!IsInViewport(placements[i], left, chartWidth))
                    continue;

                var x = placements[i].CenterX;
                var isHovered = ReferenceEquals(placements[i].Bucket, hoveredBucket);
                var text = CreateText(
                    placements[i].Bucket.PeriodLabel,
                    isHovered ? 12 : 11,
                    ParseBrush(isHovered ? "#0F766E" : "#64748B"));
                if (isHovered)
                {
                    dc.DrawRoundedRectangle(
                        new SolidColorBrush(Color.FromRgb(204, 251, 241)),
                        null,
                        new Rect(x - text.Width / 2 - 5, baselineY + 9, text.Width + 10, text.Height + 4),
                        4,
                        4);
                }
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

        private static decimal CalculateAxisMax(decimal maxValue)
        {
            if (maxValue <= 0)
                return 1;

            var rawStep = (double)maxValue / 4.0;
            var magnitude = Math.Pow(10, Math.Floor(Math.Log10(Math.Max(1, rawStep))));
            var step = (decimal)(Math.Ceiling(rawStep / magnitude) * magnitude);
            return Math.Max(1, step * 4);
        }

        private readonly record struct AllocatedExpenseBucketPlacement(
            AllocatedExpenseBucketViewModel Bucket,
            double CenterX,
            double GroupWidth);
    }

    public class MonthlyFlowChart : FrameworkElement
    {
        const double DragUnitMinimumPixels = 32;
        const double ChartLeftPadding = 58;
        const double ChartRightPadding = 58;
        const double ChartTopPadding = 46;
        const double ChartBottomPadding = 28;

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

        bool isDraggingWindow;
        Point dragStartPoint;
        Point dragCurrentPoint;
        MonthlyFlowPointViewModel? hoveredPoint;

        public event EventHandler<AllocatedExpenseWindowShiftEventArgs>? WindowShiftRequested;
        public event EventHandler<AllocatedExpensePeriodDoubleClickEventArgs>? PeriodDoubleClicked;

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

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            var position = e.GetPosition(this);
            if (e.ClickCount >= 2)
            {
                if (TryGetPointAtPoint(position, out var point))
                {
                    PeriodDoubleClicked?.Invoke(
                        this,
                        AllocatedExpensePeriodDoubleClickEventArgs.From(point.Month, isMonthly: true));
                    e.Handled = true;
                }
                return;
            }

            if (!IsPointInDragArea(position))
                return;

            isDraggingWindow = true;
            dragStartPoint = position;
            dragCurrentPoint = position;
            CaptureMouse();
            e.Handled = true;
        }

        protected override HitTestResult HitTestCore(PointHitTestParameters hitTestParameters)
        {
            return IsPointInInteractiveArea(hitTestParameters.HitPoint)
                ? new PointHitTestResult(this, hitTestParameters.HitPoint)
                : null!;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var position = e.GetPosition(this);
            Cursor = isDraggingWindow || IsPointInDragArea(position) || IsPointInDateLabelArea(position)
                ? Cursors.Hand
                : null;
            UpdateHoveredPoint(position);
            if (!isDraggingWindow)
                return;

            dragCurrentPoint = position;
            InvalidateVisual();
            e.Handled = true;
        }

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonUp(e);
            if (!isDraggingWindow)
                return;

            var unitOffset = CalculateDragUnitOffset(e.GetPosition(this));
            isDraggingWindow = false;
            ReleaseMouseCapture();
            InvalidateVisual();
            if (unitOffset != 0)
                WindowShiftRequested?.Invoke(this, new AllocatedExpenseWindowShiftEventArgs { UnitOffset = unitOffset });
            e.Handled = true;
        }

        protected override void OnLostMouseCapture(MouseEventArgs e)
        {
            base.OnLostMouseCapture(e);
            isDraggingWindow = false;
            InvalidateVisual();
        }

        protected override void OnMouseLeave(MouseEventArgs e)
        {
            base.OnMouseLeave(e);
            if (!isDraggingWindow)
                Cursor = null;
            SetHoveredPoint(null);
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
            var left = ChartLeftPadding;
            var top = ChartTopPadding;
            var right = ChartRightPadding;
            var bottom = ChartBottomPadding;
            var chartWidth = Math.Max(1, width - left - right);
            var chartHeight = Math.Max(1, height - top - bottom);
            var chartRight = left + chartWidth;
            var useExpense = String.Equals(FlowKind, "Expense", StringComparison.OrdinalIgnoreCase);
            var placements = BuildPointPlacements(points, left, chartWidth);
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
            DrawBars(dc, placements, axisMax, left, top, chartWidth, chartHeight, useExpense);
            DrawLegend(dc, width);
            DrawLabels(dc, placements, left, top + chartHeight, chartWidth, hoveredPoint);
        }

        private static void DrawBars(
            DrawingContext dc,
            List<MonthlyFlowPointPlacement> placements,
            decimal axisMax,
            double left,
            double top,
            double chartWidth,
            double chartHeight,
            bool useExpense)
        {
            if (placements.Count == 0)
                return;

            var groupWidth = placements[0].GroupWidth;
            var barWidth = Math.Min(24, Math.Max(8, groupWidth * 0.28));
            foreach (var placement in placements)
            {
                var segments = GetFlowSegments(placement.Point, useExpense);
                var total = segments.Sum(segment => segment.Value);
                DrawStackedBar(dc, segments, placement.CenterX - barWidth / 2, top, chartHeight, barWidth, axisMax);
                if (IsInViewport(placement, left, chartWidth))
                    DrawBarValueLabel(dc, placement.CenterX, top, chartHeight, total, axisMax, useExpense);
            }
        }

        private List<MonthlyFlowPointPlacement> BuildPointPlacements(
            List<MonthlyFlowPointViewModel> points,
            double left,
            double chartWidth)
        {
            if (points.Count == 0)
                return [];

            var groupWidth = chartWidth / points.Count;
            var dragPixelOffset = GetDragPixelOffset();
            return points
                .Select((point, index) =>
                {
                    var centerX = left + (index + 0.5) * groupWidth + dragPixelOffset;
                    return new MonthlyFlowPointPlacement(point, centerX, groupWidth);
                })
                .Where(placement => placement.CenterX + placement.GroupWidth >= left && placement.CenterX - placement.GroupWidth <= left + chartWidth)
                .ToList();
        }

        private bool TryGetPointAtPoint(Point point, out MonthlyFlowPointViewModel month)
        {
            var left = ChartLeftPadding;
            var chartWidth = Math.Max(1, ActualWidth - ChartLeftPadding - ChartRightPadding);
            var points = Items?.Cast<MonthlyFlowPointViewModel>().ToList() ?? [];
            var placements = BuildPointPlacements(points, left, chartWidth);
            var placement = placements
                .Where(placement => IsInViewport(placement, left, chartWidth))
                .Where(placement =>
                    point.X >= placement.CenterX - placement.GroupWidth / 2 &&
                    point.X <= placement.CenterX + placement.GroupWidth / 2)
                .OrderBy(placement => Math.Abs(point.X - placement.CenterX))
                .FirstOrDefault();
            if (placement.Point is null)
            {
                month = new MonthlyFlowPointViewModel();
                return false;
            }

            month = placement.Point;
            return true;
        }

        private void UpdateHoveredPoint(Point point)
        {
            if (!IsPointInDateLabelArea(point))
            {
                SetHoveredPoint(null);
                return;
            }

            SetHoveredPoint(TryGetPointAtPoint(point, out var month) ? month : null);
        }

        private void SetHoveredPoint(MonthlyFlowPointViewModel? point)
        {
            if (ReferenceEquals(hoveredPoint, point))
                return;

            hoveredPoint = point;
            InvalidateVisual();
        }

        private int CalculateDragUnitOffset(Point endPoint)
        {
            var deltaX = endPoint.X - dragStartPoint.X;
            var pointCount = Math.Max(1, Items?.Cast<MonthlyFlowPointViewModel>().Count() ?? 0);
            var chartWidth = Math.Max(1, ActualWidth - ChartLeftPadding - ChartRightPadding);
            var pixelsPerUnit = Math.Max(DragUnitMinimumPixels, chartWidth / pointCount);
            if (Math.Abs(deltaX) < pixelsPerUnit * 0.5)
                return 0;

            return (int)Math.Round(-deltaX / pixelsPerUnit, MidpointRounding.AwayFromZero);
        }

        private double GetDragPixelOffset()
        {
            return isDraggingWindow ? dragCurrentPoint.X - dragStartPoint.X : 0;
        }

        private static bool IsInViewport(MonthlyFlowPointPlacement placement, double left, double chartWidth)
        {
            return placement.CenterX + placement.GroupWidth / 2 >= left &&
                placement.CenterX - placement.GroupWidth / 2 <= left + chartWidth;
        }

        private bool IsPointInInteractiveArea(Point point)
        {
            return IsPointInDragArea(point) || IsPointInDateLabelArea(point);
        }

        private bool IsPointInDragArea(Point point)
        {
            var left = ChartLeftPadding;
            var top = ChartTopPadding;
            var right = Math.Max(left, ActualWidth - ChartRightPadding);
            var bottom = Math.Max(top, ActualHeight - ChartBottomPadding);
            return point.X > left && point.X < right && point.Y > top && point.Y < bottom;
        }

        private bool IsPointInDateLabelArea(Point point)
        {
            var left = ChartLeftPadding;
            var right = Math.Max(left, ActualWidth - ChartRightPadding);
            var baselineY = Math.Max(ChartTopPadding, ActualHeight - ChartBottomPadding);
            return point.X > left && point.X < right && point.Y > baselineY && point.Y < ActualHeight;
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

        private static void DrawLabels(
            DrawingContext dc,
            List<MonthlyFlowPointPlacement> placements,
            double left,
            double baselineY,
            double chartWidth,
            MonthlyFlowPointViewModel? hoveredPoint)
        {
            foreach (var placement in placements)
            {
                if (!IsInViewport(placement, left, chartWidth))
                    continue;

                var isHovered = ReferenceEquals(placement.Point, hoveredPoint);
                var text = CreateText(
                    placement.Point.MonthLabel,
                    isHovered ? 12 : 11,
                    ParseBrush(isHovered ? "#0F766E" : "#64748B"));
                if (isHovered)
                {
                    dc.DrawRoundedRectangle(
                        new SolidColorBrush(Color.FromRgb(204, 251, 241)),
                        null,
                        new Rect(placement.CenterX - text.Width / 2 - 5, baselineY + 9, text.Width + 10, text.Height + 4),
                        4,
                        4);
                }

                dc.DrawText(text, new Point(placement.CenterX - text.Width / 2, baselineY + 12));
            }
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

        private readonly record struct MonthlyFlowPointPlacement(
            MonthlyFlowPointViewModel Point,
            double CenterX,
            double GroupWidth);
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
