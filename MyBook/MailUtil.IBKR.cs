using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MailKit;
using MimeKit;

namespace MyBook
{
    // IBKR DailyMyBook CSV report discovery, parsing, and account matching.
    partial class MailUtil
    {
        private const StatementImportProvider IBKRProvider = StatementImportProvider.IBKRReportMail;
        private const string IBKRReportSender = "donotreply@interactivebrokers.com";
        private const string IBKRDailyReportType = "DailyMyBook";
        private const string IBKRInitialReportFilePrefix = "IBKR_INITIAL_";
        private const int IBKRMissingReportLimitDays = 14;
        private const string IBKRStockYieldEnhancementLoanSection = "股票收益提升计划股证券出借活动";
        private const string IBKRStockYieldEnhancementLoanSectionWithoutActivity = "股票收益提升计划股证券出借";
        private const string IBKRStockYieldEnhancementCollateralHeldSection = "在IBKRSS持有的股票收益提升计划证券抵押品";
        private const string IBKRStockYieldEnhancementLoanFeeSection = "股票收益提升计划证券出借赚取费用详情";
        private const string IBKRStockYieldEnhancementLoanRateSection = "股票收益提升计划证券出借利率详情";
        private const string IBKRInterestSection = "利息";
        private const string IBKRBondInterestReceivedSection = "收到的债券利息";
        private const string IBKRBondInterestPaidSection = "支付的债券利息";
        private const string IBKRDividendAccrualChangeSection = "应计股息的变化";
        private const string IBKRTransferSection = "转账";
        private const string IBKRDepositWithdrawSection = "存款和取款";
        private const string IBKRDividendSection = "股息";
        private const string IBKRDividendPaymentInLieuSection = "代替股息的支付";
        private const string IBKRWithholdingTaxSection = "代扣税";
        private const string IBKRCreditInterestSection = "贷方利息细节";
        private const string IBKRCommissionAdjustmentSection = "佣金调整";
        private const string IBKRStatementPeriodPnlSection = "账单期间的总损益";

        private static readonly HashSet<string> IBKRAssetGroups = new(StringComparer.Ordinal)
        {
            "股票",
            "债券",
            "外汇"
        };

        private static readonly string[] IBKRCashReportFullHeader = ["货币总结", "货币", "总数", "证券", "期货", "本月截至当前", "本年截至当前", ""];
        private static readonly string[] IBKRCashReportShortHeader = ["货币总结", "货币", "总数", "证券", "期货", ""];
        private static readonly string[] IBKRStockYieldEnhancementLoanFullHeader = ["资产分类", "货币", "代码", "日期", "描述", "", "交易号码", "数量", "抵押品金额"];
        private static readonly string[] IBKRStockYieldEnhancementLoanShortHeader = ["资产分类", "货币", "代码", "交易号码", "数量", "客户抵押品的股票收益提升计划利率 (%)", "抵押品金额"];
        private static readonly string[] IBKRStockYieldEnhancementCollateralHeldHeader = ["资产分类", "货币", "代码", "数量", "价格", "价值"];

        private static readonly HashSet<string> IBKRCsvSections = new(StringComparer.Ordinal)
        {
            "Statement",
            "账户信息",
            "净资产值",
            "净资产值变更",
            "按市值计算的表现总结",
            "持仓与以市值计的盈亏",
            "现金报告",
            "净股票持仓总结",
            "按资产类型显示的交易总结",
            "按代码显示的交易总结",
            "佣金细节",
            "应计利息",
            "未平仓应计股息",
            IBKRDividendAccrualChangeSection,
            "借方利息细节",
            "基础货币汇率",
            "代码",
            IBKRTransferSection,
            IBKRStockYieldEnhancementLoanSection,
            IBKRStockYieldEnhancementCollateralHeldSection,
            IBKRStockYieldEnhancementLoanFeeSection,
            IBKRStockYieldEnhancementLoanRateSection,
            IBKRInterestSection,
            IBKRBondInterestReceivedSection,
            IBKRBondInterestPaidSection,
            IBKRDepositWithdrawSection,
            IBKRDividendSection,
            IBKRDividendPaymentInLieuSection,
            IBKRWithholdingTaxSection,
            IBKRCreditInterestSection,
            IBKRCommissionAdjustmentSection,
            IBKRStatementPeriodPnlSection
        };

        private static readonly Dictionary<string, string[][]> IBKRCsvHeaders = new(StringComparer.Ordinal)
        {
            ["Statement"] =
            [
                ["域名称", "域值"]
            ],
            ["账户信息"] =
            [
                ["域名称", "域值"]
            ],
            ["净资产值"] =
            [
                ["资产类型", "之前合计", "当前多头", "当前空头", "当前合计", "变更"],
                ["时间加权的收益率"]
            ],
            ["净资产值变更"] =
            [
                ["域名称", "域值"]
            ],
            ["按市值计算的表现总结"] =
            [
                ["资产分类", "代码", "先前 数量", "当前 数量", "先前 价格", "当前 价格", "按市值计盈亏 持仓", "按市值计盈亏 交易", "按市值计盈亏 佣金", "按市值计盈亏 其它", "按市值计盈亏 总数", "代码"]
            ],
            ["持仓与以市值计的盈亏"] =
            [
                ["DataDiscriminator", "资产类型", "货币", "代码", "描述", "之前数量", "数量", "之前价格", "价格", "之前的市场价值", "市场价值", "持仓", "交易", "佣金", "其它", "总数"]
            ],
            ["现金报告"] =
            [
                ["货币总结", "货币", "总数", "证券", "期货", "本月截至当前", "本年截至当前", ""]
            ],
            ["净股票持仓总结"] =
            [
                ["资产分类", "货币", "代码", "描述", "在IB的股份", "借入的股份", "借出的股份", "净股份"]
            ],
            ["按资产类型显示的交易总结"] =
            [
                ["资产分类", "交易总次数", "购入股份（或合约）的总数", "售出股份（或合约）的总数", "总佣金", "佣金费率"]
            ],
            ["按代码显示的交易总结"] =
            [
                ["资产分类", "货币", "代码", "买入 数量", "买入 平均价", "买入 收益", "卖出 数量", "卖出 平均价", "卖出 收益"]
            ],
            ["佣金细节"] =
            [
                ["资产分类", "货币", "代码", "日期/时间", "数量", "佣金", "经纪商收费 执行", "经纪商收费 清算", "第三方收费 执行", "第三方收费 清算", "第三方收费 交易费", "其他"]
            ],
            ["应计利息"] =
            [
                ["货币", "域名称", "域值"]
            ],
            ["未平仓应计股息"] =
            [
                ["资产分类", "货币", "代码", "除息日", "支付日期", "数量", "税", "费用", "总价", "总额", "净额", "代码"]
            ],
            [IBKRDividendAccrualChangeSection] =
            [
                ["资产分类", "货币", "代码", "日期", "除息日", "支付日期", "数量", "税", "费用", "总股息率", "总额", "净额", "代码"]
            ],
            ["借方利息细节"] =
            [
                ["货币", "起息日", "等级差", "率（％）", "证券本金", "期货本金", "合计本金", "证券利息", "期货利息", "总利息", "代码"]
            ],
            ["基础货币汇率"] =
            [
                ["货币", "汇率"]
            ],
            ["代码"] =
            [
                ["代码", "意思", "代码 （继续）", "意思 （继续）"]
            ],
            [IBKRTransferSection] =
            [
                ["资产分类", "货币", "代码", "日期", "类型", "方向", "转账公司", "转账账户", "数量", "转账价格", "市场价值", "已实现的损益", "现金金额", "代码"]
            ],
            [IBKRStockYieldEnhancementLoanSection] =
            [
                IBKRStockYieldEnhancementLoanFullHeader
            ],
            [IBKRStockYieldEnhancementCollateralHeldSection] =
            [
                IBKRStockYieldEnhancementLoanFullHeader
            ],
            [IBKRStockYieldEnhancementLoanFeeSection] =
            [
                ["货币", "起息日", "代码", "开始日期", "数量", "抵押品金额", "基于市场的利率（%）", "股票收益提升计划利率- 客户抵押品（%）", "股票收益提升计划费用 客户赚取的", "代码"]
            ],
            [IBKRStockYieldEnhancementLoanRateSection] =
            [
                ["货币", "起息日", "代码", "开始日期", "数量", "抵押品金额", "基于市场的利率（%）", "利息- 客户抵押品（%）", "利息 支付给客户", "代码"]
            ],
            [IBKRInterestSection] =
            [
                ["货币", "日期", "描述", "金额"]
            ],
            [IBKRBondInterestReceivedSection] =
            [
                ["货币", "日期", "描述", "金额", "代码"]
            ],
            [IBKRBondInterestPaidSection] =
            [
                ["货币", "日期", "描述", "金额", "代码"]
            ],
            [IBKRDepositWithdrawSection] =
            [
                ["货币", "结算日期", "描述", "金额"]
            ],
            [IBKRDividendSection] =
            [
                ["货币", "日期", "描述", "金额"]
            ],
            [IBKRDividendPaymentInLieuSection] =
            [
                ["货币", "日期", "描述", "金额", "代码"]
            ],
            [IBKRWithholdingTaxSection] =
            [
                ["货币", "日期", "描述", "金额", "代码"]
            ],
            [IBKRCreditInterestSection] =
            [
                ["货币", "起息日", "等级差", "率（％）", "证券本金", "期货本金", "合计本金", "证券利息", "期货利息", "总利息", "代码"]
            ],
            [IBKRCommissionAdjustmentSection] =
            [
                ["货币", "日期", "描述", "金额", "代码"]
            ],
            [IBKRStatementPeriodPnlSection] =
            [
            ]
        };

        private static readonly Dictionary<string, string[][]> IBKRFlexibleCsvHeaders = new(StringComparer.Ordinal)
        {
            [IBKRDepositWithdrawSection] =
            [
                ["货币", "结算日期", "描述", "金额"],
                ["货币", "结算日期", "描述", "金额", "代码"]
            ],
            [IBKRDividendSection] =
            [
                ["货币", "日期", "描述", "金额"],
                ["货币", "日期", "描述", "金额", "代码"]
            ],
            ["佣金细节"] =
            [
                ["资产分类", "货币", "代码", "日期/时间", "数量", "佣金", "经纪商收费 执行", "经纪商收费 清算", "第三方收费 执行", "第三方收费 清算", "第三方收费 交易费", "其他"],
                ["资产分类", "货币", "代码", "日期/时间", "数量", "佣金 in USD", "经纪商收费 执行 in USD", "经纪商收费 清算 in USD", "第三方收费 执行 in USD", "第三方收费 清算 in USD", "第三方收费 交易费 in USD", "其他 in USD"]
            ],
            [IBKRStockYieldEnhancementLoanSection] =
            [
                IBKRStockYieldEnhancementLoanFullHeader,
                IBKRStockYieldEnhancementLoanShortHeader
            ],
            [IBKRStockYieldEnhancementCollateralHeldSection] =
            [
                IBKRStockYieldEnhancementLoanFullHeader,
                IBKRStockYieldEnhancementLoanShortHeader,
                IBKRStockYieldEnhancementCollateralHeldHeader
            ]
        };

        private static readonly Dictionary<string, HoldingType> IBKRFallbackHoldingTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            ["DIA"] = HoldingType.ARCA,
            ["NVDL"] = HoldingType.NASDAQ,
            ["NVDA"] = HoldingType.NASDAQ,
            ["QQQ"] = HoldingType.NASDAQ,
            ["QQQI"] = HoldingType.NASDAQ,
            ["SPY"] = HoldingType.ARCA,
            ["TLT"] = HoldingType.NASDAQ,
            ["TQQQ"] = HoldingType.NASDAQ,
        };

        private static string ReadIBKRLocalCsv(string path)
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }

        public async Task FetchIBKRReports()
        {
            await RunWithMailSessionScope(async () =>
            {
                var date = GetNextDailyStatementDate(IBKRProvider);
                Console.WriteLine($"Fetch IBKR reports from {date:yyyy-MM-dd}");
                var missingDays = 0;
                while (date <= DateTime.Today)
                {
                    Console.WriteLine($"Fetch IBKR report {date:yyyy-MM-dd}");
                    var imported = await FetchIBKRReport(date).ConfigureAwait(false);
                    if (imported)
                    {
                        missingDays = 0;
                    }
                    else
                    {
                        missingDays++;
                        if (missingDays >= IBKRMissingReportLimitDays)
                            throw new InvalidOperationException($"Missing IBKR reports for {IBKRMissingReportLimitDays} consecutive days ending {date:yyyy-MM-dd}");
                    }

                    date = date.AddDays(1);
                }
            }).ConfigureAwait(false);
        }

        private async Task<bool> FetchIBKRReport(DateTime date)
        {
            var subjectDate = date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            var expectedSubject = $"{subjectDate}的自定义活动报表";
            var messages = await SearchBills(
                IBKRReportSender,
                expectedSubject,
                date,
                summary => SummarySubjectEquals(summary, expectedSubject)
                    && SummaryIsFrom(summary, IBKRReportSender)
                    && HasIBKRReportAttachment(summary, date),
                message => String.Equals(message.Subject?.Trim(), expectedSubject, StringComparison.Ordinal)
                    && HasIBKRReportAttachment(message, date));

            if (messages.Count == 0)
            {
                Console.WriteLine($"Find no IBKR {IBKRDailyReportType} report {subjectDate}");
                return false;
            }

            var reports = new List<IBKRParsedReport>();
            foreach (var message in messages)
            {
                var reportAttachments = ReadIBKRReportAttachments(message, date);
                if (reportAttachments.Count == 0)
                {
                    Console.WriteLine($"parse IBKR report mail fail: no supported csv attachment, subject={message.Subject}");
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Supported CSV Attachment: {message.Subject}");
                }

                var mailDate = GetMailDate(message);
                foreach (var attachment in reportAttachments)
                {
                    if (attachment.ReportDate.Date != date.Date)
                        throw new MailParseException($"Parse IBKR Report Fail, Date Mismatch: expected {date:yyyy-MM-dd}, got {attachment.ReportDate:yyyy-MM-dd}");

                    Console.WriteLine($"Load IBKR csv report in memory: {attachment.FileName}, id={attachment.ReportId}, bytes={attachment.Content.Length}");
                    var csv = Encoding.UTF8.GetString(attachment.Content);
                    reports.Add(ParseIBKRReportCsv(csv, attachment.ReportDate, attachment.FileName, mailDate));
                }
            }

            SaveIBKRParsedReports(reports);
            return true;
        }

        private List<bool> SaveIBKRParsedReports(List<IBKRParsedReport> reports)
        {
            var saveItems = AddIBKRInitialReportsIfNeeded(reports);
            var saved = database.SaveStatementRecordsAndHoldingsOnce(saveItems.Select(item =>
            {
                var report = item.Report;
                return
                new StatementRecordHoldingImport(
                    IBKRProvider,
                    report.ImportDate,
                    report.StatementKey,
                    report.Account,
                    report.Records,
                    report.Holdings,
                    report.AccountBalances,
                    report.BeginningAccountBalances,
                    report.BeginningHoldings,
                    report.InternalCardNos);
            }));

            for (var i = 0; i < saveItems.Count; i++)
            {
                var report = saveItems[i].Report;
                var periodText = report.Period.StartDate == report.Period.EndDate
                    ? $"{report.ReportDate:yyyy-MM-dd}"
                    : $"{report.Period.StartDate:yyyy-MM-dd}..{report.Period.EndDate:yyyy-MM-dd}";
                var reportType = report.IsInitialReport ? "IBKR initial report" : "IBKR report";
                Console.WriteLine(
                    saved[i]
                        ? $"Import {reportType}: {periodText} {report.AccountId} -> {report.Account.name}, records={report.Records.Count}, holdings={report.Holdings.Count}"
                        : $"Skip imported {reportType}: {periodText} {report.AccountId} -> {report.Account.name}");
            }

            return saveItems
                .Zip(saved, (item, isSaved) => (item, isSaved))
                .Where(item => item.item.ReturnResult)
                .Select(item => item.isSaved)
                .ToList();
        }

        private List<IBKRReportSaveItem> AddIBKRInitialReportsIfNeeded(List<IBKRParsedReport> reports)
        {
            var result = new List<IBKRReportSaveItem>();
            var checkedAccountIds = new HashSet<int>();
            foreach (var report in reports.OrderBy(report => report.ReportDate).ThenBy(report => report.Account.name, StringComparer.Ordinal))
            {
                if (!report.IsInitialReport
                    && checkedAccountIds.Add(report.Account.Id)
                    && !database.HasAccountHistory(report.Account))
                {
                    foreach (var initialReport in LoadIBKRInitialReports(report.Account))
                        result.Add(new IBKRReportSaveItem(initialReport, false));
                }

                result.Add(new IBKRReportSaveItem(report, true));
            }

            return result;
        }

        private List<IBKRParsedReport> LoadIBKRInitialReports(Account account)
        {
            var directory = FindInitialReportsDirectory();
            if (directory is null)
                return [];

            var parsedReports = Directory.GetFiles(directory, $"{IBKRInitialReportFilePrefix}*.csv")
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
                .Select(ParseIBKRInitialReportCsv)
                .Where(report => report.Account.Id == account.Id)
                .OrderBy(report => report.Period.StartDate)
                .ThenBy(report => report.Period.EndDate)
                .ThenBy(report => report.StatementKey, StringComparer.Ordinal)
                .ToList();
            if (parsedReports.Count == 0)
                return [];

            ValidateIBKRInitialReportChain(account, parsedReports);
            return parsedReports;
        }

        private static void ValidateIBKRInitialReportChain(Account account, List<IBKRParsedReport> reports)
        {
            if (reports.Count == 0)
                return;

            ValidateIBKRZeroInitialReportStart(account, reports[0]);
            for (var i = 1; i < reports.Count; i++)
            {
                var previous = reports[i - 1];
                var current = reports[i];
                AssertIBKRBalanceSnapshotsEqual(
                    previous.AccountBalances,
                    current.BeginningAccountBalances,
                    $"IBKR initial balance chain {account.name} {previous.Period.EndDate:yyyy-MM-dd}->{current.Period.StartDate:yyyy-MM-dd}");
                AssertIBKRHoldingSnapshotsEqual(
                    previous.Holdings,
                    current.BeginningHoldings,
                    $"IBKR initial holding chain {account.name} {previous.Period.EndDate:yyyy-MM-dd}->{current.Period.StartDate:yyyy-MM-dd}");
            }
        }

        private static void ValidateIBKRZeroInitialReportStart(Account account, IBKRParsedReport report)
        {
            foreach (var balance in report.BeginningAccountBalances)
            {
                if (balance.v != 0)
                    throw new MailParseException($"Parse IBKR Report Fail, Initial Report Nonzero Beginning Balance: {account.name}/{balance.v} {balance.t}");
            }

            foreach (var holding in report.BeginningHoldings)
            {
                var hasQuantity = !Holding.IsSingleValueAsset(holding.holdingType) && holding.quantity != 0;
                if (hasQuantity || holding.totalPrice.v != 0)
                {
                    throw new MailParseException(
                        $"Parse IBKR Report Fail, Initial Report Nonzero Beginning Holding: {account.name}/{holding.code}/{holding.holdingType}/{holding.quantity}/{holding.totalPrice.v} {holding.totalPrice.t}");
                }
            }
        }

        private static void AssertIBKRBalanceSnapshotsEqual(
            List<AccountBalance> expected,
            List<AccountBalance> actual,
            string context)
        {
            var expectedValues = expected.ToDictionary(balance => balance.t, balance => balance.v);
            var actualValues = actual.ToDictionary(balance => balance.t, balance => balance.v);
            foreach (var currency in expectedValues.Keys.Concat(actualValues.Keys).Distinct())
            {
                expectedValues.TryGetValue(currency, out var expectedValue);
                actualValues.TryGetValue(currency, out var actualValue);
                if (expectedValue != actualValue)
                    throw new MailParseException($"Parse IBKR Report Fail, {context} mismatch: {currency} expected {expectedValue}, got {actualValue}");
            }
        }

        private static void AssertIBKRHoldingSnapshotsEqual(
            List<Holding> expected,
            List<Holding> actual,
            string context)
        {
            var expectedValues = expected.ToDictionary(GetIBKRHoldingKey, CreateIBKRHoldingSnapshotValue, StringComparer.OrdinalIgnoreCase);
            var actualValues = actual.ToDictionary(GetIBKRHoldingKey, CreateIBKRHoldingSnapshotValue, StringComparer.OrdinalIgnoreCase);
            foreach (var key in expectedValues.Keys.Concat(actualValues.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                expectedValues.TryGetValue(key, out var expectedValue);
                actualValues.TryGetValue(key, out var actualValue);
                if (expectedValue != actualValue)
                    throw new MailParseException($"Parse IBKR Report Fail, {context} mismatch: {key} expected {expectedValue}, got {actualValue}");
            }
        }

        private static IBKRHoldingSnapshotValue CreateIBKRHoldingSnapshotValue(Holding holding)
        {
            return new IBKRHoldingSnapshotValue(holding.quantity, holding.totalPrice);
        }

        private IBKRParsedReport ParseIBKRInitialReportCsv(string path)
        {
            var csv = ReadIBKRLocalCsv(path);
            if (String.IsNullOrWhiteSpace(csv))
                throw new MailParseException($"Parse IBKR Report Fail, Empty CSV: {path}");

            var report = ReadIBKRCsvReport(csv, path);
            var period = ValidateIBKRStatementMetadata(report, null, allowPeriodRange: true);
            var sourceName = $"{path}; initial={period.StartDate:yyyy-MM-dd}..{period.EndDate:yyyy-MM-dd}";
            return ParseIBKRReportCsvCore(report, period, period.EndDate, period.EndDate, sourceName, isInitialReport: true);
        }

        private IBKRParsedReport ParseIBKRReportCsv(string csv, DateTime reportDate, string sourceName, DateTime? importDate = null)
        {
            if (String.IsNullOrWhiteSpace(csv))
                throw new MailParseException("Parse IBKR Report Fail, Empty CSV");

            var report = ReadIBKRCsvReport(csv, sourceName);
            var period = ValidateIBKRStatementMetadata(report, reportDate, allowPeriodRange: false);
            return ParseIBKRReportCsvCore(report, period, reportDate.Date, (importDate ?? reportDate).Date, sourceName, isInitialReport: false);
        }

        private IBKRParsedReport ParseIBKRReportCsvCore(
            IBKRCsvReport report,
            IBKRStatementPeriod period,
            DateTime reportDate,
            DateTime importDate,
            string sourceName,
            bool isInitialReport)
        {
            var (accountId, baseCurrency) = ParseIBKRAccountInfo(report);
            var account = GetIBKRAccount(accountId);
            var contractInfos = BuildIBKRContractInfos(report);
            _ = ParseIBKRNavChange(report);
            var preciseNav = ParseIBKRPreciseNavValues(report, baseCurrency);
            var positionHoldings = ParseIBKRPositionAndMtmHoldings(report, account, contractInfos, out var beginningPositionHoldings);
            var beginningHoldings = ParseIBKRHoldings(
                report,
                account,
                baseCurrency,
                beginningPositionHoldings,
                sourceName,
                preciseNav.StartingCash,
                useBeginningValues: true);
            var holdings = ParseIBKRHoldings(report, account, baseCurrency, positionHoldings, sourceName, preciseNav.EndingCash);
            var records = ParseIBKRRecords(
                report,
                account,
                baseCurrency,
                contractInfos,
                reportDate.Date,
                sourceName,
                preciseNav.End - preciseNav.Start,
                preciseNav.StartingCash,
                preciseNav.EndingCash,
                isInitialReport);
            var internalCardNos = new List<AccountInternalId>
            {
                new()
                {
                    Account = account,
                    cardNo = accountId,
                    desc = "IBKR report account id",
                    currencyType = baseCurrency,
                    sourceText = $"IBKR report {reportDate:yyyy-MM-dd}; source={sourceName}; accountId={accountId}; account={account.name}"
                }
            };

            return new IBKRParsedReport(
                reportDate.Date,
                importDate.Date,
                isInitialReport
                    ? BuildIBKRInitialStatementKey(account, period)
                    : BuildIBKRStatementKey(account, reportDate),
                accountId,
                account,
                records,
                beginningHoldings,
                holdings,
                [new AccountBalance(account, new Currency(preciseNav.End, baseCurrency))],
                [new AccountBalance(account, new Currency(preciseNav.Start, baseCurrency))],
                internalCardNos,
                period,
                isInitialReport);
        }

        private static string BuildIBKRStatementKey(Account account, DateTime reportDate)
        {
            const string prefix = "IBKR_";
            var accountId = account.name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? account.name[prefix.Length..]
                : account.name;
            return $"{accountId}_{reportDate:yyyy-MM-dd}";
        }

        private static string BuildIBKRInitialStatementKey(Account account, IBKRStatementPeriod period)
        {
            const string prefix = "IBKR_";
            var accountId = account.name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? account.name[prefix.Length..]
                : account.name;
            return $"{accountId}_{period.StartDate:yyyy-MM-dd}_{period.EndDate:yyyy-MM-dd}_initial";
        }

        private Account GetIBKRAccount(string reportAccountId)
        {
            var normalizedReportAccountId = reportAccountId.Trim().ToUpperInvariant();
            var accountMatch = Regex.Match(normalizedReportAccountId, @"U(?:\*+|\d*)\d{4,}", RegexOptions.IgnoreCase);
            if (accountMatch.Success)
                normalizedReportAccountId = accountMatch.Value.ToUpperInvariant();
            var visibleDigits = Regex.Match(normalizedReportAccountId, @"\d+$").Value;
            if (String.IsNullOrWhiteSpace(visibleDigits))
                throw new InvalidOperationException($"Invalid IBKR account id: {reportAccountId}");

            var candidates = database.GetAllAccounts()
                .Where(account => account.name.StartsWith("IBKR_", StringComparison.OrdinalIgnoreCase))
                .Where(account =>
                {
                    var accountId = account.name["IBKR_".Length..].Trim().ToUpperInvariant();
                    if (normalizedReportAccountId.Contains('*'))
                        return accountId.EndsWith(visibleDigits, StringComparison.OrdinalIgnoreCase);
                    return accountId.Equals(normalizedReportAccountId, StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            if (candidates.Count == 1)
                return candidates[0];

            if (candidates.Count > 1)
            {
                throw new InvalidOperationException(
                    $"Find multiple IBKR accounts for {reportAccountId}: {String.Join(",", candidates.Select(account => account.name))}");
            }

            throw new InvalidOperationException($"Account not found: IBKR/{reportAccountId}");
        }

        private static IBKRStatementPeriod ValidateIBKRStatementMetadata(
            IBKRCsvReport report,
            DateTime? reportDate,
            bool allowPeriodRange)
        {
            var knownFields = new HashSet<string>(StringComparer.Ordinal)
            {
                "BrokerName",
                "BrokerAddress",
                "Title",
                "Period",
                "WhenGenerated"
            };
            var values = ReadIBKRKeyValueRows(report, "Statement", knownFields);
            if (!values.TryGetValue("Title", out var title) || title != "活动账单")
                throw new MailParseException($"Parse IBKR Report Fail, Invalid Statement Title: {title}");
            if (!values.TryGetValue("Period", out var period))
                throw new MailParseException("Parse IBKR Report Fail, Missing Statement Period");
            var statementPeriod = ParseIBKRStatementPeriod(period);
            if (!allowPeriodRange && statementPeriod.StartDate.Date != statementPeriod.EndDate.Date)
                throw new MailParseException($"Parse IBKR Report Fail, Unexpected Statement Period Range: {period}");
            if (reportDate.HasValue && statementPeriod.EndDate.Date != reportDate.Value.Date)
                throw new MailParseException($"Parse IBKR Report Fail, Statement Period Mismatch: {period}/{reportDate.Value:yyyy-MM-dd}");

            return statementPeriod;
        }

        private static (string AccountId, CurrencyType BaseCurrency) ParseIBKRAccountInfo(IBKRCsvReport report)
        {
            var knownFields = new HashSet<string>(StringComparer.Ordinal)
            {
                "名称",
                "账户",
                "账户类型",
                "客户类型",
                "账户能力",
                "基础货币"
            };
            var values = ReadIBKRKeyValueRows(report, "账户信息", knownFields);

            if (!values.TryGetValue("账户", out var accountId) || String.IsNullOrWhiteSpace(accountId))
                throw new MailParseException("Parse IBKR Report Fail, Missing Account");
            if (!values.TryGetValue("基础货币", out var baseCurrency) || String.IsNullOrWhiteSpace(baseCurrency))
                throw new MailParseException("Parse IBKR Report Fail, Missing Base Currency");

            return (ParseIBKRAccountId(accountId) ?? accountId, ParseIBKRCurrencyType(baseCurrency));
        }

        private static Dictionary<string, string> ReadIBKRKeyValueRows(
            IBKRCsvReport report,
            string sectionName,
            HashSet<string> knownFields)
        {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var row in report.RequireDataRows(sectionName))
            {
                AssertIBKRFieldCount(row, 2);
                var key = row.Fields[0];
                if (!knownFields.Contains(key))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown {sectionName} Field: {FormatIBKRCsvRow(row)}");
                values[key] = row.Fields[1];
            }

            return values;
        }

        private static Dictionary<string, IBKRContractInfo> BuildIBKRContractInfos(IBKRCsvReport report)
        {
            var contracts = new Dictionary<string, IBKRContractInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in report.OptionalDataRows("持仓与以市值计的盈亏"))
            {
                if (!IsIBKRPositionSummaryRow(row))
                    continue;

                var group = row.Fields[1];
                if (!IsIBKRStockOrBondGroup(group))
                    continue;

                var code = row.Fields[3];
                var description = row.Fields[4];
                AddIBKRContractInfo(contracts, CreateIBKRContractInfo(group, code, description));
            }

            foreach (var row in report.OptionalDataRows("净股票持仓总结"))
            {
                AssertIBKRFieldCount(row, 8);
                if (row.Fields[0] != "股票")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Stock Position Group: {FormatIBKRCsvRow(row)}");

                AddIBKRContractInfo(contracts, CreateIBKRContractInfo(row.Fields[0], row.Fields[2], row.Fields[3]));
            }

            foreach (var row in report.OptionalDataRows("按市值计算的表现总结"))
            {
                AssertIBKRFieldCount(row, 12);
                if (!IBKRAssetGroups.Contains(row.Fields[0]))
                    continue;
                if (row.Fields[0] == "外汇")
                    continue;
                AddIBKRContractInfo(contracts, CreateIBKRContractInfo(row.Fields[0], row.Fields[1], row.Fields[1]));
            }

            return contracts;
        }

        private static void AddIBKRContractInfo(Dictionary<string, IBKRContractInfo> contracts, IBKRContractInfo contract)
        {
            if (String.IsNullOrWhiteSpace(contract.Code))
                return;

            contracts[contract.Code] = contract;
        }

        private static IBKRContractInfo CreateIBKRContractInfo(string group, string rawCode, string description)
        {
            var code = rawCode.Trim();
            if (group == "股票")
            {
                if (!IBKRFallbackHoldingTypes.TryGetValue(code, out var holdingType))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Stock Exchange: {code}/{description}");

                return new IBKRContractInfo(code, description, holdingType, code);
            }

            if (group == "债券")
            {
                var normalizedBondCode = ExtractIBKRBondCode(String.IsNullOrWhiteSpace(code) ? description : code);
                return new IBKRContractInfo(
                    normalizedBondCode,
                    String.IsNullOrWhiteSpace(description) ? normalizedBondCode : description,
                    HoldingType.UST,
                    NormalizeIBKRBondDisplayText(normalizedBondCode, description));
            }

            if (group == "外汇" && TryParseIBKRCurrencyType(code, out _))
                return new IBKRContractInfo(code, code, HoldingType.Cash, code);

            throw new MailParseException($"Parse IBKR Report Fail, Unknown Contract Group: {group}/{rawCode}");
        }

        private Records ParseIBKRRecords(
            IBKRCsvReport report,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            DateTime reportDate,
            string sourceName,
            decimal expectedNavChange,
            decimal preciseStartingCash,
            decimal preciseEndingCash,
            bool isInitialReport)
        {
            var builder = new IBKRRecordBuilder(account, reportDate, sourceName, baseCurrency, isInitialReport);
            var tradeTotals = ParseIBKRTradeSummary(report, contractInfos);
            var quantityChangeTotals = ParseIBKRPositionQuantityChangeRecords(report, builder, baseCurrency, contractInfos, tradeTotals);
            var mtmTotals = ParseIBKRMtmRecords(report, builder, baseCurrency, contractInfos, quantityChangeTotals.CoveredTransactionImpacts);
            var commissionTotals = ParseIBKRCommissionRecords(report, builder, baseCurrency, contractInfos);
            var interestAccrualTotals = ParseIBKRInterestAccrualRecords(report, builder, baseCurrency);
            var debitInterestTotals = ParseIBKRDebitInterestDetails(report);
            var dividendAccrualTotal = ParseIBKRDividendAccruals(report);
            var dividendAccrualChangeTotal = ParseIBKRDividendAccrualChangeRecords(report, builder, baseCurrency, contractInfos);
            var transferTotal = ParseIBKRTransferRecords(report, builder, baseCurrency, account, contractInfos);
            var cashTotals = ParseIBKRCashRecords(report, builder, baseCurrency, tradeTotals, commissionTotals);
            AddIBKRPreciseFxTransactionRecord(
                builder,
                baseCurrency,
                preciseStartingCash,
                preciseEndingCash,
                tradeTotals,
                mtmTotals,
                cashTotals);

            if (mtmTotals.HasData)
            {
                if (commissionTotals.HasData)
                {
                    AssertIBKRMoneyDisplayAlmostEquals(commissionTotals.Total, ParseIBKRNavChangeComponent(report, "佣金"), "IBKR NAV change commission display");
                    AssertIBKRMoneyEquals(mtmTotals.Commission, cashTotals.Commission, "IBKR MTM cash commission");
                }
                else
                    AssertIBKRMoneyEquals(mtmTotals.Commission, cashTotals.Commission, "IBKR cash commission");

                var navOther =
                    ParseIBKRNavChangeComponent(report, "股息")
                    + ParseIBKRNavChangeComponent(report, "代扣税款")
                    + ParseIBKRNavChangeComponent(report, "利息");
                AssertIBKRMoneyEquals(mtmTotals.Other, navOther, "IBKR MTM other");
            }

            var accruedInterestRow = FindIBKRNavComponentRow(report, "应计利息");
            var accruedInterest = accruedInterestRow is null
                ? 0
                : ParseIBKRDecimalAt(accruedInterestRow, 5, "NAV component change 应计利息");
            if (interestAccrualTotals.HasData)
            {
                if (accruedInterestRow is null)
                    throw new MailParseException("Parse IBKR Report Fail, Missing Interest Accrual NAV Row");
                AssertIBKRMoneyFieldAlmostEquals(interestAccrualTotals.Total, accruedInterest, accruedInterestRow.Fields[5], "IBKR interest accrual NAV");
            }

            _ = debitInterestTotals;

            var accruedDividend = ParseIBKRNavComponent(report, "应计股息");
            if (dividendAccrualTotal.HasData)
                AssertIBKRMoneyEquals(accruedDividend, dividendAccrualTotal.Total, "IBKR open dividend accrual");

            var accruedDividendRow = FindIBKRNavComponentRow(report, "应计股息");
            var accruedDividendChange = accruedDividendRow is null
                ? 0
                : ParseIBKRDecimalAt(accruedDividendRow, 5, "NAV component change 应计股息");
            if (dividendAccrualChangeTotal.HasData)
            {
                if (accruedDividendRow is null)
                    throw new MailParseException("Parse IBKR Report Fail, Missing Dividend Accrual NAV Row");
                AssertIBKRMoneyFieldAlmostEquals(dividendAccrualChangeTotal.Total, accruedDividendChange, accruedDividendRow.Fields[5], "IBKR dividend accrual change");
            }

            var positionTransfer = ParseIBKRNavChangeComponent(report, "持仓转账");
            if (transferTotal.HasData)
                AssertIBKRMoneyEquals(positionTransfer, transferTotal.Total, "IBKR position transfer");

            AddIBKROtherFxTranslationRecord(report, builder, baseCurrency);
            if (expectedNavChange != builder.NetAssetChangeTotal)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, IBKR NAV change records mismatch: expected {expectedNavChange}, got {builder.NetAssetChangeTotal}; records={builder.DescribeNetAssetChangeTotals()}");
            }

            return builder.Records;
        }

        private static void AddIBKRPreciseFxTransactionRecord(
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            decimal preciseStartingCash,
            decimal preciseEndingCash,
            IBKRTransactionTotals tradeTotals,
            IBKRMtmTotals mtmTotals,
            IBKRCashTotals cashTotals)
        {
            if (!tradeTotals.HasFxTrades && mtmTotals.FxTransaction == 0)
                return;
            if (!mtmTotals.HasData)
                throw new MailParseException("Parse IBKR Report Fail, FX trades without MTM data");

            AssertIBKRMoneyWithin(
                mtmTotals.CashFxTranslation,
                cashTotals.CashFxTranslation,
                0.0000001m,
                "IBKR cash FX translation display");

            var cashReportFxTransaction = cashTotals.TradeBuy + cashTotals.TradeSell - tradeTotals.StockBondProceeds;
            var preciseFxTransaction =
                preciseEndingCash
                - preciseStartingCash
                - cashTotals.NonTradeCashFlow
                - tradeTotals.StockBondProceeds
                - mtmTotals.CashFxTranslation;

            AssertIBKRMoneyWithin(
                preciseFxTransaction,
                cashReportFxTransaction,
                0.0000001m,
                "IBKR cash report implied FX transaction");

            var displayTolerance = Math.Max(0.00001m, mtmTotals.FxTransactionDisplayUnit * 10);
            AssertIBKRMoneyWithin(
                preciseFxTransaction,
                mtmTotals.FxTransaction,
                displayTolerance,
                "IBKR MTM FX transaction display");

            builder.Add(
                new Currency(preciseFxTransaction, baseCurrency),
                "交易价格影响",
                $"FxTransaction/precise={preciseFxTransaction}/cashReport={cashReportFxTransaction}/mtm={mtmTotals.FxTransaction}",
                destAccount: baseCurrency.ToString());
        }

        private IBKRPositionQuantityChangeTotals ParseIBKRPositionQuantityChangeRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            IBKRTransactionTotals tradeTotals)
        {
            var coveredTransactionImpacts = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            decimal holdingValueChangeTotal = 0;
            decimal cashChangeTotal = 0;
            decimal coveredTransactionTotal = 0;
            foreach (var position in EnumerateIBKRStockBondPositionRows(report, contractInfos))
            {
                var row = position.Row;
                if (position.Currency != baseCurrency)
                    throw new MailParseException($"Parse IBKR Report Fail, Non-base Position Quantity Change: {FormatIBKRCsvRow(row)}");

                var contract = position.Contract;
                var holdingKey = GetIBKRHoldingKey(contract);
                var beginningQuantity = NormalizeIBKRHoldingQuantity(
                    contract,
                    position.BeginningQuantity);
                var endingQuantity = NormalizeIBKRHoldingQuantity(
                    contract,
                    position.EndingQuantity);
                var quantityChange = endingQuantity - beginningQuantity;
                if (quantityChange == 0)
                    continue;

                if (!tradeTotals.Holdings.TryGetValue(holdingKey, out var trade))
                    continue;
                if (trade.Quantity == 0 && trade.Proceeds == 0)
                    continue;
                if (trade.Quantity != quantityChange)
                    continue;
                if (trade.Currency != baseCurrency)
                    throw new MailParseException($"Parse IBKR Report Fail, Non-base Trade Summary: {FormatIBKRCsvRow(row)}");
                var previousMarketValue = ParseIBKRDecimalAt(row, 9, "position previous market value");
                var currentMarketValue = ParseIBKRDecimalAt(row, 10, "position current market value");
                var holdingPriceChange = ParseIBKRDecimalAt(row, 11, "position holding price change");
                var transactionPriceImpact = ParseIBKRDecimalAt(row, 12, "position transaction price impact");
                var holdingValueChange = currentMarketValue - previousMarketValue - holdingPriceChange;
                var cashChange = trade.Proceeds;
                var preciseTransactionPriceImpact = holdingValueChange + cashChange;
                AssertIBKRMoneyFieldEquals(
                    preciseTransactionPriceImpact,
                    transactionPriceImpact,
                    row.Fields[12],
                    $"IBKR position quantity change transaction {contract.Code}",
                    MidpointRounding.ToEven);

                if (holdingValueChange != 0)
                {
                    builder.Add(
                        new Currency(holdingValueChange, baseCurrency),
                        "持仓数量价值变化",
                        $"PositionQuantityChange/{FormatIBKRCsvRow(row)}",
                        isInternal: true,
                        destAccount: contract.Code,
                        holdingQuantity: quantityChange,
                        holding: contract);
                }

                if (cashChange != 0)
                {
                    builder.Add(
                        new Currency(cashChange, baseCurrency),
                        "持仓交易现金变化",
                        $"PositionQuantityCash/{trade.Source}",
                        isInternal: true,
                        destAccount: contract.Code,
                        holding: contract);
                }

                if (!coveredTransactionImpacts.TryAdd(holdingKey, preciseTransactionPriceImpact))
                    throw new MailParseException($"Parse IBKR Report Fail, Duplicate covered position transaction: {contract.Code}");
                holdingValueChangeTotal += holdingValueChange;
                cashChangeTotal += cashChange;
                coveredTransactionTotal += preciseTransactionPriceImpact;
            }

            AssertIBKRMoneyEquals(
                coveredTransactionTotal,
                holdingValueChangeTotal + cashChangeTotal,
                "IBKR covered position transaction records");
            return new IBKRPositionQuantityChangeTotals(
                coveredTransactionImpacts,
                holdingValueChangeTotal,
                cashChangeTotal,
                coveredTransactionTotal);
        }

        private static IEnumerable<IBKRStockBondPositionRow> EnumerateIBKRStockBondPositionRows(
            IBKRCsvReport report,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            foreach (var row in report.OptionalDataRows("持仓与以市值计的盈亏"))
            {
                if (!IsIBKRPositionSummaryRow(row))
                    continue;

                var group = row.Fields[1];
                if (!IsIBKRStockOrBondGroup(group))
                    continue;

                yield return new IBKRStockBondPositionRow(
                    row,
                    ResolveIBKRContract(row.Fields[3], group, contractInfos),
                    ParseIBKRCurrencyType(row.Fields[2]),
                    ParseIBKRIntegerQuantityAt(row, 5, "beginning holding quantity"),
                    ParseIBKRIntegerQuantityAt(row, 6, "ending holding quantity"));
            }
        }

        private IBKRMtmTotals ParseIBKRMtmRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            Dictionary<string, decimal> coveredTransactionImpacts)
        {
            var rows = report.OptionalDataRows("按市值计算的表现总结").ToList();
            if (rows.Count == 0)
                return new IBKRMtmTotals(false, 0, 0, 0, 0, 0, 0, 0);

            var cashFxTranslation = ParseIBKRCashFxTranslation(report, baseCurrency);
            decimal holdingTotal = 0;
            decimal transactionTotal = 0;
            decimal commissionTotal = 0;
            decimal otherTotal = 0;
            decimal brokerInterestOtherTotal = 0;
            decimal fxTransactionTotal = 0;
            decimal fxTransactionDisplayUnit = 0;
            decimal? grandHoldingTotal = null;
            decimal? grandTransactionTotal = null;
            decimal? grandCommissionTotal = null;
            decimal? grandOtherTotal = null;
            var pendingCoveredTransactionImpacts = new Dictionary<string, decimal>(coveredTransactionImpacts, StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 12);
                var assetClass = row.Fields[0];
                if (assetClass.StartsWith("总数", StringComparison.Ordinal))
                {
                    AssertIBKRMtmRowTotal(row);
                    continue;
                }

                if (assetClass.StartsWith("总计", StringComparison.Ordinal))
                {
                    AssertIBKRMtmRowTotal(row);
                    grandHoldingTotal = ParseIBKRDecimalAt(row, 6, "MTM grand holding");
                    grandTransactionTotal = ParseIBKRDecimalAt(row, 7, "MTM grand transaction");
                    grandCommissionTotal = ParseIBKRDecimalAt(row, 8, "MTM grand commission");
                    grandOtherTotal = ParseIBKRDecimalAt(row, 9, "MTM grand other");
                    continue;
                }

                if (assetClass == "支付和收到的经纪商利息")
                {
                    AssertIBKRMtmStandaloneTotalRow(row);
                    brokerInterestOtherTotal += ParseIBKRDecimalAt(row, 10, "MTM broker interest other");
                    continue;
                }

                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown MTM Asset Class: {FormatIBKRCsvRow(row)}");

                AssertIBKRMtmRowTotal(row);
                var holding = ParseIBKRDecimalAt(row, 6, "MTM holding");
                var transaction = ParseIBKRDecimalAt(row, 7, "MTM transaction");
                var commission = ParseIBKRDecimalAt(row, 8, "MTM commission");
                var other = ParseIBKRDecimalAt(row, 9, "MTM other");
                var contract = assetClass == "外汇"
                    ? null
                    : ResolveIBKRContract(row.Fields[1], assetClass, contractInfos);

                holdingTotal += holding;
                transactionTotal += transaction;
                commissionTotal += commission;
                otherTotal += other;
                if (assetClass == "外汇")
                {
                    fxTransactionTotal += transaction;
                    fxTransactionDisplayUnit = Math.Max(fxTransactionDisplayUnit, GetIBKRReportFieldUnit(row.Fields[7]));
                    continue;
                }

                builder.Add(
                    new Currency(holding, baseCurrency),
                    "持仓价格变动",
                    $"MTM/{FormatIBKRCsvRow(row)}",
                    destAccount: contract!.Code,
                    holding: contract);

                var holdingKey = contract is null
                    ? $"FX\t{row.Fields[1]}"
                    : GetIBKRHoldingKey(contract);
                if (coveredTransactionImpacts.TryGetValue(holdingKey, out var coveredTransactionImpact))
                {
                    AssertIBKRMoneyFieldEquals(
                        coveredTransactionImpact,
                        transaction,
                        row.Fields[7],
                        $"IBKR covered MTM transaction {contract?.Code ?? row.Fields[1]}",
                        MidpointRounding.ToEven);
                    pendingCoveredTransactionImpacts.Remove(holdingKey);
                }
                else
                {
                    builder.Add(
                        new Currency(transaction, baseCurrency),
                        "交易价格影响",
                        $"MTM/{FormatIBKRCsvRow(row)}",
                        destAccount: contract?.Code ?? row.Fields[1],
                        holding: contract);
                }
            }

            if (grandHoldingTotal.HasValue)
            {
                AssertIBKRMoneyEquals(grandHoldingTotal.Value, holdingTotal, "IBKR MTM grand holding");
                AssertIBKRMoneyEquals(grandTransactionTotal!.Value, transactionTotal, "IBKR MTM grand transaction");
                AssertIBKRMoneyEquals(grandCommissionTotal!.Value, commissionTotal, "IBKR MTM grand commission");
                AssertIBKRMoneyEquals(grandOtherTotal!.Value, otherTotal, "IBKR MTM grand other");
            }

            if (pendingCoveredTransactionImpacts.Count != 0)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, Missing covered MTM transactions: {String.Join(", ", pendingCoveredTransactionImpacts.Keys.OrderBy(key => key, StringComparer.OrdinalIgnoreCase))}");
            }

            builder.Add(
                new Currency(cashFxTranslation, baseCurrency),
                "现金外汇换算收益/损失",
                $"CashFxTranslation/{cashFxTranslation}",
                destAccount: baseCurrency.ToString());
            return new IBKRMtmTotals(true, holdingTotal, transactionTotal, commissionTotal, otherTotal + brokerInterestOtherTotal, cashFxTranslation, fxTransactionTotal, fxTransactionDisplayUnit);
        }

        private static decimal ParseIBKRCashFxTranslation(IBKRCsvReport report, CurrencyType baseCurrency)
        {
            var preciseValue = 0m;
            var reportedValue = 0m;
            foreach (var row in report.OptionalDataRows("持仓与以市值计的盈亏"))
            {
                if (!IsIBKRPositionSummaryRow(row) || row.Fields[1] != "外汇")
                    continue;

                var positionValues = ParseIBKRFxPositionValues(row, baseCurrency, "IBKR FX MTM");
                if (positionValues.CashCurrency == baseCurrency)
                    continue;

                var transaction = ParseIBKRDecimalAt(row, 12, "FX MTM transaction");
                var commission = ParseIBKRDecimalAt(row, 13, "FX MTM commission");
                var other = ParseIBKRDecimalAt(row, 14, "FX MTM other");
                var reportedHolding = ParseIBKRDecimalAt(row, 11, "FX MTM holding");
                var reportedTotal = ParseIBKRDecimalAt(row, 15, "FX MTM total");
                var hasQuantityChange = positionValues.PreviousQuantity != positionValues.CurrentQuantity;
                var hasCashFlow = transaction != 0 || commission != 0 || other != 0;
                var preciseHolding = !hasQuantityChange && !hasCashFlow
                    ? positionValues.PriorPositionRateImpact
                    : RoundIBKRMoneyForReportField(positionValues.PriorPositionRateImpact, row.Fields[11]) == reportedHolding
                    ? positionValues.PriorPositionRateImpact
                    : reportedHolding;
                if (hasQuantityChange || hasCashFlow)
                    AssertIBKRMoneyFieldEquals(preciseHolding, reportedHolding, row.Fields[11], $"IBKR FX MTM holding display {FormatIBKRCsvRow(row)}");
                else
                    AssertIBKRMoneyEquals(reportedHolding, reportedTotal, $"IBKR FX MTM no-flow holding display {FormatIBKRCsvRow(row)}");
                AssertIBKRMoneyFieldEquals(
                    reportedHolding + transaction + commission + other,
                    reportedTotal,
                    row.Fields[15],
                    $"IBKR FX MTM row total {FormatIBKRCsvRow(row)}");

                preciseValue += preciseHolding;
                reportedValue += reportedHolding;
            }

            ValidateIBKRCashFxTranslationDisplay(report, reportedValue);
            return preciseValue;
        }

        private static IBKRFxPositionValues ParseIBKRFxPositionValues(
            IBKRCsvRow row,
            CurrencyType baseCurrency,
            string context)
        {
            var valueCurrency = ParseIBKRCurrencyType(row.Fields[2]);
            if (valueCurrency != baseCurrency)
                throw new MailParseException($"Parse IBKR Report Fail, Non-base FX Position Value Currency: {FormatIBKRCsvRow(row)}");

            var cashCurrency = ParseIBKRCurrencyType(row.Fields[3]);
            var previousQuantity = ParseIBKRDecimalAt(row, 5, $"{context} previous quantity");
            var currentQuantity = ParseIBKRDecimalAt(row, 6, $"{context} current quantity");
            var previousRate = ParseIBKRDecimalAt(row, 7, $"{context} previous rate");
            var currentRate = ParseIBKRDecimalAt(row, 8, $"{context} current rate");
            var calculatedPreviousMarketValue = previousQuantity * previousRate;
            var calculatedCurrentMarketValue = currentQuantity * currentRate;
            var priorPositionRateImpact = previousQuantity * (currentRate - previousRate);
            var reportedPreviousMarketValue = ParseIBKRDecimalAt(row, 9, $"{context} previous market value");
            var reportedCurrentMarketValue = ParseIBKRDecimalAt(row, 10, $"{context} current market value");
            AssertIBKRMoneyFieldEquals(calculatedPreviousMarketValue, reportedPreviousMarketValue, row.Fields[9], $"{context} previous market value display {FormatIBKRCsvRow(row)}");
            AssertIBKRMoneyFieldEquals(calculatedCurrentMarketValue, reportedCurrentMarketValue, row.Fields[10], $"{context} current market value display {FormatIBKRCsvRow(row)}");
            return new IBKRFxPositionValues(
                cashCurrency,
                previousQuantity,
                currentQuantity,
                calculatedPreviousMarketValue,
                calculatedCurrentMarketValue,
                priorPositionRateImpact);
        }

        private static void ValidateIBKRCashFxTranslationDisplay(IBKRCsvReport report, decimal reportedValue)
        {
            var cashReportValue = 0m;
            var hasCashReportValue = false;
            foreach (var row in report.RequireDataRows("现金报告"))
            {
                AssertIBKRCashReportFieldCount(row);
                if (row.Fields[0] != "现金外汇换算收益/损失")
                    continue;

                cashReportValue += ParseIBKRDecimalAt(row, 2, "cash FX translation");
                hasCashReportValue = true;
            }

            if (hasCashReportValue)
                AssertIBKRMoneyEquals(reportedValue, cashReportValue, "IBKR cash FX translation cash report display");
            else
                AssertIBKRMoneyEquals(0, reportedValue, "IBKR missing cash FX translation cash report display");

            // 精确值逐行和 MTM 显示字段校验；这里校验现金报告与 MTM 总表的显示口径一致。
            var mtmReportValue = 0m;
            foreach (var row in report.OptionalDataRows("按市值计算的表现总结"))
            {
                AssertIBKRFieldCount(row, 12);
                if (row.Fields[0] != "外汇")
                    continue;

                var transaction = ParseIBKRDecimalAt(row, 7, "MTM FX transaction");
                var commission = ParseIBKRDecimalAt(row, 8, "MTM FX commission");
                var other = ParseIBKRDecimalAt(row, 9, "MTM FX other");
                var total = ParseIBKRDecimalAt(row, 10, "MTM FX total");
                AssertIBKRMoneyEquals(
                    ParseIBKRDecimalAt(row, 6, "MTM FX holding") + transaction + commission + other,
                    total,
                    $"IBKR MTM FX row total {FormatIBKRCsvRow(row)}");

                mtmReportValue += ParseIBKRDecimalAt(row, 6, "MTM FX holding");
            }

            AssertIBKRMoneyEquals(reportedValue, mtmReportValue, "IBKR cash FX translation MTM display");
        }

        private static void AssertIBKRMtmRowTotal(IBKRCsvRow row)
        {
            var total = ParseIBKRDecimalAt(row, 10, "MTM total");
            var components =
                ParseIBKRDecimalAt(row, 6, "MTM holding")
                + ParseIBKRDecimalAt(row, 7, "MTM transaction")
                + ParseIBKRDecimalAt(row, 8, "MTM commission")
                + ParseIBKRDecimalAt(row, 9, "MTM other");
            AssertIBKRMoneyEquals(total, components, $"IBKR MTM row total {FormatIBKRCsvRow(row)}");
        }

        private static void AssertIBKRMtmStandaloneTotalRow(IBKRCsvRow row)
        {
            AssertIBKRFieldCount(row, 12);
            if (row.Fields
                .Select((field, index) => (field, index))
                .Where(item => item.index != 0 && item.index != 10 && item.index != 11)
                .Any(item => !String.IsNullOrWhiteSpace(item.field)))
            {
                throw new MailParseException($"Parse IBKR Report Fail, Invalid MTM Standalone Total Row: {FormatIBKRCsvRow(row)}");
            }

            _ = ParseIBKRDecimalAt(row, 10, "MTM standalone total");
        }

        private static IBKRTransactionTotals ParseIBKRTradeSummary(
            IBKRCsvReport report,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows("按代码显示的交易总结").ToList();
            var assetRows = report.OptionalDataRows("按资产类型显示的交易总结").ToList();
            if (rows.Count == 0 && assetRows.Count == 0)
                return new IBKRTransactionTotals(false, 0, 0, 0, 0, 0, 0, false, []);

            decimal buyTotal = 0;
            decimal sellTotal = 0;
            decimal stockBondProceeds = 0;
            int buyQuantity = 0;
            int sellQuantity = 0;
            decimal? reportedBuyTotal = null;
            decimal? reportedSellTotal = null;
            var hasFxTrades = false;
            var trades = new Dictionary<string, IBKRTradeSummaryData>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 9);
                var assetClass = row.Fields[0];
                if (assetClass.StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedBuyTotal = (reportedBuyTotal ?? 0) + ParseIBKRDecimalOrZero(row.Fields[5]);
                    reportedSellTotal = (reportedSellTotal ?? 0) + ParseIBKRDecimalOrZero(row.Fields[8]);
                    continue;
                }

                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Trade Summary Asset Class: {FormatIBKRCsvRow(row)}");

                var rowBuyTotal = ParseIBKRDecimalOrZero(row.Fields[5]);
                var rowSellTotal = ParseIBKRDecimalOrZero(row.Fields[8]);
                buyTotal += rowBuyTotal;
                sellTotal += rowSellTotal;
                if (assetClass == "外汇" && (rowBuyTotal != 0 || rowSellTotal != 0))
                    hasFxTrades = true;
                if (assetClass is "股票" or "债券")
                {
                    stockBondProceeds += rowBuyTotal + rowSellTotal;
                    var rawBuyQuantity = ParseIBKRIntegerQuantityOrZero(row.Fields[3], "trade buy quantity", row);
                    var rawSellQuantity = ParseIBKRIntegerQuantityOrZero(row.Fields[6], "trade sell quantity", row);
                    buyQuantity += rawBuyQuantity;
                    sellQuantity += rawSellQuantity;
                    var currency = ParseIBKRCurrencyType(row.Fields[1]);
                    var contract = ResolveIBKRContract(row.Fields[2], assetClass, contractInfos);
                    var holdingKey = GetIBKRHoldingKey(contract);
                    var buyHoldingQuantity = NormalizeIBKRHoldingQuantity(contract, rawBuyQuantity);
                    var sellHoldingQuantity = NormalizeIBKRHoldingQuantity(contract, rawSellQuantity);
                    if (trades.TryGetValue(holdingKey, out var existing))
                    {
                        trades[holdingKey] = existing.Add(
                            currency,
                            buyHoldingQuantity,
                            sellHoldingQuantity,
                            rowBuyTotal,
                            rowSellTotal,
                            FormatIBKRCsvRow(row));
                    }
                    else
                    {
                        trades[holdingKey] = new IBKRTradeSummaryData(
                            contract,
                            currency,
                            buyHoldingQuantity,
                            sellHoldingQuantity,
                            rowBuyTotal,
                            rowSellTotal,
                            FormatIBKRCsvRow(row));
                    }
                }
            }

            if (!hasFxTrades && reportedBuyTotal.HasValue)
                AssertIBKRMoneyEquals(reportedBuyTotal.Value, buyTotal, "IBKR trade buy summary");
            if (!hasFxTrades && reportedSellTotal.HasValue)
                AssertIBKRMoneyEquals(reportedSellTotal.Value, sellTotal, "IBKR trade sell summary");

            foreach (var row in assetRows)
            {
                AssertIBKRFieldCount(row, 6);
                var assetClass = row.Fields[0];
                if (!IBKRAssetGroups.Contains(assetClass) && !assetClass.StartsWith("总数", StringComparison.Ordinal))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Trade Asset Summary Class: {FormatIBKRCsvRow(row)}");
            }

            return new IBKRTransactionTotals(true, buyTotal + sellTotal, buyTotal, sellTotal, stockBondProceeds, 0, buyQuantity + sellQuantity, hasFxTrades, trades);
        }

        private IBKRCommissionTotals ParseIBKRCommissionRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows("佣金细节").ToList();
            var adjustmentRows = report.OptionalDataRows(IBKRCommissionAdjustmentSection).ToList();
            if (rows.Count == 0 && adjustmentRows.Count == 0)
                return new IBKRCommissionTotals(false, 0);

            decimal commissionTotal = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 12);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedTotal = (reportedTotal ?? 0) + ParseIBKRDecimalAt(row, 5, "commission total");
                    ValidateIBKRCommissionComponentFields(row);
                    continue;
                }

                var assetClass = row.Fields[0];
                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Commission Asset Class: {FormatIBKRCsvRow(row)}");

                var rowCurrency = ParseIBKRCurrencyType(row.Fields[1]);
                var currency = assetClass == "外汇" ? baseCurrency : rowCurrency;
                IBKRContractInfo? contract = null;
                if (assetClass is "股票" or "债券")
                {
                    contract = ResolveIBKRContract(row.Fields[2], assetClass, contractInfos);
                    var rawQuantity = ParseIBKRIntegerQuantityAt(row, 4, "commission quantity");
                    _ = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
                }
                else
                {
                    _ = ParseIBKRDecimalAt(row, 4, "FX commission quantity");
                }

                // 佣金 record 的金额只取“佣金细节”表里的“佣金”列。这个列是 IBKR 对该次成交
                // 实际收取佣金的最精确来源，例如 2026-05-14 的两条 NVDL 佣金分别是
                // -0.025505784 和 -1.025505784，合计 -1.051011568。
                //
                // 不要用“现金报告”“净资产值变更”或“按市值计算的表现总结”里的佣金汇总值替代它们；
                // 那些表通常只显示 8 位小数，例如 -1.051011568 会显示为 -1.05101157。
                // 汇总表的数值只用于 Round(佣金明细合计, 8) == 汇总显示值 的校验，不能作为
                // record 金额入库，否则会把报表显示精度差异写进真实流水。
                var commission = ParseIBKRDecimalAt(row, 5, "commission");
                ValidateIBKRCommissionComponentFields(row);
                commissionTotal += commission;
                builder.Add(
                    new Currency(commission, currency),
                    "佣金",
                    $"Commission/{FormatIBKRCsvRow(row)}",
                    date: ParseIBKRDateTime(row.Fields[3]),
                    destAccount: contract?.Code ?? row.Fields[2],
                    holding: contract);
            }

            if (reportedTotal.HasValue && reportedTotal.Value != commissionTotal)
            {
                _ = reportedTotal.Value - commissionTotal;
            }

            if (reportedTotal.HasValue && reportedTotal.Value == commissionTotal)
                AssertIBKRMoneyEquals(reportedTotal.Value, commissionTotal, "IBKR commission total");

            _ = ParseIBKRCommissionAdjustmentRecords(report, builder, baseCurrency, contractInfos);
            return new IBKRCommissionTotals(true, commissionTotal);
        }

        private decimal ParseIBKRCommissionAdjustmentRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows(IBKRCommissionAdjustmentSection).ToList();
            if (rows.Count == 0)
                return 0;

            decimal total = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 5);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    if (reportedTotal.HasValue)
                        throw new MailParseException($"Parse IBKR Report Fail, Duplicate Commission Adjustment Total: {FormatIBKRCsvRow(row)}");
                    reportedTotal = ParseIBKRDecimalAt(row, 3, "commission adjustment total");
                    continue;
                }

                var currency = ParseIBKRCurrencyType(row.Fields[0]);
                if (currency != baseCurrency)
                    throw new MailParseException($"Parse IBKR Report Fail, Non-base Commission Adjustment: {FormatIBKRCsvRow(row)}");
                var amount = ParseIBKRDecimalAt(row, 3, "commission adjustment");
                var holding = ResolveIBKRCommissionAdjustmentHolding(row, contractInfos);
                total += amount;
                _ = ParseIBKRDate(row.Fields[1]);
                _ = holding;
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, total, "IBKR commission adjustment total");

            return total;
        }

        private static IBKRContractInfo? ResolveIBKRCommissionAdjustmentHolding(
            IBKRCsvRow row,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var code = row.Fields[4].Trim();
            if (String.IsNullOrWhiteSpace(code))
            {
                var match = Regex.Match(row.Fields[2], @"\((?<code>[^()]+)\)\s*$");
                if (match.Success)
                    code = match.Groups["code"].Value.Trim();
            }

            if (String.IsNullOrWhiteSpace(code))
                return null;
            return ResolveIBKRContract(code, "股票", contractInfos);
        }

        private static void ValidateIBKRCommissionComponentFields(IBKRCsvRow row)
        {
            _ = ParseIBKRDecimalAt(row, 6, "broker execution fee");
            _ = ParseIBKRDecimalAt(row, 7, "broker clearing fee");
            _ = ParseIBKRDecimalAt(row, 8, "third-party execution fee");
            _ = ParseIBKRDecimalAt(row, 9, "third-party clearing fee");
            _ = ParseIBKRDecimalAt(row, 10, "third-party transaction fee");
            _ = ParseIBKRDecimalAt(row, 11, "other commission fee");
        }

        private IBKRInterestTotals ParseIBKRInterestAccrualRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency)
        {
            var rows = report.OptionalDataRows("应计利息").ToList();
            if (rows.Count == 0)
                return new IBKRInterestTotals(false, 0);

            decimal total = 0;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 3);
                if (row.Fields[0] != "基础货币总结")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Interest Accrual Currency: {FormatIBKRCsvRow(row)}");

                var label = row.Fields[1];
                var amount = ParseIBKRDecimalAt(row, 2, "interest accrual");
                if (label == "期初应计余额" || label == "期末应计余额")
                    continue;
                if (label != "应计利息" && label != "应计转回" && label != "外汇换算")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Interest Accrual Row: {FormatIBKRCsvRow(row)}");

                total += amount;
                builder.Add(
                    new Currency(amount, baseCurrency),
                    label,
                    $"InterestAccruals/{FormatIBKRCsvRow(row)}",
                    destAccount: "ACCRUED_INTEREST");
            }

            return new IBKRInterestTotals(true, total);
        }

        private static IBKRInterestTotals ParseIBKRDebitInterestDetails(IBKRCsvReport report)
        {
            var rows = report.OptionalDataRows("借方利息细节").ToList();
            if (rows.Count == 0)
                return new IBKRInterestTotals(false, 0);

            decimal total = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 11);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedTotal = ParseIBKRDecimalAt(row, 9, "debit interest total");
                    continue;
                }

                _ = ParseIBKRCurrencyType(row.Fields[0]);
                _ = ParseIBKRDate(row.Fields[1]);
                total += ParseIBKRDecimalAt(row, 9, "debit interest");
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, total, "IBKR debit interest total");

            return new IBKRInterestTotals(true, total);
        }

        private static IBKRInterestTotals ParseIBKRDividendAccruals(IBKRCsvReport report)
        {
            var rows = report.OptionalDataRows("未平仓应计股息").ToList();
            if (rows.Count == 0)
                return new IBKRInterestTotals(false, 0);

            decimal total = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 12);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedTotal = ParseIBKRDecimalAt(row, 10, "open dividend net total");
                    continue;
                }

                if (row.Fields[0] != "股票")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Dividend Accrual Asset Class: {FormatIBKRCsvRow(row)}");

                _ = ParseIBKRCurrencyType(row.Fields[1]);
                _ = ParseIBKRDate(row.Fields[3]);
                _ = ParseIBKRDate(row.Fields[4]);
                var tax = Math.Abs(ParseIBKRDecimalAt(row, 6, "dividend tax"));
                var fee = Math.Abs(ParseIBKRDecimalAt(row, 7, "dividend fee"));
                var gross = ParseIBKRDecimalAt(row, 9, "gross dividend");
                var net = ParseIBKRDecimalAt(row, 10, "net dividend");
                AssertIBKRMoneyEquals(gross - tax - fee, net, "IBKR dividend accrual net");
                total += net;
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, total, "IBKR dividend accrual total");

            return new IBKRInterestTotals(true, total);
        }

        private IBKRInterestTotals ParseIBKRDividendAccrualChangeRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows(IBKRDividendAccrualChangeSection).ToList();
            if (rows.Count == 0)
                return new IBKRInterestTotals(false, 0);

            decimal detailTotal = 0;
            decimal? reportedTotal = null;
            decimal? openingTotal = null;
            decimal? endingTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 13);
                var assetClass = row.Fields[0];
                if (assetClass.StartsWith("期初应计股息", StringComparison.Ordinal))
                {
                    openingTotal = ParseIBKRDecimalAt(row, 11, "opening dividend accrual");
                    continue;
                }

                if (assetClass.StartsWith("期末应计股息", StringComparison.Ordinal))
                {
                    endingTotal = ParseIBKRDecimalAt(row, 11, "ending dividend accrual");
                    continue;
                }

                if (assetClass.StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedTotal = ParseIBKRDecimalAt(row, 11, "dividend accrual change total");
                    continue;
                }

                if (assetClass != "股票")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Dividend Accrual Change Asset Class: {FormatIBKRCsvRow(row)}");

                var currency = ParseIBKRCurrencyType(row.Fields[1]);
                if (currency != baseCurrency)
                    throw new MailParseException($"Parse IBKR Report Fail, Non-base Dividend Accrual Change: {FormatIBKRCsvRow(row)}");

                var contract = ResolveIBKRContract(row.Fields[2], assetClass, contractInfos);
                var date = ParseIBKRDate(row.Fields[3]);
                _ = ParseIBKRDate(row.Fields[4]);
                _ = ParseIBKRDate(row.Fields[5]);
                var rawQuantity = ParseIBKRIntegerQuantityAt(row, 6, "dividend accrual quantity");
                var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
                _ = ParseIBKRDecimalAt(row, 7, "dividend accrual tax");
                _ = ParseIBKRDecimalAt(row, 8, "dividend accrual fee");
                _ = ParseIBKRDecimalAt(row, 10, "dividend accrual gross");
                var net = ParseIBKRDecimalAt(row, 11, "dividend accrual net");

                detailTotal += net;
                builder.Add(
                    new Currency(net, currency),
                    "应计股息",
                    $"DividendAccrualChange/{FormatIBKRCsvRow(row)}",
                    date: date,
                    destAccount: contract.Code,
                    holding: contract);
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, detailTotal, "IBKR dividend accrual change total");
            _ = openingTotal;
            _ = endingTotal;

            return new IBKRInterestTotals(true, detailTotal);
        }

        private IBKRTransferTotals ParseIBKRTransferRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Account account,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows(IBKRTransferSection).ToList();
            if (rows.Count == 0)
                return new IBKRTransferTotals(false, 0);

            decimal detailTotal = 0;
            decimal? reportedMarketValue = null;
            decimal? reportedRealizedPnl = null;
            decimal? reportedCashAmount = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 14);
                var assetClass = row.Fields[0];
                if (assetClass.StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedMarketValue = (reportedMarketValue ?? 0) + ParseIBKRDecimalAt(row, 10, "transfer market value total");
                    reportedRealizedPnl = (reportedRealizedPnl ?? 0) + ParseIBKRDecimalAt(row, 11, "transfer realized pnl total");
                    reportedCashAmount = (reportedCashAmount ?? 0) + ParseIBKRDecimalAt(row, 12, "transfer cash amount total");
                    continue;
                }

                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Transfer Asset Class: {FormatIBKRCsvRow(row)}");

                var currency = ParseIBKRCurrencyType(row.Fields[1]);
                if (currency != baseCurrency)
                    throw new MailParseException($"Parse IBKR Report Fail, Non-base Transfer: {FormatIBKRCsvRow(row)}");

                var type = row.Fields[4];
                if (type != "内部")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Transfer Type: {FormatIBKRCsvRow(row)}");

                var direction = row.Fields[5];
                if (direction is not ("进" or "出"))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Transfer Direction: {FormatIBKRCsvRow(row)}");

                var contract = ResolveIBKRContract(row.Fields[2], assetClass, contractInfos);
                var rawQuantity = ParseIBKRIntegerQuantityAt(row, 8, "transfer quantity");
                var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
                var marketValue = ParseIBKRDecimalAt(row, 10, "transfer market value");
                var realizedPnl = ParseIBKRDecimalAt(row, 11, "transfer realized pnl");
                var cashAmount = ParseIBKRDecimalAt(row, 12, "transfer cash amount");
                if (realizedPnl != 0)
                    throw new MailParseException($"Parse IBKR Report Fail, Unsupported Transfer Realized P/L: {FormatIBKRCsvRow(row)}");

                var amount = marketValue + cashAmount;
                detailTotal += amount;
                builder.Add(
                    new Currency(amount, currency),
                    "内部转账",
                    $"Transfer/{FormatIBKRCsvRow(row)}",
                    isInternal: true,
                    date: ParseIBKRDate(row.Fields[3]),
                    destAccount: BuildIBKRTransferDestAccount(account, row.Fields[7], direction),
                    holdingQuantity: quantity,
                    holding: contract);
            }

            var reportedTotal = (reportedMarketValue ?? 0) + (reportedCashAmount ?? 0);
            if (reportedMarketValue.HasValue || reportedCashAmount.HasValue)
                AssertIBKRMoneyEquals(reportedTotal, detailTotal, "IBKR transfer total");
            if (reportedRealizedPnl.HasValue)
                AssertIBKRMoneyEquals(0, reportedRealizedPnl.Value, "IBKR transfer realized pnl total");

            return new IBKRTransferTotals(true, detailTotal);
        }

        private static string BuildIBKRTransferDestAccount(Account account, string transferAccount, string direction)
        {
            var target = transferAccount.Trim();
            if (String.IsNullOrWhiteSpace(target) || target == "--")
                return target;
            if (!target.StartsWith("IBKR_", StringComparison.OrdinalIgnoreCase))
                target = $"IBKR_{target}";

            return direction == "出"
                ? $"{account.name}->{target}"
                : $"{target}->{account.name}";
        }

        private IBKRCashTotals ParseIBKRCashRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            IBKRTransactionTotals transactionTotals,
            IBKRCommissionTotals commissionTotals)
        {
            var rows = report.RequireDataRows("现金报告");
            decimal commission = 0;
            decimal tradeBuy = 0;
            decimal tradeSell = 0;
            decimal nonTradeCashFlow = 0;
            decimal cashFxTranslation = 0;
            foreach (var row in rows)
            {
                AssertIBKRCashReportFieldCount(row);
                if (row.Fields[1] != "基础货币总结")
                    continue;

                var label = row.Fields[0];
                var amount = ParseIBKRDecimalAt(row, 2, "cash amount");
                switch (label)
                {
                    case "期初现金":
                    case "期末现金":
                    case "期末已结算现金":
                    case "期初抵押品价值":
                    case "净证券借出活动":
                    case "期末抵押品价值":
                    case "净现金结余":
                    case "净结算后现金余额":
                        break;
                    case "佣金":
                        commission += amount;
                        nonTradeCashFlow += amount;
                        if (!commissionTotals.HasData && amount != 0)
                        {
                            throw new MailParseException(
                                $"Parse IBKR Report Fail, Cash commission without commission details: {FormatIBKRCsvRow(row)}");
                        }

                        break;
                    case "交易（买入）":
                        tradeBuy += amount;
                        break;
                    case "交易（卖出）":
                        tradeSell += amount;
                        break;
                    case "支付和收到的经纪商利息":
                    case "支付和收到的债券利息":
                    case "股息":
                    case "代替股息的支付":
                    case "代扣税款":
                        nonTradeCashFlow += amount;
                        builder.Add(new Currency(amount, baseCurrency), label, $"CashReport/{FormatIBKRCsvRow(row)}", destAccount: baseCurrency.ToString());
                        break;
                    case "现金外汇换算收益/损失":
                        cashFxTranslation += amount;
                        // 已由 MTM 外汇行体现，避免同一基础货币折算影响重复生成 record。
                        break;
                    case "存款":
                    case "取款":
                    case "转入":
                    case "转出":
                    case "内部转账":
                    case "账户转账":
                        nonTradeCashFlow += amount;
                        builder.Add(
                            new Currency(amount, baseCurrency),
                            label,
                            $"CashReport/{FormatIBKRCsvRow(row)}",
                            isInternal: true,
                            destAccount: baseCurrency.ToString());
                        break;
                    default:
                        throw new MailParseException($"Parse IBKR Report Fail, Unknown Cash Row: {FormatIBKRCsvRow(row)}");
                }
            }

            if (transactionTotals.HasData && !transactionTotals.HasFxTrades)
            {
                AssertIBKRMoneyEquals(tradeBuy, transactionTotals.BuyProceeds, "IBKR cash buy transactions");
                AssertIBKRMoneyEquals(tradeSell, transactionTotals.SellProceeds, "IBKR cash sell transactions");
            }

            return new IBKRCashTotals(commission, tradeBuy, tradeSell, nonTradeCashFlow, cashFxTranslation);
        }

        private List<Holding> ParseIBKRHoldings(
            IBKRCsvReport report,
            Account account,
            CurrencyType baseCurrency,
            List<Holding> positionHoldings,
            string sourceName,
            decimal cash,
            bool useBeginningValues = false)
        {
            var holdings = new List<Holding>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var holding in positionHoldings)
                AddIBKRHolding(holdings, seen, holding);

            AddIBKRHolding(holdings, seen, new Holding(baseCurrency.ToString(), HoldingType.Cash)
            {
                Account = account,
                desc = $"All currency cash balances converted to {baseCurrency}",
                displayText = baseCurrency == CurrencyType.USD ? "All currency($)" : $"All currency({baseCurrency})",
                currentPrice = new Currency(cash, baseCurrency)
            });

            foreach (var holding in ParseIBKRNavAdjustmentHoldings(report, account, baseCurrency, useBeginningValues))
                AddIBKRHolding(holdings, seen, holding);

            if (!useBeginningValues)
                ValidateIBKRStockPositionSummary(report, holdings);
            if (holdings.Count == 0)
                throw new MailParseException($"Parse IBKR Report Fail, Missing Holdings: {sourceName}");

            return holdings;
        }

        private static List<Holding> ParseIBKRNavAdjustmentHoldings(
            IBKRCsvReport report,
            Account account,
            CurrencyType baseCurrency,
            bool useBeginningValues)
        {
            var holdings = new List<Holding>();
            foreach (var row in report.RequireDataRows("净资产值"))
            {
                if (row.Fields.Count == 1)
                    continue;
                AssertIBKRFieldCount(row, 6);

                var component = row.Fields[0];
                if (component.StartsWith("总数", StringComparison.Ordinal))
                    continue;
                if (IsIBKRNavDetailComponent(component))
                    continue;
                if (component is "现金" or "股票" or "债券")
                    continue;

                var amount = ParseIBKRDecimalAt(
                    row,
                    useBeginningValues ? 1 : 4,
                    useBeginningValues ? "beginning NAV component amount" : "NAV component amount");
                if (component is "抵押品价值" or "借出证券")
                {
                    if (amount == 0)
                        continue;

                    var adjustmentCode = component == "抵押品价值"
                        ? "COLLATERAL_VALUE"
                        : "BORROWED_SECURITIES";
                    AddOrMergeIBKRSingleValueHolding(holdings, account, adjustmentCode, HoldingType.Accrued, component, amount, baseCurrency);
                    continue;
                }

                var (code, holdingType) = component switch
                {
                    "应计利息" => ("ACCRUED_INTEREST", HoldingType.Accrued),
                    "应计股息" => ("ACCRUED_DIVIDEND", HoldingType.Accrued),
                    _ => throw new MailParseException($"Parse IBKR Report Fail, Unknown NAV Component: {component}/{amount}")
                };
                if (amount == 0)
                    continue;

                AddOrMergeIBKRSingleValueHolding(holdings, account, code, holdingType, component, amount, baseCurrency);
            }

            return holdings;
        }

        private List<Holding> ParseIBKRPositionAndMtmHoldings(
            IBKRCsvReport report,
            Account account,
            Dictionary<string, IBKRContractInfo> contractInfos,
            out List<Holding> beginningHoldings)
        {
            var holdings = new List<Holding>();
            beginningHoldings = [];
            var beginningSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var position in EnumerateIBKRStockBondPositionRows(report, contractInfos))
            {
                var row = position.Row;
                var contract = position.Contract;
                var currency = position.Currency;
                var beginningQuantity = position.BeginningQuantity;
                if (beginningQuantity != 0)
                {
                    var beginningPrice = ParseIBKRDecimalAt(row, 7, "beginning holding price");
                    var beginningValue = ParseIBKRDecimalAt(row, 9, "beginning holding value");
                    AddIBKRHolding(
                        beginningHoldings,
                        beginningSeen,
                        CreateIBKRHolding(
                            account,
                            contract,
                            beginningQuantity,
                            beginningPrice,
                            beginningValue,
                            row.Fields[9],
                            currency,
                            row.Fields[4]));
                }

                var quantity = position.EndingQuantity;
                if (quantity == 0)
                    continue;

                var currentPrice = ParseIBKRDecimalAt(row, 8, "holding price");
                var currentValue = ParseIBKRDecimalAt(row, 10, "holding value");
                holdings.Add(CreateIBKRHolding(account, contract, quantity, currentPrice, currentValue, row.Fields[10], currency, row.Fields[4]));
            }

            return holdings;
        }

        private static bool IsIBKRStockOrBondGroup(string group)
        {
            return group is "股票" or "债券";
        }

        private static bool IsIBKRPositionSummaryRow(IBKRCsvRow row)
        {
            if (row.Fields.Count == 16 && row.Fields[0] == "Summary")
                return true;
            if (row.Fields.Count == 16 && row.Fields[0] == "Details")
                return false;
            if (row.Fields.Count > 1 && row.Fields[1] is "其他项目" or "其它项目")
                return false;
            if (row.Fields.Count == 16 && row.Fields[1].StartsWith("总数", StringComparison.Ordinal))
                return false;
            if (row.Fields.Count == 14 && row.Fields[1].StartsWith("总计", StringComparison.Ordinal))
                return false;

            if (row.Section == "持仓与以市值计的盈亏")
                throw new MailParseException($"Parse IBKR Report Fail, Unknown Position Row: {FormatIBKRCsvRow(row)}");
            return false;
        }

        private static Holding CreateIBKRHolding(
            Account account,
            IBKRContractInfo contract,
            int rawQuantity,
            decimal statementPrice,
            decimal currentValue,
            string currentValueText,
            CurrencyType currency,
            string rowDescription)
        {
            var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
            var currentPrice = statementPrice;
            if (contract.HoldingType == HoldingType.UST)
            {
                AssertIBKRMoneyFieldEquals(currentPrice * quantity, currentValue, currentValueText, $"IBKR bond holding value {contract.Code}");
            }
            else if (quantity != 0 && currentPrice * quantity != currentValue)
            {
                currentPrice = currentValue / quantity;
                AssertIBKRMoneyFieldEquals(currentPrice * quantity, currentValue, currentValueText, $"IBKR holding value {contract.Code}");
            }

            var description = String.IsNullOrWhiteSpace(rowDescription) ? contract.Description : rowDescription;
            return new Holding(contract.Code, contract.HoldingType)
            {
                Account = account,
                quantity = quantity,
                desc = description,
                displayText = contract.DisplayText,
                currentPrice = new Currency(currentPrice, currency)
            };
        }

        private static int NormalizeIBKRHoldingQuantity(IBKRContractInfo contract, int rawQuantity)
        {
            if (contract.HoldingType == HoldingType.UST)
            {
                if (rawQuantity % 100 != 0)
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid UST Quantity: {contract.Code}/{rawQuantity}");

                return rawQuantity / 100;
            }

            return rawQuantity;
        }

        private static void AddIBKRHolding(List<Holding> holdings, HashSet<string> seen, Holding holding)
        {
            var key = GetIBKRHoldingKey(holding);
            if (!seen.Add(key))
                throw new MailParseException($"Parse IBKR Report Fail, Duplicate Holding: {holding.code}/{holding.holdingType}");
            holdings.Add(holding);
        }

        private static void AddOrMergeIBKRSingleValueHolding(
            List<Holding> holdings,
            Account account,
            string code,
            HoldingType holdingType,
            string component,
            decimal amount,
            CurrencyType currency)
        {
            var existing = holdings.FirstOrDefault(holding =>
                holding.code == code && holding.holdingType == holdingType);
            if (existing is null)
            {
                holdings.Add(new Holding(code, holdingType)
                {
                    Account = account,
                    desc = $"IBKR NAV {component}",
                    displayText = component,
                    currentPrice = new Currency(amount, currency)
                });
                return;
            }

            if (existing.currentPrice.t != currency)
                throw new MailParseException($"Parse IBKR Report Fail, NAV adjustment currency mismatch: {code}/{holdingType}");
            existing.currentPrice = new Currency(existing.currentPrice.v + amount, currency);
        }

        private static string GetIBKRHoldingKey(Holding holding)
        {
            return $"{holding.code}\t{holding.holdingType}";
        }

        private static string GetIBKRHoldingKey(IBKRContractInfo contract)
        {
            return $"{contract.Code}\t{contract.HoldingType}";
        }

        private static IBKRContractInfo ResolveIBKRContract(
            string rawCode,
            string group,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var code = rawCode.Trim();
            if (contractInfos.TryGetValue(code, out var contract))
                return contract;

            var created = CreateIBKRContractInfo(group, code, code);
            if (contractInfos.TryGetValue(created.Code, out contract))
            {
                contractInfos[code] = contract;
                return contract;
            }

            contractInfos[created.Code] = created;
            if (!String.Equals(code, created.Code, StringComparison.OrdinalIgnoreCase))
                contractInfos[code] = created;
            return created;
        }

        private static string ExtractIBKRBondCode(string text)
        {
            var match = Regex.Match(text, @"T\s+\d+(?:(?:\.\d+)|(?:\s+\d+/\d+))?\s+\d{2}/\d{2}/\d{2}", RegexOptions.IgnoreCase);
            if (match.Success)
                return Regex.Replace(match.Value, @"\s+", " ").Trim();
            return text;
        }

        private static string NormalizeIBKRBondDisplayText(string code, string issuer)
        {
            var normalizedCode = ExtractIBKRBondCode(String.IsNullOrWhiteSpace(code) ? issuer : code);
            if (normalizedCode.StartsWith("T ", StringComparison.OrdinalIgnoreCase))
                return $"US {normalizedCode}";
            if (!String.IsNullOrWhiteSpace(issuer))
                return issuer;
            return normalizedCode;
        }

        private static void ValidateIBKRStockPositionSummary(IBKRCsvReport report, List<Holding> holdings)
        {
            var stockHoldings = holdings
                .Where(holding => holding.holdingType is HoldingType.NASDAQ or HoldingType.ARCA)
                .ToDictionary(holding => holding.code, holding => holding.quantity, StringComparer.OrdinalIgnoreCase);
            foreach (var row in report.OptionalDataRows("净股票持仓总结"))
            {
                AssertIBKRFieldCount(row, 8);
                if (row.Fields[0] != "股票")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Stock Position Summary Asset: {FormatIBKRCsvRow(row)}");

                var code = row.Fields[2];
                var inIBQuantity = ParseIBKRIntegerQuantityAt(row, 4, "in-IB stock quantity");
                var borrowedQuantity = ParseIBKRIntegerQuantityAt(row, 5, "borrowed stock quantity");
                var lentQuantity = ParseIBKRIntegerQuantityAt(row, 6, "lent stock quantity");
                var netQuantity = ParseIBKRIntegerQuantityAt(row, 7, "net stock quantity");
                if (inIBQuantity + borrowedQuantity + lentQuantity != netQuantity)
                    throw new MailParseException($"Parse IBKR Report Fail, Stock Position Summary Quantity Mismatch: {FormatIBKRCsvRow(row)}");
                if (!stockHoldings.TryGetValue(code, out var holdingQuantity))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Holding: {FormatIBKRCsvRow(row)}");
                if (holdingQuantity != inIBQuantity)
                    throw new MailParseException($"Parse IBKR Report Fail, Stock Holding Quantity Mismatch: {code}/{holdingQuantity}/{inIBQuantity}");
            }
        }

        private static decimal ParseIBKREndingCash(IBKRCsvReport report)
        {
            var row = report.RequireDataRows("现金报告").FirstOrDefault(row => row.Fields.Count > 2 && row.Fields[0] == "期末现金")
                ?? throw new MailParseException("Parse IBKR Report Fail, Missing Ending Cash");
            return ParseIBKRDecimalAt(row, 2, "ending cash");
        }

        private static decimal ParseIBKRStartingCash(IBKRCsvReport report)
        {
            return ParseIBKRCashReportAmount(report, IBKRStartingCashLabel, "starting cash");
        }

        private static decimal ParseIBKRCashReportAmount(IBKRCsvReport report, string label, string context)
        {
            var row = FindIBKRCashReportRow(report, label);
            return ParseIBKRDecimalAt(row, 2, context);
        }

        private const string IBKRStartingCashLabel = "期初现金";
        private const string IBKREndingCashLabel = "期末现金";
        private const string IBKRNavCashComponent = "现金";
        private const string IBKRNavTotalLabel = "总数";

        private static IBKRCsvRow FindIBKRCashReportRow(IBKRCsvReport report, string label)
        {
            return report.Sections.Values.SelectMany(section => section.DataRows).FirstOrDefault(row => row.Fields.Count > 2 && row.Fields[0] == label)
                ?? throw new MailParseException($"Parse IBKR Report Fail, Missing Cash Row: {label}");
        }

        private static IBKRPreciseNavValues ParseIBKRPreciseNavValues(IBKRCsvReport report, CurrencyType baseCurrency)
        {
            // IBKR 净资产总值表本身是汇总视图，底层实际计算依赖各资产明细的精确值。
            // 本程序也按同样口径重建净资产：现金用外汇持仓行中的数量和汇率精确计算，
            // 股票、债券、应计项目等使用报表明细给出的当前值。净资产总值表里的总数
            // 只用于校验重建结果按报表显示精度舍入后是否一致，不能作为余额或 record 的来源。
            var preciseCash = ParseIBKRPreciseCashValues(report, baseCurrency);
            var navSectionName = FindIBKRNavRow(report, IBKRNavTotalLabel).Section;
            decimal previousTotal = 0;
            decimal currentTotal = 0;
            foreach (var row in report.RequireDataRows(navSectionName))
            {
                if (row.Fields.Count == 1)
                    continue;
                AssertIBKRFieldCount(row, 6);
                var component = row.Fields[0];
                if (component.StartsWith(IBKRNavTotalLabel, StringComparison.Ordinal))
                    continue;
                if (IsIBKRNavDetailComponent(component))
                    continue;

                _ = ParseIBKRKnownNavComponent(component);
                var previous = ParseIBKRDecimalAt(row, 1, "precise NAV previous component");
                var current = ParseIBKRDecimalAt(row, 4, "precise NAV current component");
                if (component == IBKRNavCashComponent)
                {
                    AssertIBKRMoneyFieldEquals(preciseCash.Start, previous, row.Fields[1], "IBKR precise NAV cash previous display");
                    AssertIBKRMoneyFieldEquals(preciseCash.End, current, row.Fields[4], "IBKR precise NAV cash current display");
                    previous = preciseCash.Start;
                    current = preciseCash.End;
                }

                previousTotal += previous;
                currentTotal += current;
            }

            var totalRow = FindIBKRNavRow(report, IBKRNavTotalLabel);
            AssertIBKRMoneyFieldEquals(previousTotal, ParseIBKRDecimalAt(totalRow, 1, "precise NAV previous total"), totalRow.Fields[1], "IBKR precise NAV previous total display");
            AssertIBKRMoneyFieldEquals(currentTotal, ParseIBKRDecimalAt(totalRow, 4, "precise NAV current total"), totalRow.Fields[4], "IBKR precise NAV current total display");
            AssertIBKRMoneyFieldEquals(currentTotal - previousTotal, ParseIBKRDecimalAt(totalRow, 5, "precise NAV total change"), totalRow.Fields[5], "IBKR precise NAV total change display");
            return new IBKRPreciseNavValues(previousTotal, currentTotal, preciseCash.Start, preciseCash.End);
        }

        private static IBKRPreciseCashValues ParseIBKRPreciseCashValues(IBKRCsvReport report, CurrencyType baseCurrency)
        {
            decimal previousTotal = 0;
            decimal currentTotal = 0;
            var hasPositionCashRows = false;
            foreach (var row in report.Sections.Values.SelectMany(section => section.DataRows))
            {
                if (!IsIBKRPositionSummaryRow(row))
                    continue;
                if (!TryParseIBKRCurrencyType(row.Fields[3], out _))
                    continue;

                var positionValues = ParseIBKRFxPositionValues(row, baseCurrency, "IBKR precise FX");
                previousTotal += positionValues.PreviousMarketValue;
                currentTotal += positionValues.CurrentMarketValue;
                hasPositionCashRows = true;
            }

            if (!hasPositionCashRows)
                return new IBKRPreciseCashValues(ParseIBKRStartingCash(report), ParseIBKREndingCash(report));

            var startingCashRow = FindIBKRCashReportRow(report, IBKRStartingCashLabel);
            var endingCashRow = FindIBKRCashReportRow(report, IBKREndingCashLabel);
            AssertIBKRMoneyFieldEquals(previousTotal, ParseIBKRDecimalAt(startingCashRow, 2, "starting cash"), startingCashRow.Fields[2], "IBKR precise starting cash display");
            AssertIBKRMoneyFieldEquals(currentTotal, ParseIBKRDecimalAt(endingCashRow, 2, "ending cash"), endingCashRow.Fields[2], "IBKR precise ending cash display");
            return new IBKRPreciseCashValues(previousTotal, currentTotal);
        }

        private static decimal ParseIBKRNavChange(IBKRCsvReport report)
        {
            var (start, end) = ParseIBKRNavChangeValues(report);
            ValidateIBKRNavTotals(report, start, end);
            return end - start;
        }

        private static (decimal Start, decimal End) ParseIBKRNavChangeValues(IBKRCsvReport report)
        {
            var values = new Dictionary<string, decimal>(StringComparer.Ordinal);
            foreach (var row in report.RequireDataRows("净资产值变更"))
            {
                AssertIBKRFieldCount(row, 2);
                var label = row.Fields[0];
                if (values.ContainsKey(label))
                    throw new MailParseException($"Parse IBKR Report Fail, Duplicate NAV Change Row: {FormatIBKRCsvRow(row)}");

                values[label] = ParseIBKRDecimalAt(row, 1, $"NAV change {label}");
            }

            if (!values.TryGetValue("开始价值", out var start) || !values.TryGetValue("结束价值", out var end))
                throw new MailParseException("Parse IBKR Report Fail, Invalid NAV Change Table");

            decimal componentTotal = 0;
            foreach (var (label, amount) in values)
            {
                if (label is "开始价值" or "结束价值")
                    continue;
                if (label is "按市值计价"
                    or "存款和取款"
                    or "持仓转账"
                    or "股息"
                    or "代扣税款"
                    or "利息"
                    or "应计利息变更"
                    or "应计股息的变化"
                    or "其它外汇换算"
                    or "佣金")
                {
                    componentTotal += amount;
                    continue;
                }

                throw new MailParseException($"Parse IBKR Report Fail, Unknown NAV Change Component: {label}/{amount}");
            }

            return (start, end);
        }

        private static decimal ParseIBKRNavChangeComponent(IBKRCsvReport report, string component)
        {
            var row = report.RequireDataRows("净资产值变更")
                .FirstOrDefault(row => row.Fields.Count > 1 && row.Fields[0] == component);
            return row is null ? 0 : ParseIBKRDecimalAt(row, 1, $"NAV change component {component}");
        }

        private static void AddIBKROtherFxTranslationRecord(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency)
        {
            var amount = ParseIBKRNavChangeComponent(report, "其它外汇换算");
            builder.Add(
                new Currency(amount, baseCurrency),
                "其它外汇换算",
                $"NAVChange/其它外汇换算/{amount}",
                destAccount: baseCurrency.ToString());
        }

        private static void ValidateIBKRNavTotals(IBKRCsvReport report, decimal start, decimal end)
        {
            var totalRow = FindIBKRNavRow(report, "总数");
            AssertIBKRMoneyEquals(start, ParseIBKRDecimalAt(totalRow, 1, "NAV previous total"), "IBKR NAV previous total");
            AssertIBKRMoneyEquals(end, ParseIBKRDecimalAt(totalRow, 4, "NAV current total"), "IBKR NAV current total");
            decimal previousTotal = 0;
            decimal currentLongTotal = 0;
            decimal currentShortTotal = 0;
            decimal currentTotal = 0;
            decimal changeTotal = 0;
            foreach (var row in report.RequireDataRows("净资产值"))
            {
                if (row.Fields.Count == 1)
                    continue;
                AssertIBKRFieldCount(row, 6);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                    continue;
                if (IsIBKRNavDetailComponent(row.Fields[0]))
                    continue;

                _ = ParseIBKRKnownNavComponent(row.Fields[0]);
                previousTotal += ParseIBKRDecimalAt(row, 1, "NAV previous component");
                currentLongTotal += ParseIBKRDecimalAt(row, 2, "NAV current long component");
                currentShortTotal += ParseIBKRDecimalAt(row, 3, "NAV current short component");
                currentTotal += ParseIBKRDecimalAt(row, 4, "NAV current component");
                changeTotal += ParseIBKRDecimalAt(row, 5, "NAV change component");
            }

            AssertIBKRMoneyEquals(ParseIBKRDecimalAt(totalRow, 1, "NAV previous component sum"), previousTotal, "IBKR NAV previous component sum");
            AssertIBKRMoneyEquals(ParseIBKRDecimalAt(totalRow, 2, "NAV current long component sum"), currentLongTotal, "IBKR NAV current long component sum");
            AssertIBKRMoneyEquals(ParseIBKRDecimalAt(totalRow, 3, "NAV current short component sum"), currentShortTotal, "IBKR NAV current short component sum");
            AssertIBKRMoneyEquals(ParseIBKRDecimalAt(totalRow, 4, "NAV current component sum"), currentTotal, "IBKR NAV current component sum");
            AssertIBKRMoneyEquals(ParseIBKRDecimalAt(totalRow, 5, "NAV change component sum"), changeTotal, "IBKR NAV change component sum");
        }

        private static string ParseIBKRKnownNavComponent(string component)
        {
            return component switch
            {
                "现金" or "股票" or "债券" or "应计利息" or "应计股息" or "抵押品价值" or "借出证券" => component,
                _ => throw new MailParseException($"Parse IBKR Report Fail, Unknown NAV Component: {component}")
            };
        }

        private static bool IsIBKRNavDetailComponent(string component)
        {
            return component is "应计经纪商利息" or "应计债券利息";
        }

        private static decimal ParseIBKRNavComponent(IBKRCsvReport report, string component)
        {
            var row = report.RequireDataRows("净资产值")
                .FirstOrDefault(row => row.Fields.Count > 4 && row.Fields[0] == component);
            return row is null ? 0 : ParseIBKRDecimalAt(row, 4, $"NAV component {component}");
        }

        private static decimal ParseIBKRNavComponentChange(IBKRCsvReport report, string component)
        {
            var row = FindIBKRNavComponentRow(report, component);
            return row is null ? 0 : ParseIBKRDecimalAt(row, 5, $"NAV component change {component}");
        }

        private static IBKRCsvRow? FindIBKRNavComponentRow(IBKRCsvReport report, string component)
        {
            return report.RequireDataRows("净资产值")
                .FirstOrDefault(row => row.Fields.Count > 5 && row.Fields[0] == component);
        }

        private static IBKRCsvRow FindIBKRNavRow(IBKRCsvReport report, string label)
        {
            return report.RequireDataRows("净资产值")
                .FirstOrDefault(row => row.Fields.Count > 4 && row.Fields[0].StartsWith(label, StringComparison.Ordinal))
                ?? throw new MailParseException($"Parse IBKR Report Fail, Missing NAV Row: {label}");
        }

        private static IBKRCsvReport ReadIBKRCsvReport(string csv, string sourceName)
        {
            var report = new IBKRCsvReport(sourceName);
            using var reader = new StringReader(csv);
            string? line;
            var lineNumber = 0;
            while ((line = reader.ReadLine()) is not null)
            {
                lineNumber++;
                if (String.IsNullOrWhiteSpace(line))
                    continue;

                var fields = ParseIBKRCsvLine(line);
                if (fields.Count < 2)
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid CSV Row {sourceName}:{lineNumber}: {line}");

                var sectionName = NormalizeIBKRCsvSectionName(fields[0].Trim('\uFEFF', ' '));
                var rowType = fields[1].Trim();
                if (!IBKRCsvSections.Contains(sectionName))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown CSV Section: {sectionName}");
                if (sectionName == IBKRStatementPeriodPnlSection && String.IsNullOrWhiteSpace(rowType))
                    rowType = "Data";
                if (rowType is not ("Header" or "Headers" or "Data" or "Notes"))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown CSV Row Type: {rowType}/{sectionName}");

                var row = new IBKRCsvRow(
                    lineNumber,
                    sectionName,
                    rowType,
                    fields.Skip(2).Select(CleanIBKRCsvField).ToList());
                report.Add(row);
            }

            ValidateIBKRCsvHeaders(report);
            ValidateIBKRInformationalSections(report);
            return report;
        }

        private static string NormalizeIBKRCsvSectionName(string sectionName)
        {
            return sectionName == IBKRStockYieldEnhancementLoanSectionWithoutActivity
                ? IBKRStockYieldEnhancementLoanSection
                : sectionName;
        }

        private static void ValidateIBKRCsvHeaders(IBKRCsvReport report)
        {
            foreach (var (sectionName, section) in report.Sections)
            {
                if (IBKRFlexibleCsvHeaders.TryGetValue(sectionName, out var allowedHeaders))
                {
                    foreach (var headerRow in section.HeaderRows)
                    {
                        if (!allowedHeaders.Any(expectedHeader => headerRow.Fields.SequenceEqual(expectedHeader, StringComparer.Ordinal)))
                        {
                            throw new MailParseException(
                                $"Parse IBKR Report Fail, Header Mismatch: {sectionName}, actual={String.Join("|", headerRow.Fields)}");
                        }
                    }

                    continue;
                }

                var expectedHeaders = IBKRCsvHeaders[sectionName];
                if (expectedHeaders.Length == 0)
                {
                    if (section.HeaderRows.Count != 0)
                    {
                        throw new MailParseException(
                            $"Parse IBKR Report Fail, Header Count Mismatch: {sectionName}, expected=0, actual={section.HeaderRows.Count}");
                    }

                    continue;
                }

                if (section.HeaderRows.Count < expectedHeaders.Length
                    || section.HeaderRows.Count % expectedHeaders.Length != 0)
                {
                    throw new MailParseException(
                        $"Parse IBKR Report Fail, Header Count Mismatch: {sectionName}, expected={expectedHeaders.Length}, actual={section.HeaderRows.Count}");
                }

                for (var i = 0; i < section.HeaderRows.Count; i++)
                {
                    if (sectionName == "现金报告" && IsKnownIBKRCashReportHeaders(section.HeaderRows))
                        continue;

                    var expectedHeader = expectedHeaders[i % expectedHeaders.Length];
                    if (!section.HeaderRows[i].Fields.SequenceEqual(expectedHeader, StringComparer.Ordinal))
                    {
                        throw new MailParseException(
                            $"Parse IBKR Report Fail, Header Mismatch: {sectionName}, expected={String.Join("|", expectedHeader)}, actual={String.Join("|", section.HeaderRows[i].Fields)}");
                    }
                }
            }
        }

        private static bool IsKnownIBKRCashReportHeaders(List<IBKRCsvRow> headerRows)
        {
            return headerRows.Count == 1
                && (headerRows[0].Fields.SequenceEqual(IBKRCashReportFullHeader, StringComparer.Ordinal)
                    || headerRows[0].Fields.SequenceEqual(IBKRCashReportShortHeader, StringComparer.Ordinal));
        }

        private static void AssertIBKRCashReportFieldCount(IBKRCsvRow row)
        {
            if (row.Fields.Count is not (6 or 8))
                throw new MailParseException($"Parse IBKR Report Fail, Field Count Mismatch: {row.Section}, expected=6/8, actual={row.Fields.Count}, row={FormatIBKRCsvRow(row)}");
        }

        private static void ValidateIBKRInformationalSections(IBKRCsvReport report)
        {
            foreach (var row in report.OptionalDataRows("基础货币汇率"))
            {
                AssertIBKRFieldCount(row, 2);
                if (String.IsNullOrWhiteSpace(row.Fields[0]))
                    throw new MailParseException($"Parse IBKR Report Fail, Empty Base Currency Rate: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDecimalAt(row, 1, "base currency rate");
            }

            foreach (var row in report.OptionalDataRows("代码"))
            {
                if (row.Fields.Count is not (2 or 4))
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Code Legend Row: {FormatIBKRCsvRow(row)}");
            }

            ValidateIBKRStockYieldEnhancementLoanSection(report, IBKRStockYieldEnhancementLoanSection);
            ValidateIBKRStockYieldEnhancementLoanSection(report, IBKRStockYieldEnhancementCollateralHeldSection);
            ValidateIBKRStockYieldEnhancementLoanFeeDetailSection(report, IBKRStockYieldEnhancementLoanFeeSection);
            ValidateIBKRStockYieldEnhancementLoanFeeDetailSection(report, IBKRStockYieldEnhancementLoanRateSection);
            ValidateIBKRMoneyDetailSection(report, IBKRInterestSection, 4);
            ValidateIBKRMoneyDetailSection(report, IBKRBondInterestReceivedSection, 5, 3);
            ValidateIBKRMoneyDetailSection(report, IBKRBondInterestPaidSection, 5, 3);
            ValidateIBKRFlexibleMoneyDetailSection(report, IBKRDepositWithdrawSection, [4, 5], allowMultipleTotals: true);
            ValidateIBKRFlexibleMoneyDetailSection(report, IBKRDividendSection, [4, 5], allowMultipleTotals: true);
            ValidateIBKRMoneyDetailSection(report, IBKRDividendPaymentInLieuSection, 5, 3);
            ValidateIBKRMoneyDetailSection(report, IBKRWithholdingTaxSection, 5, 3);
            ValidateIBKRCommissionAdjustmentSection(report);
            ValidateIBKRCreditInterestDetails(report);
            ValidateIBKRStatementPeriodPnlSection(report);
        }

        private static void ValidateIBKRStockYieldEnhancementLoanSection(IBKRCsvReport report, string sectionName)
        {
            var rows = report.OptionalDataRows(sectionName).ToList();
            if (rows.Count == 0)
                return;

            decimal collateralTotal = 0;
            decimal reportedTotal = 0;
            var reportedTotalCount = 0;
            foreach (var row in rows)
            {
                if (row.Fields.Count == 7)
                {
                    ValidateIBKRStockYieldEnhancementLoanShortRow(row, sectionName, ref collateralTotal, ref reportedTotal, ref reportedTotalCount);
                    continue;
                }

                if (row.Fields.Count == 6 && sectionName == IBKRStockYieldEnhancementCollateralHeldSection)
                {
                    ValidateIBKRStockYieldEnhancementCollateralHeldRow(row, sectionName, ref collateralTotal, ref reportedTotal, ref reportedTotalCount);
                    continue;
                }

                AssertIBKRFieldCount(row, 9);
                if (row.Fields[0] == "总数")
                {
                    if (row.Fields.Take(8).Skip(1).Any(field => !String.IsNullOrWhiteSpace(field)))
                        throw new MailParseException($"Parse IBKR Report Fail, Invalid Stock Yield Enhancement Total: {sectionName}, {FormatIBKRCsvRow(row)}");

                    reportedTotal += ParseIBKRDecimalAt(row, 8, "stock yield enhancement collateral total");
                    reportedTotalCount++;
                    continue;
                }

                ValidateIBKRStockYieldEnhancementAssetClass(row, sectionName);
                _ = ParseIBKRCurrencyType(row.Fields[1]);
                if (String.IsNullOrWhiteSpace(row.Fields[2]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Code: {sectionName}, {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDate(row.Fields[3]);
                if (String.IsNullOrWhiteSpace(row.Fields[4]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Description: {sectionName}, {FormatIBKRCsvRow(row)}");
                if (String.IsNullOrWhiteSpace(row.Fields[6]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Transaction Number: {sectionName}, {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDecimalAt(row, 7, "stock yield enhancement quantity");
                collateralTotal += ParseIBKRDecimalAt(row, 8, "stock yield enhancement collateral");
            }

            if (reportedTotalCount == 0)
                throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Total: {sectionName}, {report.SourceName}");
            AssertIBKRMoneyEquals(reportedTotal, collateralTotal, "IBKR stock yield enhancement collateral total");

            foreach (var row in report.OptionalNoteRows(sectionName))
            {
                if (row.Fields.Count != 1 || String.IsNullOrWhiteSpace(row.Fields[0]))
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Stock Yield Enhancement Note: {sectionName}, {FormatIBKRCsvRow(row)}");
            }
        }

        private static void ValidateIBKRStockYieldEnhancementAssetClass(IBKRCsvRow row, string sectionName)
        {
            if (IBKRAssetGroups.Contains(row.Fields[0]) || row.Fields[0] == "证券")
                return;

            throw new MailParseException($"Parse IBKR Report Fail, Unknown Stock Yield Enhancement Asset Class: {sectionName}, {FormatIBKRCsvRow(row)}");
        }

        private static void ValidateIBKRStockYieldEnhancementLoanShortRow(
            IBKRCsvRow row,
            string sectionName,
            ref decimal collateralTotal,
            ref decimal reportedTotal,
            ref int reportedTotalCount)
        {
            if (row.Fields[0] == "总数")
            {
                if (row.Fields.Take(6).Skip(1).Any(field => !String.IsNullOrWhiteSpace(field)))
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Stock Yield Enhancement Total: {sectionName}, {FormatIBKRCsvRow(row)}");

                reportedTotal += ParseIBKRDecimalAt(row, 6, "stock yield enhancement collateral total");
                reportedTotalCount++;
                return;
            }

            ValidateIBKRStockYieldEnhancementAssetClass(row, sectionName);
            _ = ParseIBKRCurrencyType(row.Fields[1]);
            if (String.IsNullOrWhiteSpace(row.Fields[2]))
                throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Code: {sectionName}, {FormatIBKRCsvRow(row)}");
            if (String.IsNullOrWhiteSpace(row.Fields[3]))
                throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Transaction Number: {sectionName}, {FormatIBKRCsvRow(row)}");
            _ = ParseIBKRDecimalAt(row, 4, "stock yield enhancement quantity");
            _ = ParseIBKRDecimalAt(row, 5, "stock yield enhancement collateral rate");
            collateralTotal += ParseIBKRDecimalAt(row, 6, "stock yield enhancement collateral");
        }

        private static void ValidateIBKRStockYieldEnhancementCollateralHeldRow(
            IBKRCsvRow row,
            string sectionName,
            ref decimal collateralTotal,
            ref decimal reportedTotal,
            ref int reportedTotalCount)
        {
            if (row.Fields[0] == "总数")
            {
                if (row.Fields.Take(5).Skip(1).Any(field => !String.IsNullOrWhiteSpace(field)))
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Stock Yield Enhancement Total: {sectionName}, {FormatIBKRCsvRow(row)}");

                reportedTotal += ParseIBKRDecimalAt(row, 5, "stock yield enhancement collateral held total");
                reportedTotalCount++;
                return;
            }

            ValidateIBKRStockYieldEnhancementAssetClass(row, sectionName);
            _ = ParseIBKRCurrencyType(row.Fields[1]);
            if (String.IsNullOrWhiteSpace(row.Fields[2]))
                throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Code: {sectionName}, {FormatIBKRCsvRow(row)}");
            var quantity = ParseIBKRDecimalAt(row, 3, "stock yield enhancement collateral held quantity");
            var price = ParseIBKRDecimalAt(row, 4, "stock yield enhancement collateral held price");
            var value = ParseIBKRDecimalAt(row, 5, "stock yield enhancement collateral held value");
            var priceDivisor = row.Fields[0] == "证券" ? 100m : 1m;
            AssertIBKRValueMatchesDisplayedUnitPrice(
                quantity,
                price,
                row.Fields[4],
                value,
                row.Fields[5],
                priceDivisor,
                $"IBKR stock yield enhancement collateral held value {row.Fields[2]}");
            collateralTotal += value;
        }

        private static void ValidateIBKRStockYieldEnhancementLoanFeeDetailSection(IBKRCsvReport report, string sectionName)
        {
            var rows = report.OptionalDataRows(sectionName).ToList();
            if (rows.Count == 0)
                return;

            decimal total = 0;
            decimal? reportedTotal = null;
            string? reportedTotalText = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 10);
                if (row.Fields[0] == "总数")
                {
                    if (reportedTotal.HasValue)
                        throw new MailParseException($"Parse IBKR Report Fail, Duplicate {sectionName} Total: {FormatIBKRCsvRow(row)}");
                    if (row.Fields
                        .Select((field, index) => (field, index))
                        .Where(item => item.index != 0 && item.index != 8)
                        .Any(item => !String.IsNullOrWhiteSpace(item.field)))
                    {
                        throw new MailParseException($"Parse IBKR Report Fail, Invalid {sectionName} Total: {FormatIBKRCsvRow(row)}");
                    }

                    reportedTotal = ParseIBKRDecimalAt(row, 8, $"{sectionName} total");
                    reportedTotalText = row.Fields[8];
                    continue;
                }

                _ = ParseIBKRCurrencyType(row.Fields[0]);
                _ = ParseIBKRDate(row.Fields[1]);
                if (String.IsNullOrWhiteSpace(row.Fields[2]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing {sectionName} Code: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDate(row.Fields[3]);
                _ = ParseIBKRDecimalAt(row, 4, $"{sectionName} quantity");
                _ = ParseIBKRDecimalAt(row, 5, $"{sectionName} collateral");
                _ = ParseIBKRDecimalAt(row, 6, $"{sectionName} market rate");
                _ = ParseIBKRDecimalAt(row, 7, $"{sectionName} client rate");
                total += ParseIBKRDecimalAt(row, 8, sectionName);
            }

            if (!reportedTotal.HasValue)
                throw new MailParseException($"Parse IBKR Report Fail, Missing {sectionName} Total: {report.SourceName}");
            AssertIBKRMoneyFieldEquals(total, reportedTotal.Value, reportedTotalText!, $"IBKR {sectionName} total");
        }

        private static void ValidateIBKRFlexibleMoneyDetailSection(
            IBKRCsvReport report,
            string sectionName,
            int[] fieldCounts,
            bool allowMultipleTotals)
        {
            var rows = report.OptionalDataRows(sectionName).ToList();
            if (rows.Count == 0)
                return;

            var totalCount = 0;
            foreach (var row in rows)
            {
                if (!fieldCounts.Contains(row.Fields.Count))
                {
                    throw new MailParseException(
                        $"Parse IBKR Report Fail, Field Count Mismatch: {row.Section}:{row.LineNumber}, expected={String.Join("/", fieldCounts)}, actual={row.Fields.Count}, row={FormatIBKRCsvRow(row)}");
                }

                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    totalCount++;
                    if (!allowMultipleTotals && totalCount > 1)
                        throw new MailParseException($"Parse IBKR Report Fail, Duplicate {sectionName} Total: {FormatIBKRCsvRow(row)}");
                    _ = ParseIBKRDecimalAt(row, 3, $"{sectionName} total");
                    continue;
                }

                _ = ParseIBKRCurrencyType(row.Fields[0]);
                _ = ParseIBKRDate(row.Fields[1]);
                if (String.IsNullOrWhiteSpace(row.Fields[2]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing {sectionName} Description: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDecimalAt(row, 3, sectionName);
            }
        }

        private static void ValidateIBKRCommissionAdjustmentSection(IBKRCsvReport report)
        {
            ValidateIBKRMoneyDetailSection(report, IBKRCommissionAdjustmentSection, 5, 3);
        }

        private static void ValidateIBKRCreditInterestDetails(IBKRCsvReport report)
        {
            var rows = report.OptionalDataRows(IBKRCreditInterestSection).ToList();
            if (rows.Count == 0)
                return;

            decimal total = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 11);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    if (reportedTotal.HasValue)
                        throw new MailParseException($"Parse IBKR Report Fail, Duplicate Credit Interest Total: {FormatIBKRCsvRow(row)}");
                    _ = ParseIBKRDecimalAt(row, 3, "credit interest total rate");
                    _ = ParseIBKRDecimalAt(row, 6, "credit interest total principal");
                    _ = ParseIBKRDecimalAt(row, 7, "credit interest total securities interest");
                    _ = ParseIBKRDecimalAt(row, 8, "credit interest total futures interest");
                    reportedTotal = ParseIBKRDecimalAt(row, 9, "credit interest total");
                    continue;
                }

                _ = ParseIBKRCurrencyType(row.Fields[0]);
                _ = ParseIBKRDate(row.Fields[1]);
                if (String.IsNullOrWhiteSpace(row.Fields[2]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Credit Interest Tier: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDecimalAt(row, 3, "credit interest rate");
                _ = ParseIBKRDecimalAt(row, 4, "credit interest securities principal");
                _ = ParseIBKRDecimalAt(row, 5, "credit interest futures principal");
                _ = ParseIBKRDecimalAt(row, 6, "credit interest total principal");
                _ = ParseIBKRDecimalAt(row, 7, "credit interest securities interest");
                _ = ParseIBKRDecimalAt(row, 8, "credit interest futures interest");
                total += ParseIBKRDecimalAt(row, 9, "credit interest");
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, total, "IBKR credit interest total");
        }

        private static void ValidateIBKRStatementPeriodPnlSection(IBKRCsvReport report)
        {
            foreach (var row in report.OptionalDataRows(IBKRStatementPeriodPnlSection))
            {
                AssertIBKRFieldCount(row, 12);
                _ = ParseIBKRDecimalAt(row, 10, "statement period P/L");
            }
        }

        private static decimal? ValidateIBKRMoneyDetailSection(
            IBKRCsvReport report,
            string sectionName,
            int fieldCount,
            int? totalAmountIndex = null)
        {
            var rows = report.OptionalDataRows(sectionName).ToList();
            if (rows.Count == 0)
                return null;

            var totalIndex = totalAmountIndex ?? fieldCount - 1;
            decimal total = 0;
            decimal? reportedTotal = null;
            string? reportedTotalText = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, fieldCount);
                if (row.Fields[0] == "总数")
                {
                    if (reportedTotal.HasValue)
                        throw new MailParseException($"Parse IBKR Report Fail, Duplicate {sectionName} Total: {FormatIBKRCsvRow(row)}");
                    if (row.Fields
                        .Select((field, index) => (field, index))
                        .Where(item => item.index != 0 && item.index != totalIndex)
                        .Any(item => !String.IsNullOrWhiteSpace(item.field)))
                    {
                        throw new MailParseException($"Parse IBKR Report Fail, Invalid {sectionName} Total: {FormatIBKRCsvRow(row)}");
                    }

                    reportedTotal = ParseIBKRDecimalAt(row, totalIndex, $"{sectionName} total");
                    reportedTotalText = row.Fields[totalIndex];
                    continue;
                }

                _ = ParseIBKRCurrencyType(row.Fields[0]);
                _ = ParseIBKRDate(row.Fields[1]);
                if (String.IsNullOrWhiteSpace(row.Fields[2]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing {sectionName} Description: {FormatIBKRCsvRow(row)}");
                total += ParseIBKRDecimalAt(row, 3, sectionName);
            }

            if (!reportedTotal.HasValue)
                throw new MailParseException($"Parse IBKR Report Fail, Missing {sectionName} Total: {report.SourceName}");
            AssertIBKRMoneyFieldEquals(total, reportedTotal.Value, reportedTotalText!, $"IBKR {sectionName} total");
            return total;
        }

        private static List<string> ParseIBKRCsvLine(string line)
        {
            var fields = new List<string>();
            var current = new StringBuilder();
            var inQuotes = false;
            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            current.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        fields.Add(current.ToString());
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }

            if (inQuotes)
                throw new MailParseException($"Parse IBKR Report Fail, Unterminated CSV Quote: {line}");

            fields.Add(current.ToString());
            return fields;
        }

        private static string CleanIBKRCsvField(string text)
        {
            return text.Replace('\u00a0', ' ').Trim();
        }

        private static void AssertIBKRFieldCount(IBKRCsvRow row, int count)
        {
            if (row.Fields.Count != count)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, Field Count Mismatch: {row.Section}:{row.LineNumber}, expected={count}, actual={row.Fields.Count}, row={FormatIBKRCsvRow(row)}");
            }
        }

        private static string FormatIBKRCsvRow(IBKRCsvRow row)
        {
            return $"{row.Section},{row.RowType},{String.Join(",", row.Fields)}";
        }

        private static int ParseIBKRIntegerQuantityAt(IBKRCsvRow row, int index, string context)
        {
            var value = ParseIBKRDecimalAt(row, index, context);
            if (value != Decimal.Truncate(value))
                throw new MailParseException($"Parse IBKR Report Fail, Non-integer {context}: {FormatIBKRCsvRow(row)}");

            return checked((int)value);
        }

        private static int ParseIBKRIntegerQuantityOrZero(string text, string context, IBKRCsvRow row)
        {
            if (String.IsNullOrWhiteSpace(text))
                return 0;
            if (!TryParseIBKRDecimal(text, out var value) || value != Decimal.Truncate(value))
                throw new MailParseException($"Parse IBKR Report Fail, Non-integer {context}: {FormatIBKRCsvRow(row)}");
            return checked((int)value);
        }

        private static decimal ParseIBKRDecimalAt(IBKRCsvRow row, int index, string context)
        {
            if (index >= row.Fields.Count || !TryParseIBKRDecimal(row.Fields[index], out var value))
                throw new MailParseException($"Parse IBKR Report Fail, Invalid {context}: {FormatIBKRCsvRow(row)}");
            return value;
        }

        private static bool TryParseIBKRDecimal(string text, out decimal value)
        {
            value = 0;
            var normalized = text.Trim().Replace(",", "");
            if (String.IsNullOrWhiteSpace(normalized))
                return false;
            if (normalized.EndsWith("%", StringComparison.Ordinal))
                normalized = normalized[..^1];
            return Decimal.TryParse(
                normalized,
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture,
                out value);
        }

        private static decimal ParseIBKRDecimalOrZero(string text)
        {
            if (String.IsNullOrWhiteSpace(text) || text.Trim() == "-")
                return 0;
            if (TryParseIBKRDecimal(text, out var value))
                return value;
            throw new MailParseException($"Parse IBKR Report Fail, Invalid Decimal: {text}");
        }

        private static CurrencyType ParseIBKRCurrencyType(string text)
        {
            if (TryParseIBKRCurrencyType(text, out var currencyType))
                return currencyType;
            throw new MailParseException($"Parse IBKR Report Fail, Unknown Currency: {text}");
        }

        private static bool TryParseIBKRCurrencyType(string text, out CurrencyType currencyType)
        {
            return Enum.TryParse(text.Trim(), true, out currencyType);
        }

        private static DateTime ParseIBKRDate(string text)
        {
            if (TryParseIBKRDate(text, out var date))
                return date;
            throw new MailParseException($"Parse IBKR Report Fail, Invalid Date: {text}");
        }

        private static bool TryParseIBKRDate(string text, out DateTime date)
        {
            return DateTime.TryParseExact(text.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        private static DateTime ParseIBKRDateTime(string text)
        {
            if (DateTime.TryParseExact(text.Trim(), "yyyy-MM-dd, HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                return date;

            return ParseIBKRDate(text);
        }

        private static IBKRStatementPeriod ParseIBKRStatementPeriod(string text)
        {
            var parts = Regex.Split(text.Trim(), @"\s+-\s+");
            if (parts.Length == 1)
            {
                var date = ParseIBKRStatementPeriodDate(parts[0]);
                return new IBKRStatementPeriod(date, date);
            }

            if (parts.Length == 2)
            {
                var start = ParseIBKRStatementPeriodDate(parts[0]);
                var end = ParseIBKRStatementPeriodDate(parts[1]);
                if (start.Date > end.Date)
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Statement Period Range: {text}");
                return new IBKRStatementPeriod(start, end);
            }

            throw new MailParseException($"Parse IBKR Report Fail, Invalid Statement Period: {text}");
        }

        private static DateTime ParseIBKRStatementPeriodDate(string text)
        {
            var match = Regex.Match(text.Trim(), @"^(?<month>\S+)\s+(?<day>\d{1,2}),\s*(?<year>\d{4})$");
            if (!match.Success)
                throw new MailParseException($"Parse IBKR Report Fail, Invalid Statement Period: {text}");

            var monthName = match.Groups["month"].Value;
            var month = monthName switch
            {
                "一月" or "January" => 1,
                "二月" or "February" => 2,
                "三月" or "March" => 3,
                "四月" or "April" => 4,
                "五月" or "May" => 5,
                "六月" or "June" => 6,
                "七月" or "July" => 7,
                "八月" or "August" => 8,
                "九月" or "September" => 9,
                "十月" or "October" => 10,
                "十一月" or "November" => 11,
                "十二月" or "December" => 12,
                _ => throw new MailParseException($"Parse IBKR Report Fail, Unknown Statement Month: {text}")
            };
            return new DateTime(
                Int32.Parse(match.Groups["year"].Value, CultureInfo.InvariantCulture),
                month,
                Int32.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture));
        }

        private static void AssertIBKRMoneyEquals(decimal expected, decimal actual, string context)
        {
            if (expected != actual)
                throw new MailParseException($"Parse IBKR Report Fail, {context} mismatch: expected {expected}, got {actual}");
        }

        private static void AssertIBKRMoneyWithin(decimal expected, decimal actual, decimal tolerance, string context)
        {
            if (Math.Abs(expected - actual) > tolerance)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, {context} mismatch: expected {expected}, got {actual}, tolerance {tolerance}");
            }
        }

        private static void AssertIBKRMoneyDisplayEquals(decimal preciseValue, decimal displayedValue, string context)
        {
            var roundedPreciseValue = RoundIBKRMoneyForReportDisplay(preciseValue);
            if (roundedPreciseValue != displayedValue)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, {context} mismatch: expected display {roundedPreciseValue} from precise {preciseValue}, got {displayedValue}");
            }
        }

        private static void AssertIBKRMoneyDisplayAlmostEquals(decimal preciseValue, decimal displayedValue, string context)
        {
            var roundedPreciseValue = RoundIBKRMoneyForReportDisplay(preciseValue);
            if (roundedPreciseValue != displayedValue && Math.Abs(roundedPreciseValue - displayedValue) > 0.00000001m)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, {context} mismatch: expected display {roundedPreciseValue} from precise {preciseValue}, got {displayedValue}");
            }
        }

        private static void AssertIBKRMoneyFieldEquals(
            decimal preciseValue,
            decimal displayedValue,
            string reportText,
            string context,
            MidpointRounding rounding = MidpointRounding.AwayFromZero)
        {
            var roundedPreciseValue = RoundIBKRMoneyForReportField(preciseValue, reportText, rounding);
            if (roundedPreciseValue != displayedValue)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, {context} mismatch: expected display {roundedPreciseValue} from precise {preciseValue}, got {displayedValue}");
            }
        }

        private static void AssertIBKRMoneyFieldAlmostEquals(
            decimal preciseValue,
            decimal displayedValue,
            string reportText,
            string context,
            MidpointRounding rounding = MidpointRounding.AwayFromZero)
        {
            var roundedPreciseValue = RoundIBKRMoneyForReportField(preciseValue, reportText, rounding);
            if (roundedPreciseValue == displayedValue)
                return;

            var tolerance = GetIBKRReportFieldUnit(reportText);
            if (Math.Abs(roundedPreciseValue - displayedValue) > tolerance)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, {context} mismatch: expected display {roundedPreciseValue} from precise {preciseValue}, got {displayedValue}");
            }
        }

        private static void AssertIBKRValueMatchesDisplayedUnitPrice(
            decimal quantity,
            decimal displayedUnitPrice,
            string unitPriceText,
            decimal displayedValue,
            string valueText,
            decimal priceDivisor,
            string context)
        {
            var expectedValue = quantity * displayedUnitPrice / priceDivisor;
            var unitPriceTolerance = Math.Abs(quantity) * GetIBKRReportFieldUnit(unitPriceText) / priceDivisor / 2m;
            var valueTolerance = GetIBKRReportFieldUnit(valueText);
            var tolerance = unitPriceTolerance + valueTolerance;
            if (Math.Abs(expectedValue - displayedValue) > tolerance)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, {context} mismatch: expected {expectedValue} from displayed unit price {displayedUnitPrice}, got {displayedValue}, tolerance {tolerance}");
            }
        }

        private static decimal GetIBKRReportFieldUnit(string reportText)
        {
            var text = reportText.Trim();
            var decimalPoint = text.IndexOf('.');
            if (decimalPoint < 0)
                return 1m;

            var decimalPlaces = text.Length - decimalPoint - 1;
            var unit = 1m;
            for (var i = 0; i < decimalPlaces; i++)
                unit /= 10m;
            return unit;
        }

        private static decimal RoundIBKRMoneyForReportDisplay(decimal value)
        {
            return Math.Round(value, 8, MidpointRounding.AwayFromZero);
        }

        private static decimal RoundIBKRMoneyForReportField(
            decimal value,
            string reportText,
            MidpointRounding rounding = MidpointRounding.AwayFromZero)
        {
            var text = reportText.Trim();
            var decimalPoint = text.IndexOf('.');
            if (decimalPoint < 0)
                return Math.Round(value, 0, rounding);

            var decimalPlaces = text.Length - decimalPoint - 1;
            return Math.Round(value, decimalPlaces, rounding);
        }

        private static string LimitIBKRRecordText(string text)
        {
            if (text.Length <= 1024)
                return text;
            return text[..1024];
        }

        private static string? ParseIBKRAccountId(string text)
        {
            var normalizedText = NormalizeMailText(text);
            var labeledMatch = Regex.Match(
                normalizedText,
                @"(?:Account(?:\s*ID|\s*Number)?|账号|账户)\s*[:：]?\s*(?<account>U(?:\*+|\d*)\d{4,})",
                RegexOptions.IgnoreCase);
            if (labeledMatch.Success)
                return labeledMatch.Groups["account"].Value.ToUpperInvariant();

            var match = Regex.Match(normalizedText, @"\b(?<account>U(?:\*{2,}\d{4,}|\d{5,}))\b", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups["account"].Value.ToUpperInvariant() : null;
        }

        private static List<InMemoryIBKRReportAttachment> ReadIBKRReportAttachments(
            MimeMessage message,
            DateTime reportDate)
        {
            return ReadMatchingAttachments(message, (attachment, fileName) =>
            {
                if (!TryParseIBKRReportAttachmentName(fileName, out var attachmentInfo)
                    || !IsDailyMyBookReportType(attachmentInfo.ReportType)
                    || attachmentInfo.ReportDate.Date != reportDate.Date
                    || !IsIBKRCsvAttachment(fileName))
                    return null;

                return new InMemoryIBKRReportAttachment(
                    fileName,
                    attachmentInfo.ReportType,
                    attachmentInfo.ReportId,
                    attachmentInfo.ReportDate,
                    ReadMimePartBytes(attachment));
            });
        }

        private static bool HasIBKRReportAttachment(MimeMessage message, DateTime reportDate)
        {
            return HasMatchingAttachment(message, (_, fileName) =>
            {
                return TryParseIBKRReportAttachmentName(fileName, out var attachmentInfo)
                    && IsDailyMyBookReportType(attachmentInfo.ReportType)
                    && attachmentInfo.ReportDate.Date == reportDate.Date
                    && IsIBKRCsvAttachment(fileName);
            });
        }

        private static bool HasIBKRReportAttachment(IMessageSummary summary, DateTime reportDate)
        {
            return SummaryHasMatchingAttachment(summary, fileName =>
            {
                return TryParseIBKRReportAttachmentName(fileName, out var attachmentInfo)
                    && IsDailyMyBookReportType(attachmentInfo.ReportType)
                    && attachmentInfo.ReportDate.Date == reportDate.Date
                    && IsIBKRCsvAttachment(fileName);
            });
        }

        private static bool IsDailyMyBookReportType(string reportType)
        {
            return String.Equals(IBKRDailyReportType, reportType, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsIBKRCsvAttachment(string fileName)
        {
            return Path.GetExtension(fileName).Equals(".csv", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryParseIBKRReportAttachmentName(string fileName, out IBKRReportAttachmentInfo attachmentInfo)
        {
            attachmentInfo = new IBKRReportAttachmentInfo("", "", default);
            var name = Path.GetFileName(fileName.Trim());
            var match = Regex.Match(name, @"^(?<type>[^.]+)\.(?<id>[^.]+)\.(?<date>\d{8})(?:\..*)?$");
            if (!match.Success)
                return false;

            if (!DateTime.TryParseExact(
                match.Groups["date"].Value,
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var reportDate))
            {
                return false;
            }

            attachmentInfo = new IBKRReportAttachmentInfo(
                match.Groups["type"].Value,
                match.Groups["id"].Value,
                reportDate);
            return true;
        }

        private sealed record IBKRParsedReport(
            DateTime ReportDate,
            DateTime ImportDate,
            string StatementKey,
            string AccountId,
            Account Account,
            Records Records,
            List<Holding> BeginningHoldings,
            List<Holding> Holdings,
            List<AccountBalance> AccountBalances,
            List<AccountBalance> BeginningAccountBalances,
            List<AccountInternalId> InternalCardNos,
            IBKRStatementPeriod Period,
            bool IsInitialReport);

        private sealed record IBKRReportSaveItem(IBKRParsedReport Report, bool ReturnResult);

        private sealed record IBKRStatementPeriod(DateTime StartDate, DateTime EndDate);

        private sealed record IBKRHoldingSnapshotValue(int Quantity, Currency TotalPrice);

        private sealed record IBKRContractInfo(string Code, string Description, HoldingType HoldingType, string DisplayText);

        private sealed record IBKRStockBondPositionRow(
            IBKRCsvRow Row,
            IBKRContractInfo Contract,
            CurrencyType Currency,
            int BeginningQuantity,
            int EndingQuantity);

        private sealed record IBKRPreciseNavValues(decimal Start, decimal End, decimal StartingCash, decimal EndingCash);

        private sealed record IBKRPreciseCashValues(decimal Start, decimal End);

        private sealed record IBKRFxPositionValues(
            CurrencyType CashCurrency,
            decimal PreviousQuantity,
            decimal CurrentQuantity,
            decimal PreviousMarketValue,
            decimal CurrentMarketValue,
            decimal PriorPositionRateImpact);

        private sealed record IBKRMtmTotals(
            bool HasData,
            decimal Holding,
            decimal Transaction,
            decimal Commission,
            decimal Other,
            decimal CashFxTranslation,
            decimal FxTransaction,
            decimal FxTransactionDisplayUnit);

        private sealed record IBKRPositionQuantityChangeTotals(
            Dictionary<string, decimal> CoveredTransactionImpacts,
            decimal HoldingValueChange,
            decimal CashChange,
            decimal TransactionPriceImpact);

        private sealed record IBKRTransactionTotals(
            bool HasData,
            decimal Proceeds,
            decimal BuyProceeds,
            decimal SellProceeds,
            decimal StockBondProceeds,
            decimal Commission,
            int Quantity,
            bool HasFxTrades,
            Dictionary<string, IBKRTradeSummaryData> Holdings);

        private sealed record IBKRTradeSummaryData(
            IBKRContractInfo Contract,
            CurrencyType Currency,
            int BuyQuantity,
            int SellQuantity,
            decimal BuyProceeds,
            decimal SellProceeds,
            string Source)
        {
            public int Quantity => BuyQuantity + SellQuantity;
            public decimal Proceeds => BuyProceeds + SellProceeds;

            public IBKRTradeSummaryData Add(
                CurrencyType currency,
                int buyQuantity,
                int sellQuantity,
                decimal buyProceeds,
                decimal sellProceeds,
                string source)
            {
                if (Currency != currency)
                    throw new MailParseException($"Parse IBKR Report Fail, Trade Summary Currency Mismatch: {Contract.Code}");

                return this with
                {
                    BuyQuantity = BuyQuantity + buyQuantity,
                    SellQuantity = SellQuantity + sellQuantity,
                    BuyProceeds = BuyProceeds + buyProceeds,
                    SellProceeds = SellProceeds + sellProceeds,
                    Source = $"{Source}; {source}"
                };
            }
        }

        private sealed record IBKRCommissionTotals(bool HasData, decimal Total);

        private sealed record IBKRInterestTotals(bool HasData, decimal Total);

        private sealed record IBKRCashTotals(
            decimal Commission,
            decimal TradeBuy,
            decimal TradeSell,
            decimal NonTradeCashFlow,
            decimal CashFxTranslation);

        private sealed record IBKRTransferTotals(bool HasData, decimal Total);

        private sealed record IBKRReportAttachmentInfo(string ReportType, string ReportId, DateTime ReportDate);

        private sealed record InMemoryIBKRReportAttachment(
            string FileName,
            string ReportType,
            string ReportId,
            DateTime ReportDate,
            byte[] Content);

        private sealed record IBKRCsvRow(int LineNumber, string Section, string RowType, List<string> Fields);

        private sealed class IBKRCsvSection
        {
            public List<IBKRCsvRow> HeaderRows { get; } = [];
            public List<IBKRCsvRow> DataRows { get; } = [];
            public List<IBKRCsvRow> NoteRows { get; } = [];
        }

        private sealed class IBKRCsvReport
        {
            public IBKRCsvReport(string sourceName)
            {
                SourceName = sourceName;
            }

            public string SourceName { get; }
            public Dictionary<string, IBKRCsvSection> Sections { get; } = new(StringComparer.Ordinal);

            public void Add(IBKRCsvRow row)
            {
                if (!Sections.TryGetValue(row.Section, out var section))
                {
                    section = new IBKRCsvSection();
                    Sections[row.Section] = section;
                }

                if (row.RowType is "Header" or "Headers")
                    section.HeaderRows.Add(row);
                else if (row.RowType == "Notes")
                    section.NoteRows.Add(row);
                else
                    section.DataRows.Add(row);
            }

            public List<IBKRCsvRow> RequireDataRows(string sectionName)
            {
                if (!Sections.TryGetValue(sectionName, out var section))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing CSV Section: {sectionName}");
                return section.DataRows;
            }

            public IEnumerable<IBKRCsvRow> OptionalDataRows(string sectionName)
            {
                return Sections.TryGetValue(sectionName, out var section)
                    ? section.DataRows
                    : Enumerable.Empty<IBKRCsvRow>();
            }

            public IEnumerable<IBKRCsvRow> OptionalNoteRows(string sectionName)
            {
                return Sections.TryGetValue(sectionName, out var section)
                    ? section.NoteRows
                    : Enumerable.Empty<IBKRCsvRow>();
            }
        }

        private sealed class IBKRRecordBuilder
        {
            private readonly Account account;
            private readonly DateTime reportDate;
            private readonly string sourceName;
            private readonly CurrencyType baseCurrency;

            public IBKRRecordBuilder(
                Account account,
                DateTime reportDate,
                string sourceName,
                CurrencyType baseCurrency,
                bool allowLargeStatementResidual)
            {
                this.account = account;
                this.reportDate = reportDate;
                this.sourceName = sourceName;
                this.baseCurrency = baseCurrency;
                AllowLargeStatementResidual = allowLargeStatementResidual;
            }

            public Records Records { get; } = new();
            public decimal NetAssetChangeTotal { get; private set; }
            public bool AllowLargeStatementResidual { get; }

            public string DescribeNetAssetChangeTotals()
            {
                return String.Join(
                    "; ",
                    Records
                        .GroupBy(record => $"{(record.isInternal ? "internal " : "")}{record.Reason}", StringComparer.Ordinal)
                        .Select(group => $"{group.Key}={group.Sum(record => record.v)}")
                        .OrderBy(text => text, StringComparer.Ordinal));
            }

            public void Add(
                Currency amount,
                string reason,
                string source,
                bool isInternal = false,
                bool affectsNetAsset = true,
                DateTime? date = null,
                string destAccount = "",
                int holdingQuantity = 0,
                IBKRContractInfo? holding = null)
            {
                if (amount.v == 0)
                    return;
                if (affectsNetAsset && amount.t != baseCurrency)
                    throw new MailParseException($"Parse IBKR Report Fail, Non-base NAV record: {reason} {amount.v}/{amount.t}");

                var record = new Record
                {
                    Account = account,
                    date = date ?? reportDate,
                    updateTime = DateTime.Now,
                    DestAccount = destAccount,
                    isInternal = isInternal,
                    HoldingQuantity = holdingQuantity,
                    Holding = holding is null
                        ? null
                        : new Holding(holding.Code, holding.HoldingType)
                        {
                            Account = account,
                            desc = holding.Description,
                            displayText = holding.DisplayText
                        },
                    Source = LimitIBKRRecordText($"IBKR CSV {sourceName}/{source}"),
                    Reason = LimitIBKRRecordText(reason)
                };
                record.CopyFrom(amount);
                Records.Add(record);

                if (affectsNetAsset)
                    NetAssetChangeTotal += amount.v;
            }
        }
    }
}
