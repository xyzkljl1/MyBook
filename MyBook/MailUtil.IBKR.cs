using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MimeKit;

namespace MyBook
{
    // IBKR DailyMyBook CSV report discovery, parsing, and account matching.
    partial class MailUtil
    {
        private const StatementImportProvider IBKRProvider = StatementImportProvider.IBKRReportMail;
        private const string IBKRReportSender = "donotreply@interactivebrokers.com";
        private const string IBKRDailyReportType = "DailyMyBook";
        private const int IBKRMissingReportLimitDays = 14;
        private const decimal IBKRPrecisionResidualLimit = 0.0000001m;
        private const string IBKRStockYieldEnhancementLoanSection = "股票收益提升计划股证券出借活动";
        private const string IBKRInterestSection = "利息";
        private const string IBKRBondInterestReceivedSection = "收到的债券利息";
        private const string IBKRDividendAccrualChangeSection = "应计股息的变化";
        private const string IBKRTransferSection = "转账";

        private static readonly HashSet<string> IBKRAssetGroups = new(StringComparer.Ordinal)
        {
            "股票",
            "债券",
            "外汇"
        };

        private static readonly string[] IBKRCashReportFullHeader = ["货币总结", "货币", "总数", "证券", "期货", "本月截至当前", "本年截至当前", ""];
        private static readonly string[] IBKRCashReportShortHeader = ["货币总结", "货币", "总数", "证券", "期货", ""];

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
            IBKRInterestSection,
            IBKRBondInterestReceivedSection
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
                ["资产分类", "货币", "代码", "日期", "描述", "", "交易号码", "数量", "抵押品金额"]
            ],
            [IBKRInterestSection] =
            [
                ["货币", "日期", "描述", "金额"]
            ],
            [IBKRBondInterestReceivedSection] =
            [
                ["货币", "日期", "描述", "金额", "代码"]
            ]
        };

        private static readonly Dictionary<string, HoldingType> IBKRFallbackHoldingTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            ["DIA"] = HoldingType.ARCA,
            ["NVDL"] = HoldingType.NASDAQ,
            ["NVDA"] = HoldingType.NASDAQ,
            ["QQQ"] = HoldingType.NASDAQ,
            ["SPY"] = HoldingType.ARCA,
            ["TLT"] = HoldingType.NASDAQ,
            ["TQQQ"] = HoldingType.NASDAQ,
        };

        public void DebugFetchLocalIBKRReports()
        {
            var files = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.csv")
                .Where(file => TryParseIBKRReportAttachmentName(Path.GetFileName(file), out var attachmentInfo)
                    && IsDailyMyBookReportType(attachmentInfo.ReportType))
                .OrderBy(file =>
                {
                    TryParseIBKRReportAttachmentName(Path.GetFileName(file), out var attachmentInfo);
                    return attachmentInfo.ReportDate;
                })
                .ToList();
            if (files.Count == 0)
                throw new FileNotFoundException("No local IBKR csv report found.");

            var reports = files
                .Select(file =>
                {
                    TryParseIBKRReportAttachmentName(Path.GetFileName(file), out var attachmentInfo);
                    var csv = ReadIBKRLocalCsv(file);
                    return ParseIBKRReportCsv(csv, attachmentInfo.ReportDate, file);
                })
                .ToList();
            SaveIBKRParsedReports(reports);
        }

        public bool DebugFetchLocalIBKRReport(string path)
        {
            if (!TryParseIBKRReportAttachmentName(Path.GetFileName(path), out var attachmentInfo)
                || !IsDailyMyBookReportType(attachmentInfo.ReportType)
                || !IsIBKRCsvAttachment(path))
            {
                throw new ArgumentException($"Invalid IBKR csv report file name: {path}");
            }

            var csv = ReadIBKRLocalCsv(path);
            return SaveIBKRParsedReports([ParseIBKRReportCsv(csv, attachmentInfo.ReportDate, path)])[0];
        }

        private static string ReadIBKRLocalCsv(string path)
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }

        public async Task FetchIBKRReports()
        {
            var date = GetNextDailyStatementDate(IBKRProvider);
            Console.WriteLine($"Fetch IBKR reports from {date:yyyy-MM-dd}");
            var missingDays = 0;
            while (date <= DateTime.Today)
            {
                Console.WriteLine($"Fetch IBKR report {date:yyyy-MM-dd}");
                var imported = await FetchIBKRReport(date);
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
        }

        private async Task<bool> FetchIBKRReport(DateTime date)
        {
            var subjectDate = date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            var expectedSubject = $"{subjectDate}的自定义活动报表";
            var messages = await SearchBills(
                IBKRReportSender,
                expectedSubject,
                date,
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
            var saved = database.SaveStatementRecordsAndHoldingsOnce(reports.Select(report =>
                new StatementRecordHoldingImport(
                    IBKRProvider,
                    report.ImportDate,
                    report.StatementKey,
                    report.Account,
                    report.Records,
                    report.Holdings,
                    report.AccountBalances,
                    report.BeginningAccountBalances,
                    report.InternalCardNos)));

            for (var i = 0; i < reports.Count; i++)
            {
                var report = reports[i];
                Console.WriteLine(
                    saved[i]
                        ? $"Import IBKR report: {report.ReportDate:yyyy-MM-dd} {report.AccountId} -> {report.Account.name}, records={report.Records.Count}, holdings={report.Holdings.Count}"
                        : $"Skip imported IBKR report: {report.ReportDate:yyyy-MM-dd} {report.AccountId} -> {report.Account.name}");
            }

            return saved;
        }

        private IBKRParsedReport ParseIBKRReportCsv(string csv, DateTime reportDate, string sourceName, DateTime? importDate = null)
        {
            if (String.IsNullOrWhiteSpace(csv))
                throw new MailParseException("Parse IBKR Report Fail, Empty CSV");

            var report = ReadIBKRCsvReport(csv, sourceName);
            ValidateIBKRStatementMetadata(report, reportDate);
            var (accountId, baseCurrency) = ParseIBKRAccountInfo(report);
            var account = GetIBKRAccount(accountId);
            var contractInfos = BuildIBKRContractInfos(report);
            var holdings = ParseIBKRHoldings(report, account, baseCurrency, contractInfos, sourceName);
            var startingNav = ParseIBKRStartingNav(report);
            var endingNav = ParseIBKREndingNav(report);
            var navChange = ParseIBKRNavChange(report);
            var records = ParseIBKRRecords(report, account, baseCurrency, contractInfos, reportDate.Date, sourceName, navChange);
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
                (importDate ?? reportDate).Date,
                BuildIBKRStatementKey(account, reportDate),
                accountId,
                account,
                records,
                holdings,
                [new AccountBalance(account, new Currency(endingNav, baseCurrency))],
                [new AccountBalance(account, new Currency(startingNav, baseCurrency))],
                internalCardNos);
        }

        private static string BuildIBKRStatementKey(Account account, DateTime reportDate)
        {
            const string prefix = "IBKR_";
            var accountId = account.name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? account.name[prefix.Length..]
                : account.name;
            return $"{accountId}_{reportDate:yyyy-MM-dd}";
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

        private static void ValidateIBKRStatementMetadata(IBKRCsvReport report, DateTime reportDate)
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
            if (ParseIBKRStatementPeriod(period).Date != reportDate.Date)
                throw new MailParseException($"Parse IBKR Report Fail, Statement Period Mismatch: {period}/{reportDate:yyyy-MM-dd}");
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
                if (group != "股票" && group != "债券")
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
            decimal expectedNavChange)
        {
            var builder = new IBKRRecordBuilder(account, reportDate, sourceName, baseCurrency);
            var mtmTotals = ParseIBKRMtmRecords(report, builder, baseCurrency, contractInfos);
            var tradeTotals = ParseIBKRTradeSummary(report);
            var commissionTotals = ParseIBKRCommissionRecords(report, builder, baseCurrency, contractInfos);
            var interestAccrualTotals = ParseIBKRInterestAccrualRecords(report, builder, baseCurrency);
            var debitInterestTotals = ParseIBKRDebitInterestDetails(report);
            var dividendAccrualTotal = ParseIBKRDividendAccruals(report);
            var dividendAccrualChangeTotal = ParseIBKRDividendAccrualChangeRecords(report, builder, baseCurrency, contractInfos);
            var transferTotal = ParseIBKRTransferRecords(report, builder, baseCurrency, account, contractInfos);
            var cashTotals = ParseIBKRCashRecords(report, builder, baseCurrency, tradeTotals, commissionTotals);

            if (mtmTotals.HasData)
            {
                if (commissionTotals.HasData)
                    AssertIBKRMoneyEquals(mtmTotals.Commission, commissionTotals.Total, "IBKR commission", IBKRPrecisionResidualLimit);
                else
                    AssertIBKRMoneyEquals(mtmTotals.Commission, cashTotals.Commission, "IBKR cash commission", IBKRPrecisionResidualLimit);

                AssertIBKRMoneyEquals(mtmTotals.Other, ParseIBKRNavChangeComponent(report, "利息"), "IBKR MTM other");
            }

            var accruedInterest = ParseIBKRNavComponentChange(report, "应计利息");
            if (interestAccrualTotals.HasData)
                AssertIBKRMoneyEquals(accruedInterest, interestAccrualTotals.Total, "IBKR interest accrual NAV");

            if (debitInterestTotals.HasData)
            {
                var accruedBrokerInterest = TryParseIBKRNavComponentChange(report, "应计经纪商利息") ?? accruedInterest;
                AssertIBKRMoneyEquals(accruedBrokerInterest, debitInterestTotals.Total, "IBKR debit interest details");
            }

            var accruedDividend = ParseIBKRNavComponent(report, "应计股息");
            if (dividendAccrualTotal.HasData)
                AssertIBKRMoneyEquals(accruedDividend, dividendAccrualTotal.Total, "IBKR open dividend accrual");

            var accruedDividendChange = ParseIBKRNavComponentChange(report, "应计股息");
            if (dividendAccrualChangeTotal.HasData)
                AssertIBKRMoneyEquals(accruedDividendChange, dividendAccrualChangeTotal.Total, "IBKR dividend accrual change");

            var positionTransfer = ParseIBKRNavChangeComponent(report, "持仓转账");
            if (transferTotal.HasData)
                AssertIBKRMoneyEquals(positionTransfer, transferTotal.Total, "IBKR position transfer");

            AddIBKROtherFxTranslationRecord(report, builder, baseCurrency);
            AddIBKRPrecisionResidualRecord(builder, baseCurrency, expectedNavChange);
            AssertIBKRMoneyEquals(expectedNavChange, builder.NetAssetChangeTotal, "IBKR NAV change records");
            return builder.Records;
        }

        private static void AddIBKRPrecisionResidualRecord(
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            decimal expectedNavChange)
        {
            var residual = expectedNavChange - builder.NetAssetChangeTotal;
            if (residual == 0)
                return;
            if (Math.Abs(residual) > IBKRPrecisionResidualLimit)
            {
                throw new MailParseException(
                    $"Parse IBKR Report Fail, IBKR NAV change records mismatch: expected {expectedNavChange}, got {builder.NetAssetChangeTotal}");
            }

            builder.Add(
                new Currency(residual, baseCurrency),
                "报表精度残差",
                $"PrecisionResidual/{expectedNavChange}/{builder.NetAssetChangeTotal}",
                destAccount: baseCurrency.ToString());
        }

        private IBKRMtmTotals ParseIBKRMtmRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows("按市值计算的表现总结").ToList();
            if (rows.Count == 0)
                return new IBKRMtmTotals(false, 0, 0, 0, 0);

            decimal holdingTotal = 0;
            decimal transactionTotal = 0;
            decimal commissionTotal = 0;
            decimal otherTotal = 0;
            decimal? grandHoldingTotal = null;
            decimal? grandTransactionTotal = null;
            decimal? grandCommissionTotal = null;
            decimal? grandOtherTotal = null;

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

                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown MTM Asset Class: {FormatIBKRCsvRow(row)}");

                AssertIBKRMtmRowTotal(row);
                var holding = ParseIBKRDecimalAt(row, 6, "MTM holding");
                var transaction = ParseIBKRDecimalAt(row, 7, "MTM transaction");
                var commission = ParseIBKRDecimalAt(row, 8, "MTM commission");
                var other = ParseIBKRDecimalAt(row, 9, "MTM other");
                var contract = ResolveIBKRContract(row.Fields[1], assetClass, contractInfos);

                holdingTotal += holding;
                transactionTotal += transaction;
                commissionTotal += commission;
                otherTotal += other;
                builder.Add(
                    new Currency(holding, baseCurrency),
                    "持仓价格变动",
                    $"MTM/{FormatIBKRCsvRow(row)}",
                    destAccount: contract.Code);
                builder.Add(
                    new Currency(transaction, baseCurrency),
                    "交易价格影响",
                    $"MTM/{FormatIBKRCsvRow(row)}",
                    destAccount: contract.Code);
            }

            if (grandHoldingTotal.HasValue)
            {
                AssertIBKRMoneyEquals(grandHoldingTotal.Value, holdingTotal, "IBKR MTM grand holding");
                AssertIBKRMoneyEquals(grandTransactionTotal!.Value, transactionTotal, "IBKR MTM grand transaction");
                AssertIBKRMoneyEquals(grandCommissionTotal!.Value, commissionTotal, "IBKR MTM grand commission");
                AssertIBKRMoneyEquals(grandOtherTotal!.Value, otherTotal, "IBKR MTM grand other");
            }

            return new IBKRMtmTotals(true, holdingTotal, transactionTotal, commissionTotal, otherTotal);
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

        private static IBKRTransactionTotals ParseIBKRTradeSummary(IBKRCsvReport report)
        {
            var rows = report.OptionalDataRows("按代码显示的交易总结").ToList();
            var assetRows = report.OptionalDataRows("按资产类型显示的交易总结").ToList();
            if (rows.Count == 0 && assetRows.Count == 0)
                return new IBKRTransactionTotals(false, 0, 0, 0, 0, 0);

            decimal buyTotal = 0;
            decimal sellTotal = 0;
            int buyQuantity = 0;
            int sellQuantity = 0;
            decimal? reportedBuyTotal = null;
            decimal? reportedSellTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 9);
                var assetClass = row.Fields[0];
                if (assetClass.StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedBuyTotal = ParseIBKRDecimalOrZero(row.Fields[5]);
                    reportedSellTotal = ParseIBKRDecimalOrZero(row.Fields[8]);
                    continue;
                }

                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Trade Summary Asset Class: {FormatIBKRCsvRow(row)}");

                buyQuantity += ParseIBKRIntegerQuantityOrZero(row.Fields[3], "trade buy quantity", row);
                sellQuantity += ParseIBKRIntegerQuantityOrZero(row.Fields[6], "trade sell quantity", row);
                buyTotal += ParseIBKRDecimalOrZero(row.Fields[5]);
                sellTotal += ParseIBKRDecimalOrZero(row.Fields[8]);
            }

            if (reportedBuyTotal.HasValue)
                AssertIBKRMoneyEquals(reportedBuyTotal.Value, buyTotal, "IBKR trade buy summary");
            if (reportedSellTotal.HasValue)
                AssertIBKRMoneyEquals(reportedSellTotal.Value, sellTotal, "IBKR trade sell summary");

            foreach (var row in assetRows)
            {
                AssertIBKRFieldCount(row, 6);
                var assetClass = row.Fields[0];
                if (!IBKRAssetGroups.Contains(assetClass) && !assetClass.StartsWith("总数", StringComparison.Ordinal))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Trade Asset Summary Class: {FormatIBKRCsvRow(row)}");
            }

            return new IBKRTransactionTotals(true, buyTotal + sellTotal, buyTotal, sellTotal, 0, buyQuantity + sellQuantity);
        }

        private IBKRCommissionTotals ParseIBKRCommissionRecords(
            IBKRCsvReport report,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = report.OptionalDataRows("佣金细节").ToList();
            if (rows.Count == 0)
                return new IBKRCommissionTotals(false, 0);

            decimal commissionTotal = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 12);
                if (row.Fields[0].StartsWith("总数", StringComparison.Ordinal))
                {
                    reportedTotal = ParseIBKRDecimalAt(row, 5, "commission total");
                    AssertIBKRCommissionComponents(row);
                    continue;
                }

                var assetClass = row.Fields[0];
                if (!IBKRAssetGroups.Contains(assetClass))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Commission Asset Class: {FormatIBKRCsvRow(row)}");

                var currency = ParseIBKRCurrencyType(row.Fields[1]);
                var contract = ResolveIBKRContract(row.Fields[2], assetClass, contractInfos);
                var rawQuantity = ParseIBKRIntegerQuantityAt(row, 4, "commission quantity");
                var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
                var commission = ParseIBKRDecimalAt(row, 5, "commission");
                AssertIBKRCommissionComponents(row);
                commissionTotal += commission;
                builder.Add(
                    new Currency(commission, currency),
                    "佣金",
                    $"Commission/{FormatIBKRCsvRow(row)}",
                    date: ParseIBKRDateTime(row.Fields[3]),
                    destAccount: contract.Code,
                    holdingQuantity: quantity);
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, commissionTotal, "IBKR commission total");

            return new IBKRCommissionTotals(true, commissionTotal);
        }

        private static void AssertIBKRCommissionComponents(IBKRCsvRow row)
        {
            var commission = ParseIBKRDecimalAt(row, 5, "commission");
            var components =
                ParseIBKRDecimalAt(row, 6, "broker execution fee")
                + ParseIBKRDecimalAt(row, 7, "broker clearing fee")
                + ParseIBKRDecimalAt(row, 8, "third-party execution fee")
                + ParseIBKRDecimalAt(row, 9, "third-party clearing fee")
                + ParseIBKRDecimalAt(row, 10, "third-party transaction fee")
                + ParseIBKRDecimalAt(row, 11, "other commission fee");
            AssertIBKRMoneyEquals(commission, components, $"IBKR commission components {FormatIBKRCsvRow(row)}", IBKRPrecisionResidualLimit);
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
                if (label != "应计利息" && label != "应计转回")
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
                var tax = Math.Abs(ParseIBKRDecimalAt(row, 7, "dividend accrual tax"));
                var fee = Math.Abs(ParseIBKRDecimalAt(row, 8, "dividend accrual fee"));
                var gross = ParseIBKRDecimalAt(row, 10, "dividend accrual gross");
                var net = ParseIBKRDecimalAt(row, 11, "dividend accrual net");
                AssertIBKRMoneyEquals(gross - tax - fee, net, "IBKR dividend accrual change net");

                detailTotal += net;
                builder.Add(
                    new Currency(net, currency),
                    "应计股息",
                    $"DividendAccrualChange/{FormatIBKRCsvRow(row)}",
                    date: date,
                    destAccount: contract.Code,
                    holdingQuantity: quantity);
            }

            if (reportedTotal.HasValue)
                AssertIBKRMoneyEquals(reportedTotal.Value, detailTotal, "IBKR dividend accrual change total");
            if (openingTotal.HasValue && endingTotal.HasValue)
                AssertIBKRMoneyEquals(endingTotal.Value - openingTotal.Value, detailTotal, "IBKR dividend accrual opening ending");

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
                    holdingQuantity: quantity);
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
            foreach (var row in rows)
            {
                AssertIBKRCashReportFieldCount(row);
                if (row.Fields[1] != "基础货币总结")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Cash Currency Summary: {FormatIBKRCsvRow(row)}");

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
                        if (!commissionTotals.HasData)
                            builder.Add(new Currency(amount, baseCurrency), "佣金", $"CashReport/{FormatIBKRCsvRow(row)}", destAccount: baseCurrency.ToString());
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
                        builder.Add(new Currency(amount, baseCurrency), label, $"CashReport/{FormatIBKRCsvRow(row)}", destAccount: baseCurrency.ToString());
                        break;
                    case "现金外汇换算收益/损失":
                        // 已由 MTM 外汇行体现，避免同一基础货币折算影响重复生成 record。
                        break;
                    case "存款":
                    case "取款":
                    case "转入":
                    case "转出":
                    case "内部转账":
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

            if (transactionTotals.HasData)
            {
                AssertIBKRMoneyEquals(tradeBuy, transactionTotals.BuyProceeds, "IBKR cash buy transactions");
                AssertIBKRMoneyEquals(tradeSell, transactionTotals.SellProceeds, "IBKR cash sell transactions");
            }

            if (commissionTotals.HasData)
                AssertIBKRMoneyEquals(commission, commissionTotals.Total, "IBKR cash commission", IBKRPrecisionResidualLimit);

            return new IBKRCashTotals(commission, tradeBuy, tradeSell);
        }

        private List<Holding> ParseIBKRHoldings(
            IBKRCsvReport report,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            string sourceName)
        {
            var holdings = new List<Holding>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var holding in ParseIBKRPositionAndMtmHoldings(report, account, baseCurrency, contractInfos))
                AddIBKRHolding(holdings, seen, holding);

            var cash = ParseIBKREndingCash(report);
            AddIBKRHolding(holdings, seen, new Holding(baseCurrency.ToString(), HoldingType.Cash)
            {
                Account = account,
                desc = $"IBKR cash {baseCurrency}",
                displayText = baseCurrency.ToString(),
                currentPrice = new Currency(cash, baseCurrency)
            });

            foreach (var holding in ParseIBKRNavAdjustmentHoldings(report, account, baseCurrency))
                AddIBKRHolding(holdings, seen, holding);

            ValidateIBKRStockPositionSummary(report, holdings);
            if (holdings.Count == 0)
                throw new MailParseException($"Parse IBKR Report Fail, Missing Holdings: {sourceName}");

            return holdings;
        }

        private static List<Holding> ParseIBKRNavAdjustmentHoldings(
            IBKRCsvReport report,
            Account account,
            CurrencyType baseCurrency)
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

                var amount = ParseIBKRDecimalAt(row, 4, "NAV component amount");
                if (component is "抵押品价值" or "借出证券")
                {
                    if (amount != 0)
                        throw new MailParseException($"Parse IBKR Report Fail, Unsupported Nonzero Stock Yield Enhancement NAV Component: {component}/{amount}");
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

                holdings.Add(new Holding(code, holdingType)
                {
                    Account = account,
                    desc = $"IBKR NAV {component}",
                    displayText = component,
                    currentPrice = new Currency(amount, baseCurrency)
                });
            }

            return holdings;
        }

        private List<Holding> ParseIBKRPositionAndMtmHoldings(
            IBKRCsvReport report,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var holdings = new List<Holding>();
            foreach (var row in report.OptionalDataRows("持仓与以市值计的盈亏"))
            {
                if (!IsIBKRPositionSummaryRow(row))
                    continue;

                var group = row.Fields[1];
                if (group != "股票" && group != "债券")
                    continue;

                var currency = ParseIBKRCurrencyType(row.Fields[2]);
                var quantity = ParseIBKRIntegerQuantityAt(row, 6, "holding quantity");
                if (quantity == 0)
                    continue;

                var currentPrice = ParseIBKRDecimalAt(row, 8, "holding price");
                var currentValue = ParseIBKRDecimalAt(row, 10, "holding value");
                var contract = ResolveIBKRContract(row.Fields[3], group, contractInfos);
                holdings.Add(CreateIBKRHolding(account, contract, quantity, currentPrice, currentValue, currency, row.Fields[4]));
            }

            return holdings;
        }

        private static bool IsIBKRPositionSummaryRow(IBKRCsvRow row)
        {
            if (row.Fields.Count == 16 && row.Fields[0] == "Summary")
                return true;
            if (row.Fields.Count == 16 && row.Fields[0] == "Details")
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
            CurrencyType currency,
            string rowDescription)
        {
            var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
            var currentPrice = statementPrice;
            if (quantity != 0 && currentPrice * quantity != currentValue)
                currentPrice = currentValue / quantity;
            if (quantity != 0)
                AssertIBKRMoneyEquals(currentValue, currentPrice * quantity, $"IBKR holding value {contract.Code}", IBKRPrecisionResidualLimit);

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

        private static string GetIBKRHoldingKey(Holding holding)
        {
            return $"{holding.code}\t{holding.holdingType}";
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
            contractInfos[created.Code] = created;
            return created;
        }

        private static string ExtractIBKRBondCode(string text)
        {
            var match = Regex.Match(text, @"T\s+\d+(?:\.\d+)?\s+\d{2}/\d{2}/\d{2}", RegexOptions.IgnoreCase);
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
                var netQuantity = ParseIBKRIntegerQuantityAt(row, 7, "net stock quantity");
                if (!stockHoldings.TryGetValue(code, out var holdingQuantity))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Holding: {FormatIBKRCsvRow(row)}");
                if (holdingQuantity != netQuantity)
                    throw new MailParseException($"Parse IBKR Report Fail, Stock Holding Quantity Mismatch: {code}/{holdingQuantity}/{netQuantity}");
            }
        }

        private static decimal ParseIBKREndingCash(IBKRCsvReport report)
        {
            var row = report.RequireDataRows("现金报告").FirstOrDefault(row => row.Fields.Count > 2 && row.Fields[0] == "期末现金")
                ?? throw new MailParseException("Parse IBKR Report Fail, Missing Ending Cash");
            return ParseIBKRDecimalAt(row, 2, "ending cash");
        }

        private static decimal ParseIBKREndingNav(IBKRCsvReport report)
        {
            var row = FindIBKRNavRow(report, "总数");
            return ParseIBKRDecimalAt(row, 4, "ending NAV");
        }

        private static decimal ParseIBKRStartingNav(IBKRCsvReport report)
        {
            return ParseIBKRNavChangeValues(report).Start;
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
                if (label is "按市值计价" or "持仓转账" or "利息" or "应计利息变更" or "应计股息的变化" or "其它外汇换算" or "佣金")
                {
                    componentTotal += amount;
                    continue;
                }

                throw new MailParseException($"Parse IBKR Report Fail, Unknown NAV Change Component: {label}/{amount}");
            }

            AssertIBKRMoneyEquals(end - start, componentTotal, "IBKR NAV change components", IBKRPrecisionResidualLimit);
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
            AssertIBKRMoneyEquals(end - start, ParseIBKRDecimalAt(totalRow, 5, "NAV change"), "IBKR NAV total change", IBKRPrecisionResidualLimit);

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
            var row = report.RequireDataRows("净资产值")
                .FirstOrDefault(row => row.Fields.Count > 5 && row.Fields[0] == component);
            return row is null ? 0 : ParseIBKRDecimalAt(row, 5, $"NAV component change {component}");
        }

        private static decimal? TryParseIBKRNavComponentChange(IBKRCsvReport report, string component)
        {
            var row = report.RequireDataRows("净资产值")
                .FirstOrDefault(row => row.Fields.Count > 5 && row.Fields[0] == component);
            return row is null ? null : ParseIBKRDecimalAt(row, 5, $"NAV component change {component}");
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

                var sectionName = fields[0].Trim('\uFEFF', ' ');
                var rowType = fields[1].Trim();
                if (!IBKRCsvSections.Contains(sectionName))
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown CSV Section: {sectionName}");
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

        private static void ValidateIBKRCsvHeaders(IBKRCsvReport report)
        {
            foreach (var (sectionName, section) in report.Sections)
            {
                var expectedHeaders = IBKRCsvHeaders[sectionName];
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

            ValidateIBKRStockYieldEnhancementLoanSection(report);
            var interestTotal = ValidateIBKRMoneyDetailSection(report, IBKRInterestSection, 4);
            var bondInterestTotal = ValidateIBKRMoneyDetailSection(report, IBKRBondInterestReceivedSection, 5, 3);
            if (interestTotal.HasValue && bondInterestTotal.HasValue)
                AssertIBKRMoneyEquals(interestTotal.Value, bondInterestTotal.Value, "IBKR interest detail sections", IBKRPrecisionResidualLimit);
        }

        private static void ValidateIBKRStockYieldEnhancementLoanSection(IBKRCsvReport report)
        {
            var rows = report.OptionalDataRows(IBKRStockYieldEnhancementLoanSection).ToList();
            if (rows.Count == 0)
                return;

            decimal collateralTotal = 0;
            decimal? reportedTotal = null;
            foreach (var row in rows)
            {
                AssertIBKRFieldCount(row, 9);
                if (row.Fields[0] == "总数")
                {
                    if (reportedTotal.HasValue)
                        throw new MailParseException($"Parse IBKR Report Fail, Duplicate Stock Yield Enhancement Total: {FormatIBKRCsvRow(row)}");
                    if (row.Fields.Take(8).Skip(1).Any(field => !String.IsNullOrWhiteSpace(field)))
                        throw new MailParseException($"Parse IBKR Report Fail, Invalid Stock Yield Enhancement Total: {FormatIBKRCsvRow(row)}");

                    reportedTotal = ParseIBKRDecimalAt(row, 8, "stock yield enhancement collateral total");
                    continue;
                }

                if (row.Fields[0] != "股票")
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Stock Yield Enhancement Asset Class: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRCurrencyType(row.Fields[1]);
                if (String.IsNullOrWhiteSpace(row.Fields[2]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Code: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDate(row.Fields[3]);
                if (String.IsNullOrWhiteSpace(row.Fields[4]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Description: {FormatIBKRCsvRow(row)}");
                if (String.IsNullOrWhiteSpace(row.Fields[6]))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Transaction Number: {FormatIBKRCsvRow(row)}");
                _ = ParseIBKRDecimalAt(row, 7, "stock yield enhancement quantity");
                collateralTotal += ParseIBKRDecimalAt(row, 8, "stock yield enhancement collateral");
            }

            if (!reportedTotal.HasValue)
                throw new MailParseException($"Parse IBKR Report Fail, Missing Stock Yield Enhancement Total: {report.SourceName}");
            AssertIBKRMoneyEquals(reportedTotal.Value, collateralTotal, "IBKR stock yield enhancement collateral total");

            foreach (var row in report.OptionalNoteRows(IBKRStockYieldEnhancementLoanSection))
            {
                if (row.Fields.Count != 1 || String.IsNullOrWhiteSpace(row.Fields[0]))
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Stock Yield Enhancement Note: {FormatIBKRCsvRow(row)}");
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
            AssertIBKRMoneyEquals(reportedTotal.Value, total, $"IBKR {sectionName} total", IBKRPrecisionResidualLimit);
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

        private static DateTime ParseIBKRStatementPeriod(string text)
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

        private static void AssertIBKRMoneyEquals(decimal expected, decimal actual, string context, decimal tolerance)
        {
            if (Math.Abs(expected - actual) > tolerance)
                throw new MailParseException($"Parse IBKR Report Fail, {context} mismatch: expected {expected}, got {actual}");
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
            var attachments = new List<InMemoryIBKRReportAttachment>();
            foreach (var attachment in message.Attachments)
            {
                if (attachment is not MimePart mimePart)
                    continue;

                var fileName = GetAttachmentFileName(attachment);
                if (!TryParseIBKRReportAttachmentName(fileName, out var attachmentInfo)
                    || !IsDailyMyBookReportType(attachmentInfo.ReportType)
                    || attachmentInfo.ReportDate.Date != reportDate.Date
                    || !IsIBKRCsvAttachment(fileName))
                {
                    continue;
                }

                attachments.Add(new InMemoryIBKRReportAttachment(
                    fileName,
                    attachmentInfo.ReportType,
                    attachmentInfo.ReportId,
                    attachmentInfo.ReportDate,
                    ReadMimePartBytes(mimePart)));
            }

            return attachments;
        }

        private static bool HasIBKRReportAttachment(MimeMessage message, DateTime reportDate)
        {
            return message.Attachments.Any(attachment =>
            {
                var fileName = GetAttachmentFileName(attachment);
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
            List<Holding> Holdings,
            List<AccountBalance> AccountBalances,
            List<AccountBalance> BeginningAccountBalances,
            List<AccountInternalId> InternalCardNos);

        private sealed record IBKRContractInfo(string Code, string Description, HoldingType HoldingType, string DisplayText);

        private sealed record IBKRMtmTotals(bool HasData, decimal Holding, decimal Transaction, decimal Commission, decimal Other);

        private sealed record IBKRTransactionTotals(
            bool HasData,
            decimal Proceeds,
            decimal BuyProceeds,
            decimal SellProceeds,
            decimal Commission,
            int Quantity);

        private sealed record IBKRCommissionTotals(bool HasData, decimal Total);

        private sealed record IBKRInterestTotals(bool HasData, decimal Total);

        private sealed record IBKRCashTotals(decimal Commission, decimal TradeBuy, decimal TradeSell);

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

            public IBKRRecordBuilder(Account account, DateTime reportDate, string sourceName, CurrencyType baseCurrency)
            {
                this.account = account;
                this.reportDate = reportDate;
                this.sourceName = sourceName;
                this.baseCurrency = baseCurrency;
            }

            public Records Records { get; } = new();
            public decimal NetAssetChangeTotal { get; private set; }

            public void Add(
                Currency amount,
                string reason,
                string source,
                bool isInternal = false,
                bool affectsNetAsset = true,
                DateTime? date = null,
                string destAccount = "",
                int holdingQuantity = 0)
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
