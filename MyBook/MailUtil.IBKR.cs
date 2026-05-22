using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using MimeKit;

namespace MyBook
{
    // IBKR report discovery, HTML parsing, and account matching.
    partial class MailUtil
    {
        private const StatementImportProvider IBKRProvider = StatementImportProvider.IBKRReportMail;
        private const string IBKRReportSender = "donotreply@interactivebrokers.com";
        private const string IBKRDailyReportType = "DailyMyBook";
        private const int IBKRMissingReportLimitDays = 14;

        private static readonly HashSet<string> IBKRAssetGroups = new(StringComparer.Ordinal)
        {
            "股票",
            "债券",
            "外汇"
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
            var files = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.html")
                .Where(file => TryParseIBKRReportAttachmentName(Path.GetFileName(file), out var attachmentInfo)
                    && IsDailyMyBookReportType(attachmentInfo.ReportType))
                .OrderBy(file =>
                {
                    TryParseIBKRReportAttachmentName(Path.GetFileName(file), out var attachmentInfo);
                    return attachmentInfo.ReportDate;
                })
                .ToList();
            if (files.Count == 0)
                throw new FileNotFoundException("No local IBKR html report found.");

            var reports = files
                .Select(file =>
                {
                    TryParseIBKRReportAttachmentName(Path.GetFileName(file), out var attachmentInfo);
                    var html = File.ReadAllText(file, Encoding.UTF8);
                    return ParseIBKRReportHtml(html, attachmentInfo.ReportDate, file);
                })
                .ToList();
            SaveIBKRParsedReports(reports);
        }

        public bool DebugFetchLocalIBKRReport(string path)
        {
            if (!TryParseIBKRReportAttachmentName(Path.GetFileName(path), out var attachmentInfo)
                || !IsDailyMyBookReportType(attachmentInfo.ReportType))
            {
                throw new ArgumentException($"Invalid IBKR report file name: {path}");
            }

            var html = File.ReadAllText(path, Encoding.UTF8);
            return SaveIBKRParsedReports([ParseIBKRReportHtml(html, attachmentInfo.ReportDate, path)])[0];
        }

        public async Task FetchIBKRReports()
        {
            var date = GetNextDailyStatementDate(IBKRProvider);
            var missingDays = 0;
            while (date <= DateTime.Today)
            {
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
                    Console.WriteLine($"parse IBKR report mail fail: no supported html attachment, subject={message.Subject}");
                    throw new MailParseException($"Parse IBKR Report Fail, Missing Supported HTML Attachment: {message.Subject}");
                }

                var mailDate = GetMailDate(message);
                foreach (var attachment in reportAttachments)
                {
                    if (attachment.ReportDate.Date != date.Date)
                        throw new MailParseException($"Parse IBKR Report Fail, Date Mismatch: expected {date:yyyy-MM-dd}, got {attachment.ReportDate:yyyy-MM-dd}");

                    Console.WriteLine($"Load IBKR report in memory: {attachment.FileName}, id={attachment.ReportId}, bytes={attachment.Content.Length}");
                    var html = Encoding.UTF8.GetString(attachment.Content);
                    reports.Add(ParseIBKRReportHtml(html, attachment.ReportDate, attachment.FileName, mailDate));
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
                    report.BeginningAccountBalances)));

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

        private IBKRParsedReport ParseIBKRReportHtml(string html, DateTime reportDate, string sourceName, DateTime? importDate = null)
        {
            if (String.IsNullOrWhiteSpace(html))
                throw new MailParseException("Parse IBKR Report Fail, Empty HTML");

            var doc = new HtmlDocument();
            doc.OptionFixNestedTags = true;
            doc.LoadHtml(html);

            var (accountId, baseCurrency) = ParseIBKRAccountInfo(doc);
            var account = GetIBKRAccount(accountId);
            var contractInfos = ParseIBKRContractInfos(doc);
            var holdings = ParseIBKRHoldings(doc, account, baseCurrency, contractInfos, sourceName);
            var startingNav = ParseIBKRStartingNav(doc);
            var endingNav = ParseIBKREndingNav(doc);
            var navChange = ParseIBKRNavChange(doc);
            var records = ParseIBKRRecords(doc, account, baseCurrency, contractInfos, reportDate.Date, sourceName, navChange);

            return new IBKRParsedReport(
                reportDate.Date,
                (importDate ?? reportDate).Date,
                BuildIBKRStatementKey(account, reportDate),
                accountId,
                account,
                records,
                holdings,
                [new AccountBalance(account, new Currency(endingNav, baseCurrency))],
                [new AccountBalance(account, new Currency(startingNav, baseCurrency))]);
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

        private static (string AccountId, CurrencyType BaseCurrency) ParseIBKRAccountInfo(HtmlDocument doc)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblAccountInformation_");
            var accountId = "";
            var baseCurrency = "";
            foreach (var row in rows)
            {
                if (row.Count < 2)
                    continue;

                if (row[0] == "账户")
                    accountId = ParseIBKRAccountId(row[1]) ?? row[1];
                else if (row[0] == "基础货币")
                    baseCurrency = row[1];
            }

            if (String.IsNullOrWhiteSpace(accountId))
                throw new MailParseException("Parse IBKR Report Fail, Missing Account");
            if (String.IsNullOrWhiteSpace(baseCurrency))
                throw new MailParseException("Parse IBKR Report Fail, Missing Base Currency");

            return (accountId, ParseIBKRCurrencyType(baseCurrency));
        }

        private Dictionary<string, IBKRContractInfo> ParseIBKRContractInfos(HtmlDocument doc)
        {
            var contracts = new Dictionary<string, IBKRContractInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var table in ReadIBKRSectionTables(doc, "tblContractInfo"))
            {
                var currentGroup = "";
                foreach (var row in table)
                {
                    if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                        continue;

                    if (TryReadIBKRSingleValue(row, out var singleValue))
                    {
                        if (IBKRAssetGroups.Contains(singleValue))
                            currentGroup = singleValue;
                        continue;
                    }

                    if (row.Count < 8 || String.IsNullOrWhiteSpace(row[0]))
                        continue;

                    if (currentGroup == "股票")
                    {
                        var holdingType = ParseIBKRHoldingType(row.Count > 5 ? row[5] : "");
                        AddIBKRContractInfo(
                            contracts,
                            new IBKRContractInfo(row[0], row[1], holdingType, row[0]));
                    }
                    else if (currentGroup == "债券")
                    {
                        var issuer = row.Count > 8 ? row[8] : "";
                        AddIBKRContractInfo(
                            contracts,
                            new IBKRContractInfo(row[0], row[1], HoldingType.UST, NormalizeIBKRBondDisplayText(row[0], issuer)));
                    }
                    else if (!String.IsNullOrWhiteSpace(currentGroup))
                    {
                        throw new MailParseException($"Parse IBKR Report Fail, Unknown Contract Group: {currentGroup}");
                    }
                }
            }

            return contracts;
        }

        private static void AddIBKRContractInfo(Dictionary<string, IBKRContractInfo> contracts, IBKRContractInfo contract)
        {
            if (String.IsNullOrWhiteSpace(contract.Code))
                return;

            contracts[contract.Code] = contract;
        }

        private Records ParseIBKRRecords(
            HtmlDocument doc,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            DateTime reportDate,
            string sourceName,
            decimal expectedNavChange)
        {
            var builder = new IBKRRecordBuilder(account, reportDate, sourceName, baseCurrency);
            var mtmTotals = ParseIBKRMtmRecords(doc, builder, baseCurrency, contractInfos);
            var transactionTotals = ParseIBKRTransactionRecords(doc, builder, baseCurrency, contractInfos, !mtmTotals.HasData);
            var commissionTotals = ParseIBKRCommissionRecords(doc, builder, baseCurrency, contractInfos);
            var interestTotals = ParseIBKRInterestRecords(doc, builder, baseCurrency);
            var interestAccrualTotal = ParseIBKRInterestAccrualRecords(doc, builder, baseCurrency);
            var dividendAccrualTotal = ParseIBKRDividendAccrualRecords(doc, builder, baseCurrency);
            var positionTransferTotal = ParseIBKRPositionTransferRecords(doc, builder, baseCurrency, contractInfos);
            var cashTotals = ParseIBKRCashRecords(doc, builder, baseCurrency, transactionTotals, commissionTotals, interestTotals);
            AssertIBKRMoneyEquals(ParseIBKRNavChangeComponent(doc, "持仓转账"), positionTransferTotal, "IBKR position transfer");

            if (mtmTotals.HasData)
            {
                AssertIBKRMoneyEquals(mtmTotals.Transaction, transactionTotals.Mtm, "IBKR transaction MTM");
                var commissionCheck = commissionTotals.HasData ? commissionTotals.Total : cashTotals.Commission;
                AssertIBKRMoneyEquals(mtmTotals.Commission, commissionCheck, "IBKR commission");
                AssertIBKRMoneyEquals(mtmTotals.Other, interestTotals.Total, "IBKR MTM other");
            }

            var recordNavChange = Decimal.Round(builder.NetAssetChangeTotal, 2);
            AssertIBKRMoneyEquals(expectedNavChange, recordNavChange, "IBKR NAV change records");
            _ = interestAccrualTotal + dividendAccrualTotal;
            return builder.Records;
        }

        private IBKRMtmTotals ParseIBKRMtmRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblPosAndMTM_");
            var usePositionAndMtm = rows.Count > 0;
            if (!usePositionAndMtm)
                rows = ReadIBKRSectionFirstTableRows(doc, "tblMtmPerfSumByUnderlying_");
            if (rows.Count == 0)
                return new IBKRMtmTotals(false, 0, 0, 0, 0);

            var holdingColumn = usePositionAndMtm ? FindIBKRGroupedColumn(rows, "以市值计的盈亏", "持仓") : 5;
            var transactionColumn = usePositionAndMtm ? FindIBKRGroupedColumn(rows, "以市值计的盈亏", "交易") : 6;
            var commissionColumn = usePositionAndMtm ? FindIBKRGroupedColumn(rows, "以市值计的盈亏", "佣金") : 7;
            var otherColumn = usePositionAndMtm ? FindIBKRGroupedColumn(rows, "以市值计的盈亏", "其它") : 8;
            var currentGroup = "";
            var currentCurrency = baseCurrency;
            decimal holdingTotal = 0;
            decimal transactionTotal = 0;
            decimal commissionTotal = 0;
            decimal otherTotal = 0;

            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                if (!IBKRAssetGroups.Contains(currentGroup))
                    throw new MailParseException($"Parse IBKR Report Fail, Missing MTM Group: {String.Join(",", row)}");
                if (usePositionAndMtm && IsIBKRPositionDetailRow(row))
                    continue;

                var holding = ParseIBKRDecimalAt(row, holdingColumn, "MTM holding");
                var transaction = ParseIBKRDecimalAt(row, transactionColumn, "MTM transaction");
                var commission = ParseIBKRDecimalAt(row, commissionColumn, "MTM commission");
                var other = ParseIBKRDecimalAt(row, otherColumn, "MTM other");
                var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                var code = contract.Code;

                holdingTotal += holding;
                transactionTotal += transaction;
                commissionTotal += commission;
                otherTotal += other;
                builder.Add(
                    new Currency(holding, currentCurrency),
                    "持仓价格变动",
                    $"MTM/{String.Join(",", row)}",
                    destAccount: code);
                builder.Add(
                    new Currency(transaction, currentCurrency),
                    "交易价格影响",
                    $"MTM/{String.Join(",", row)}",
                    destAccount: code);
            }

            return new IBKRMtmTotals(true, holdingTotal, transactionTotal, commissionTotal, otherTotal);
        }

        private IBKRTransactionTotals ParseIBKRTransactionRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            bool recordTransactionMtm)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblTransactions_");
            if (rows.Count == 0)
                return new IBKRTransactionTotals(false, 0, 0, 0, 0, 0);

            var currentGroup = "";
            var currentCurrency = baseCurrency;
            decimal proceedsTotal = 0;
            decimal buyTotal = 0;
            decimal sellTotal = 0;
            decimal commissionTotal = 0;
            decimal mtmTotal = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                var code = contract.Code;
                var tradeTime = ParseIBKRDateTime(row[1]);
                var rawQuantity = ParseIBKRIntegerQuantityAt(row, 2, "transaction quantity");
                var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
                var proceeds = ParseIBKRDecimalAt(row, 5, "transaction proceeds");
                var commission = ParseIBKRDecimalAt(row, 6, "transaction commission");
                var mtm = ParseIBKRDecimalAt(row, 9, "transaction MTM");
                proceedsTotal += proceeds;
                commissionTotal += commission;
                mtmTotal += mtm;
                if (proceeds < 0)
                    buyTotal += proceeds;
                else
                    sellTotal += proceeds;

                builder.Add(
                    new Currency(proceeds, currentCurrency),
                    quantity < 0 ? "卖出" : "买入",
                    $"Transactions/{String.Join(",", row)}",
                    isInternal: true,
                    affectsNetAsset: false,
                    date: tradeTime,
                    destAccount: code,
                    holdingQuantity: quantity);

                if (recordTransactionMtm)
                {
                    builder.Add(
                        new Currency(mtm, currentCurrency),
                        "交易价格影响",
                        $"Transactions/{String.Join(",", row)}",
                        date: tradeTime,
                        destAccount: code,
                        holdingQuantity: quantity);
                }
            }

            return new IBKRTransactionTotals(true, proceedsTotal, buyTotal, sellTotal, commissionTotal, mtmTotal);
        }

        private IBKRCommissionTotals ParseIBKRCommissionRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblUnbundledCommissionDetails_");
            if (rows.Count == 0)
                return new IBKRCommissionTotals(false, 0);

            var currentGroup = "";
            var currentCurrency = baseCurrency;
            decimal commissionTotal = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                var commission = ParseIBKRDecimalAt(row, 3, "commission");
                commissionTotal += commission;
                builder.Add(
                    new Currency(commission, currentCurrency),
                    "佣金",
                    $"Commission/{String.Join(",", row)}",
                    destAccount: contract.Code);
            }

            return new IBKRCommissionTotals(true, commissionTotal);
        }

        private IBKRInterestTotals ParseIBKRInterestRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblCombInt_");
            if (rows.Count == 0)
                return new IBKRInterestTotals(false, 0);

            var currentCurrency = baseCurrency;
            decimal total = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue)
                    && TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                {
                    currentCurrency = parsedCurrency;
                    continue;
                }

                if (row.Count < 3)
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Interest Row: {String.Join(",", row)}");

                var amount = ParseIBKRDecimalAt(row, 2, "interest");
                total += amount;
                builder.Add(
                    new Currency(amount, currentCurrency),
                    $"利息:{row[1]}",
                    $"Interest/{String.Join(",", row)}",
                    date: ParseIBKRDate(row[0]),
                    destAccount: currentCurrency.ToString());
            }

            return new IBKRInterestTotals(true, total);
        }

        private decimal ParseIBKRInterestAccrualRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblInterestAccruals_");
            decimal total = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || TryReadIBKRSingleValue(row, out _))
                    continue;
                if (row.Count < 2)
                    continue;

                var label = row[0];
                var amount = ParseIBKRDecimalAt(row, 1, "interest accrual");
                if (label == "期初应计余额" || label == "期末应计余额")
                    continue;
                if (label != "应计利息" && label != "应计转回")
                {
                    if (amount != 0)
                        throw new MailParseException($"Parse IBKR Report Fail, Unknown Interest Accrual Row: {String.Join(",", row)}");
                    continue;
                }

                total += amount;
                builder.Add(
                    new Currency(amount, baseCurrency),
                    label,
                    $"InterestAccruals/{String.Join(",", row)}",
                    destAccount: "ACCRUED_INTEREST");
            }

            return total;
        }

        private decimal ParseIBKRDividendAccrualRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblChangeInDividend_");
            var currentGroup = "";
            var currentCurrency = baseCurrency;
            decimal total = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                if (row[0].StartsWith("期初应计股息", StringComparison.Ordinal)
                    || row[0].StartsWith("期末应计股息", StringComparison.Ordinal))
                {
                    continue;
                }

                if (currentGroup != "股票" || row.Count < 10)
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid Dividend Accrual Row: {String.Join(",", row)}");

                var tax = Math.Abs(ParseIBKRDecimalAt(row, 5, "dividend tax"));
                var fee = Math.Abs(ParseIBKRDecimalAt(row, 6, "dividend fee"));
                var gross = ParseIBKRDecimalAt(row, 8, "gross dividend");
                var net = ParseIBKRDecimalAt(row, 9, "net dividend");
                AssertIBKRMoneyEquals(gross - tax - fee, net, "IBKR dividend accrual net");

                total += net;
                builder.Add(
                    new Currency(gross, currentCurrency),
                    "应计股息",
                    $"DividendAccrual/{String.Join(",", row)}",
                    date: ParseIBKRDate(row[1]),
                    destAccount: row[0]);
                builder.Add(
                    new Currency(-tax, currentCurrency),
                    "应计股息税",
                    $"DividendAccrual/{String.Join(",", row)}",
                    date: ParseIBKRDate(row[1]),
                    destAccount: row[0]);
                builder.Add(
                    new Currency(-fee, currentCurrency),
                    "应计股息费用",
                    $"DividendAccrual/{String.Join(",", row)}",
                    date: ParseIBKRDate(row[1]),
                    destAccount: row[0]);
            }

            return total;
        }

        private decimal ParseIBKRPositionTransferRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblPosAndMTM_");
            if (rows.Count == 0)
                return 0;

            var quantityColumn = FindIBKRGroupedColumn(rows, "数量", "当前");
            var currentValueColumn = FindIBKRGroupedColumn(rows, "市场价值", "当前");
            var totalMtmColumn = FindIBKRGroupedColumn(rows, "以市值计的盈亏", "总数");
            var currentGroup = "";
            var currentCurrency = baseCurrency;
            decimal total = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                if (!TryGetIBKRPositionDetail(row, out var detailDate, out var action))
                    continue;

                var currentValue = ParseIBKRDecimalOrZero(row.Count > currentValueColumn ? row[currentValueColumn] : "");
                var totalMtm = ParseIBKRDecimalOrZero(row.Count > totalMtmColumn ? row[totalMtmColumn] : "");
                if (action == "转账")
                {
                    var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                    var rawQuantity = ParseIBKRIntegerQuantityAt(row, quantityColumn, "position transfer quantity");
                    var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
                    total += currentValue;
                    builder.Add(
                        new Currency(currentValue, currentCurrency),
                        "持仓转账",
                        $"PositionTransfer/{String.Join(",", row)}",
                        isInternal: true,
                        date: detailDate,
                        destAccount: contract.Code,
                        holdingQuantity: quantity);
                    continue;
                }

                if (currentValue != 0 || totalMtm != 0)
                    throw new MailParseException($"Parse IBKR Report Fail, Unknown Position Detail Row: {String.Join(",", row)}");
            }

            return total;
        }

        private IBKRCashTotals ParseIBKRCashRecords(
            HtmlDocument doc,
            IBKRRecordBuilder builder,
            CurrencyType baseCurrency,
            IBKRTransactionTotals transactionTotals,
            IBKRCommissionTotals commissionTotals,
            IBKRInterestTotals interestTotals)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblCashReport_");
            decimal commission = 0;
            decimal tradeBuy = 0;
            decimal tradeSell = 0;
            decimal interest = 0;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || TryReadIBKRSingleValue(row, out _))
                    continue;
                if (row.Count < 2 || !TryParseIBKRDecimal(row[1], out var amount))
                    continue;

                var label = row[0];
                switch (label)
                {
                    case "期初现金":
                    case "期末现金":
                    case "期末已结算现金":
                        break;
                    case "佣金":
                        commission += amount;
                        if (!commissionTotals.HasData)
                            builder.Add(new Currency(amount, baseCurrency), "佣金", $"CashReport/{String.Join(",", row)}", destAccount: baseCurrency.ToString());
                        break;
                    case "交易（买入）":
                        tradeBuy += amount;
                        if (!transactionTotals.HasData)
                        {
                            builder.Add(
                                new Currency(amount, baseCurrency),
                                "交易买入现金",
                                $"CashReport/{String.Join(",", row)}",
                                isInternal: true,
                                affectsNetAsset: false,
                                destAccount: baseCurrency.ToString());
                        }
                        break;
                    case "交易（卖出）":
                        tradeSell += amount;
                        if (!transactionTotals.HasData)
                        {
                            builder.Add(
                                new Currency(amount, baseCurrency),
                                "交易卖出现金",
                                $"CashReport/{String.Join(",", row)}",
                                isInternal: true,
                                affectsNetAsset: false,
                                destAccount: baseCurrency.ToString());
                        }
                        break;
                    case "支付和收到的经纪商利息":
                    case "支付和收到的债券利息":
                        interest += amount;
                        if (!interestTotals.HasData)
                            builder.Add(new Currency(amount, baseCurrency), label, $"CashReport/{String.Join(",", row)}", destAccount: baseCurrency.ToString());
                        break;
                    case "存款":
                    case "取款":
                    case "转入":
                    case "转出":
                    case "内部转账":
                        builder.Add(
                            new Currency(amount, baseCurrency),
                            label,
                            $"CashReport/{String.Join(",", row)}",
                            isInternal: true,
                            destAccount: baseCurrency.ToString());
                        break;
                    case "股息":
                    case "代替股息的支付":
                    case "代扣税款":
                    case "现金外汇换算收益/损失":
                        builder.Add(new Currency(amount, baseCurrency), label, $"CashReport/{String.Join(",", row)}", destAccount: baseCurrency.ToString());
                        break;
                    default:
                        if (amount != 0)
                            throw new MailParseException($"Parse IBKR Report Fail, Unknown Cash Row: {String.Join(",", row)}");
                        break;
                }
            }

            if (transactionTotals.HasData)
            {
                AssertIBKRMoneyEquals(tradeBuy, transactionTotals.BuyProceeds, "IBKR cash buy transactions");
                AssertIBKRMoneyEquals(tradeSell, transactionTotals.SellProceeds, "IBKR cash sell transactions");
            }

            if (commissionTotals.HasData)
                AssertIBKRMoneyEquals(commission, commissionTotals.Total, "IBKR cash commission");
            if (interestTotals.HasData)
                AssertIBKRMoneyEquals(interest, interestTotals.Total, "IBKR cash interest");

            return new IBKRCashTotals(commission, tradeBuy, tradeSell, interest);
        }

        private List<Holding> ParseIBKRHoldings(
            HtmlDocument doc,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos,
            string sourceName)
        {
            var holdings = new List<Holding>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var parsedPositionHoldings = ParseIBKRPositionAndMtmHoldings(doc, account, baseCurrency, contractInfos);
            if (parsedPositionHoldings.Count == 0)
                parsedPositionHoldings = ParseIBKROpenPositionHoldings(doc, account, baseCurrency, contractInfos);

            foreach (var holding in parsedPositionHoldings)
                AddIBKRHolding(holdings, seen, holding);

            var cash = RoundIBKRMoney(ParseIBKREndingCash(doc));
            AddIBKRHolding(holdings, seen, new Holding(baseCurrency.ToString(), HoldingType.Cash)
            {
                Account = account,
                desc = $"IBKR cash {baseCurrency}",
                displayText = baseCurrency.ToString(),
                currentPrice = new Currency(cash, baseCurrency)
            });
            foreach (var holding in ParseIBKRNavAdjustmentHoldings(doc, account, baseCurrency))
                AddIBKRHolding(holdings, seen, holding);

            if (holdings.Count == 0)
                throw new MailParseException($"Parse IBKR Report Fail, Missing Holdings: {sourceName}");

            return holdings;
        }

        private static List<Holding> ParseIBKRNavAdjustmentHoldings(
            HtmlDocument doc,
            Account account,
            CurrencyType baseCurrency)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblNAV_");
            var holdings = new List<Holding>();
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row) || row.Count <= 4)
                    continue;

                var component = row[0];
                if (component is "现金" or "股票" or "债券")
                    continue;

                if (!TryParseIBKRDecimal(row[4], out var parsedAmount))
                    continue;

                var amount = RoundIBKRMoney(parsedAmount);
                if (amount == 0)
                    continue;

                var (code, holdingType) = component switch
                {
                    "应计利息" => ("ACCRUED_INTEREST", HoldingType.Accrued),
                    "应计股息" => ("ACCRUED_DIVIDEND", HoldingType.Accrued),
                    _ => throw new MailParseException($"Parse IBKR Report Fail, Unknown NAV Component: {component}/{amount}")
                };
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
            HtmlDocument doc,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblPosAndMTM_");
            var holdings = new List<Holding>();
            if (rows.Count == 0)
                return holdings;

            var quantityColumn = FindIBKRGroupedColumn(rows, "数量", "当前");
            var currentPriceColumn = FindIBKRGroupedColumn(rows, "价格", "当前");
            var currentValueColumn = FindIBKRGroupedColumn(rows, "市场价值", "当前");
            var precisePrices = ParseIBKRPreciseHoldingPrices(doc, contractInfos);
            var currentGroup = "";
            var currentCurrency = baseCurrency;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                if (currentGroup != "股票" && currentGroup != "债券")
                    continue;
                if (IsIBKRPositionDetailRow(row))
                    continue;

                var quantity = ParseIBKRIntegerQuantityAt(row, quantityColumn, "holding quantity");
                if (quantity == 0)
                    continue;

                var currentPriceText = row[currentPriceColumn];
                var currentPrice = ParseIBKRDecimalAt(row, currentPriceColumn, "holding price");
                var currentValue = ParseIBKRDecimalAt(row, currentValueColumn, "holding value");
                var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                if (precisePrices.TryGetValue(GetIBKRHoldingKey(contract), out var precisePrice))
                {
                    currentPriceText = precisePrice.PriceText;
                    currentPrice = precisePrice.Price;
                }
                holdings.Add(CreateIBKRHolding(account, contract, quantity, currentPrice, currentPriceText, currentValue, currentCurrency, row.Count > 1 ? row[1] : ""));
            }

            return holdings;
        }

        private Dictionary<string, IBKRPreciseHoldingPrice> ParseIBKRPreciseHoldingPrices(
            HtmlDocument doc,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblMtmPerfSumByUnderlying_");
            var prices = new Dictionary<string, IBKRPreciseHoldingPrice>(StringComparer.OrdinalIgnoreCase);
            if (rows.Count == 0)
                return prices;

            var currentPriceColumn = FindIBKRGroupedColumn(rows, "价格", "当前");
            var currentGroup = "";
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out _))
                        continue;
                }

                if (currentGroup != "股票" && currentGroup != "债券")
                    continue;

                var priceText = row[currentPriceColumn];
                if (!TryParseIBKRDecimal(priceText, out var price))
                    continue;

                var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                prices[GetIBKRHoldingKey(contract)] = new IBKRPreciseHoldingPrice(price, priceText);
            }

            return prices;
        }

        private List<Holding> ParseIBKROpenPositionHoldings(
            HtmlDocument doc,
            Account account,
            CurrencyType baseCurrency,
            Dictionary<string, IBKRContractInfo> contractInfos)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblOpenPositions_");
            var holdings = new List<Holding>();
            var currentGroup = "";
            var currentCurrency = baseCurrency;
            foreach (var row in rows)
            {
                if (IsIBKRBlankRow(row) || IsIBKRHeaderRow(row) || IsIBKRTotalRow(row))
                    continue;

                if (TryReadIBKRSingleValue(row, out var singleValue))
                {
                    if (IBKRAssetGroups.Contains(singleValue))
                    {
                        currentGroup = singleValue;
                        continue;
                    }

                    if (TryParseIBKRCurrencyType(singleValue, out var parsedCurrency))
                    {
                        currentCurrency = parsedCurrency;
                        continue;
                    }
                }

                if (currentGroup != "股票" && currentGroup != "债券")
                    continue;

                var quantity = ParseIBKRIntegerQuantityAt(row, 1, "holding quantity");
                if (quantity == 0)
                    continue;

                var currentPriceText = row[5];
                var currentPrice = ParseIBKRDecimalAt(row, 5, "holding price");
                var currentValue = ParseIBKRDecimalAt(row, 6, "holding value");
                var contract = ResolveIBKRContract(row[0], currentGroup, contractInfos);
                holdings.Add(CreateIBKRHolding(account, contract, quantity, currentPrice, currentPriceText, currentValue, currentCurrency, ""));
            }

            return holdings;
        }

        private static Holding CreateIBKRHolding(
            Account account,
            IBKRContractInfo contract,
            int rawQuantity,
            decimal statementPrice,
            string statementPriceText,
            decimal currentValue,
            CurrencyType currency,
            string rowDescription)
        {
            var quantity = NormalizeIBKRHoldingQuantity(contract, rawQuantity);
            var currentPrice = RoundIBKRUnitPrice(statementPrice, GetIBKRUnitPriceDecimals(contract, statementPriceText));
            var roundedCurrentValue = RoundIBKRMoney(currentValue);
            if (!IsIBKRMoneyEqual(roundedCurrentValue, quantity * currentPrice) && quantity != 0)
                currentPrice = RoundIBKRUnitPrice(currentValue / quantity, 7);
            AssertIBKRMoneyEquals(roundedCurrentValue, quantity * currentPrice, $"IBKR holding value {contract.Code}");
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

        private static int GetIBKRUnitPriceDecimals(IBKRContractInfo contract, string statementPriceText)
        {
            var statementDecimals = CountIBKRSignificantDecimalPlaces(statementPriceText);
            var minimumDecimals = contract.HoldingType switch
            {
                HoldingType.UST => 6,
                HoldingType.NASDAQ or HoldingType.ARCA => 2,
                _ => 2
            };
            return Math.Min(7, Math.Max(minimumDecimals, statementDecimals));
        }

        private static decimal RoundIBKRUnitPrice(decimal value, int decimals)
        {
            return Decimal.Round(value, decimals, MidpointRounding.ToEven);
        }

        private static int CountIBKRSignificantDecimalPlaces(string text)
        {
            var normalized = text.Trim().Replace(",", "");
            var decimalPoint = normalized.IndexOf('.');
            if (decimalPoint < 0)
                return 0;

            var decimals = normalized[(decimalPoint + 1)..]
                .TakeWhile(Char.IsDigit)
                .ToArray();
            return new String(decimals).TrimEnd('0').Length;
        }

        private static int NormalizeIBKRHoldingQuantity(IBKRContractInfo contract, int rawQuantity)
        {
            if (contract.HoldingType == HoldingType.UST)
            {
                if (rawQuantity % 100 != 0)
                    throw new MailParseException($"Parse IBKR Report Fail, Invalid UST Quantity: {contract.Code}/{rawQuantity}");

                var ustQuantity = rawQuantity / 100;
                return ustQuantity;
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

            if (group == "债券")
            {
                var matchedBond = contractInfos.Values
                    .Where(item => item.HoldingType == HoldingType.UST && code.Contains(item.Code, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(item => item.Code.Length)
                    .FirstOrDefault();
                if (matchedBond is not null)
                    return matchedBond;

                var normalizedBondCode = ExtractIBKRBondCode(code);
                return new IBKRContractInfo(normalizedBondCode, normalizedBondCode, HoldingType.UST, NormalizeIBKRBondDisplayText(normalizedBondCode, ""));
            }

            if (group == "股票" && IBKRFallbackHoldingTypes.TryGetValue(code, out var holdingType))
                return new IBKRContractInfo(code, code, holdingType, code);

            if (group == "外汇" && TryParseIBKRCurrencyType(code, out _))
                return new IBKRContractInfo(code, code, HoldingType.Cash, code);

            throw new MailParseException($"Parse IBKR Report Fail, Unknown Contract: {rawCode}/{group}");
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

        private static HoldingType ParseIBKRHoldingType(string exchange)
        {
            return exchange.Trim().ToUpperInvariant() switch
            {
                "NASDAQ" => HoldingType.NASDAQ,
                "ISLAND" => HoldingType.NASDAQ,
                "ARCA" => HoldingType.ARCA,
                "NYSEARCA" => HoldingType.ARCA,
                "NYSE ARCA" => HoldingType.ARCA,
                _ => throw new MailParseException($"Parse IBKR Report Fail, Unknown Stock Exchange: {exchange}")
            };
        }

        private static decimal ParseIBKREndingCash(HtmlDocument doc)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblCashReport_");
            var row = rows.FirstOrDefault(row => row.Count > 1 && row[0] == "期末现金")
                ?? throw new MailParseException("Parse IBKR Report Fail, Missing Ending Cash");
            return RoundIBKRMoney(ParseIBKRDecimalAt(row, 1, "ending cash"));
        }

        private static decimal ParseIBKREndingNav(HtmlDocument doc)
        {
            var rows = ReadIBKRSectionFirstTableRows(doc, "tblNAV_");
            var row = rows.FirstOrDefault(row => row.Count > 4 && row[0] == "总数")
                ?? throw new MailParseException("Parse IBKR Report Fail, Missing Ending NAV");
            return RoundIBKRMoney(ParseIBKRDecimalAt(row, 4, "ending NAV"));
        }

        private static decimal ParseIBKRStartingNav(HtmlDocument doc)
        {
            return ParseIBKRNavChangeValues(doc).Start;
        }

        private static decimal ParseIBKRNavChange(HtmlDocument doc)
        {
            var (start, end) = ParseIBKRNavChangeValues(doc);
            return end - start;
        }

        private static (decimal Start, decimal End) ParseIBKRNavChangeValues(HtmlDocument doc)
        {
            var tables = ReadIBKRSectionTables(doc, "tblNAV_");
            if (tables.Count < 2)
                throw new MailParseException("Parse IBKR Report Fail, Missing NAV Change Table");

            decimal? start = null;
            decimal? end = null;
            foreach (var row in tables[1])
            {
                if (row.Count < 2)
                    continue;
                if (row[0] == "开始价值")
                    start = ParseIBKRDecimalAt(row, 1, "starting NAV");
                else if (row[0] == "结束价值")
                    end = ParseIBKRDecimalAt(row, 1, "ending NAV");
            }

            if (start is null || end is null)
                throw new MailParseException("Parse IBKR Report Fail, Invalid NAV Change Table");
            return (RoundIBKRMoney(start.Value), RoundIBKRMoney(end.Value));
        }

        private static decimal ParseIBKRNavChangeComponent(HtmlDocument doc, string componentName)
        {
            var tables = ReadIBKRSectionTables(doc, "tblNAV_");
            if (tables.Count < 2)
                return 0;

            var row = tables[1].FirstOrDefault(row => row.Count > 1 && row[0] == componentName);
            return row is null ? 0 : ParseIBKRDecimalAt(row, 1, $"NAV change {componentName}");
        }

        private static List<List<string>> ReadIBKRSectionFirstTableRows(HtmlDocument doc, string idPrefix)
        {
            var tables = ReadIBKRSectionTables(doc, idPrefix);
            return tables.Count == 0 ? [] : tables[0];
        }

        private static List<List<List<string>>> ReadIBKRSectionTables(HtmlDocument doc, string idPrefix)
        {
            var section = doc.DocumentNode.SelectSingleNode(
                $"//*[starts-with(@id, '{idPrefix}') and contains(concat(' ', normalize-space(@class), ' '), ' sectionContent ')]");
            if (section is null)
                return [];

            return (section.SelectNodes(".//table") ?? Enumerable.Empty<HtmlNode>())
                .Select(ReadIBKRTableRows)
                .ToList();
        }

        private static int FindIBKRGroupedColumn(List<List<string>> rows, string groupName, string columnName)
        {
            if (rows.Count < 2)
                throw new MailParseException($"Parse IBKR Report Fail, Missing Grouped Headers: {groupName}/{columnName}");

            var groupRow = rows[0];
            var columnRow = rows[1];
            for (var i = 0; i < groupRow.Count; i++)
            {
                if (groupRow[i] != groupName)
                    continue;

                var end = groupRow.Count;
                for (var j = i + 1; j < groupRow.Count; j++)
                {
                    if (!String.IsNullOrWhiteSpace(groupRow[j]))
                    {
                        end = j;
                        break;
                    }
                }

                for (var j = i; j < end && j < columnRow.Count; j++)
                {
                    if (columnRow[j] == columnName)
                        return j;
                }
            }

            throw new MailParseException($"Parse IBKR Report Fail, Missing Grouped Column: {groupName}/{columnName}");
        }

        private static List<List<string>> ReadIBKRTableRows(HtmlNode table)
        {
            var rows = new List<List<string>>();
            var rowSpans = new Dictionary<int, IBKRRowSpanCell>();
            var maxColumns = 0;

            foreach (var tr in table.SelectNodes(".//tr") ?? Enumerable.Empty<HtmlNode>())
            {
                var row = new List<string>();
                var col = 0;
                var cells = tr.ChildNodes
                    .Where(node => node.Name.Equals("td", StringComparison.OrdinalIgnoreCase)
                        || node.Name.Equals("th", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var cell in cells)
                {
                    while (rowSpans.ContainsKey(col))
                    {
                        AddIBKRRowSpanValue(row, rowSpans, col);
                        col++;
                    }

                    var colspan = ReadIBKRSpan(cell, "colspan");
                    var rowspan = ReadIBKRSpan(cell, "rowspan");
                    var text = CleanIBKRText(cell.InnerText);

                    for (var i = 0; i < colspan; i++)
                    {
                        var cellText = i == 0 ? text : "";
                        row.Add(cellText);
                        if (rowspan > 1)
                            rowSpans[col] = new IBKRRowSpanCell(cellText, rowspan - 1);
                        col++;
                    }
                }

                while (rowSpans.ContainsKey(col))
                {
                    AddIBKRRowSpanValue(row, rowSpans, col);
                    col++;
                }

                maxColumns = Math.Max(maxColumns, row.Count);
                rows.Add(row);
            }

            foreach (var row in rows)
            {
                while (row.Count < maxColumns)
                    row.Add("");
            }

            return rows.Where(row => row.Any(cell => !String.IsNullOrWhiteSpace(cell))).ToList();
        }

        private static void AddIBKRRowSpanValue(List<string> row, Dictionary<int, IBKRRowSpanCell> rowSpans, int col)
        {
            var rowSpan = rowSpans[col];
            row.Add(rowSpan.Text);
            rowSpan.RemainingRows--;
            if (rowSpan.RemainingRows <= 0)
                rowSpans.Remove(col);
        }

        private static int ReadIBKRSpan(HtmlNode cell, string attributeName)
        {
            var value = cell.GetAttributeValue(attributeName, "1");
            return Int32.TryParse(value, out var span) && span > 0 ? span : 1;
        }

        private static string CleanIBKRText(string text)
        {
            var decoded = (HtmlEntity.DeEntitize(text) ?? "").Replace('\u00a0', ' ');
            return Regex.Replace(decoded, @"\s+", " ").Trim();
        }

        private static bool IsIBKRBlankRow(List<string> row)
        {
            return row.All(String.IsNullOrWhiteSpace);
        }

        private static bool IsIBKRHeaderRow(List<string> row)
        {
            return row.Count > 0 && (row[0] == "代码" || row[0] == "日期" || row[0] == "起息日" || row[0] == "");
        }

        private static bool IsIBKRTotalRow(List<string> row)
        {
            return row.Count > 0
                && (row[0].StartsWith("总数", StringComparison.Ordinal)
                    || row[0].StartsWith("总计", StringComparison.Ordinal)
                    || row[0].StartsWith("合计", StringComparison.Ordinal));
        }

        private static bool IsIBKRPositionDetailRow(List<string> row)
        {
            return TryGetIBKRPositionDetail(row, out _, out _);
        }

        private static bool TryGetIBKRPositionDetail(List<string> row, out DateTime date, out string action)
        {
            date = default;
            action = "";
            if (row.Count > 3
                && TryParseIBKRDate(row[2], out date)
                && !String.IsNullOrWhiteSpace(row[3]))
            {
                action = row[3];
                return true;
            }

            if (row.Count > 2
                && TryParseIBKRDate(row[1], out date)
                && !String.IsNullOrWhiteSpace(row[2]))
            {
                action = row[2];
                return true;
            }

            return false;
        }

        private static bool TryReadIBKRSingleValue(List<string> row, out string value)
        {
            var nonEmpty = row.Where(cell => !String.IsNullOrWhiteSpace(cell)).ToList();
            if (nonEmpty.Count == 1)
            {
                value = nonEmpty[0];
                return true;
            }

            value = "";
            return false;
        }

        private static decimal ParseIBKRDecimalAt(List<string> row, int index, string context)
        {
            if (index >= row.Count || !TryParseIBKRDecimal(row[index], out var value))
                throw new MailParseException($"Parse IBKR Report Fail, Invalid {context}: {String.Join(",", row)}");
            return value;
        }

        private static int ParseIBKRIntegerQuantityAt(List<string> row, int index, string context)
        {
            var value = ParseIBKRDecimalAt(row, index, context);
            if (value != Decimal.Truncate(value))
                throw new MailParseException($"Parse IBKR Report Fail, Non-integer {context}: {String.Join(",", row)}");

            return checked((int)value);
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
            return TryParseIBKRDecimal(text, out var value) ? value : 0;
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

        private static void AssertIBKRMoneyEquals(decimal expected, decimal actual, string context)
        {
            if (!IsIBKRMoneyEqual(expected, actual))
                throw new MailParseException($"Parse IBKR Report Fail, {context} mismatch: expected {expected}, got {actual}");
        }

        private static bool IsIBKRMoneyEqual(decimal expected, decimal actual)
        {
            return RoundIBKRMoney(expected) == RoundIBKRMoney(actual);
        }

        private static decimal RoundIBKRMoney(decimal value)
        {
            return Currency.RoundMoney(value);
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
                    || attachmentInfo.ReportDate.Date != reportDate.Date)
                    continue;

                using var memory = new MemoryStream();
                mimePart.Content.DecodeTo(memory);
                attachments.Add(new InMemoryIBKRReportAttachment(
                    fileName,
                    attachmentInfo.ReportType,
                    attachmentInfo.ReportId,
                    attachmentInfo.ReportDate,
                    memory.ToArray()));
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
                    && attachmentInfo.ReportDate.Date == reportDate.Date;
            });
        }

        private static bool IsDailyMyBookReportType(string reportType)
        {
            return String.Equals(IBKRDailyReportType, reportType, StringComparison.OrdinalIgnoreCase);
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

        private static string GetAttachmentFileName(MimeEntity attachment)
        {
            return attachment.ContentDisposition?.FileName
                ?? attachment.ContentType.Name
                ?? "";
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
            List<AccountBalance> BeginningAccountBalances);

        private sealed record IBKRContractInfo(string Code, string Description, HoldingType HoldingType, string DisplayText);

        private sealed record IBKRPreciseHoldingPrice(decimal Price, string PriceText);

        private sealed record IBKRMtmTotals(bool HasData, decimal Holding, decimal Transaction, decimal Commission, decimal Other);

        private sealed record IBKRTransactionTotals(
            bool HasData,
            decimal Proceeds,
            decimal BuyProceeds,
            decimal SellProceeds,
            decimal Commission,
            decimal Mtm);

        private sealed record IBKRCommissionTotals(bool HasData, decimal Total);

        private sealed record IBKRInterestTotals(bool HasData, decimal Total);

        private sealed record IBKRCashTotals(decimal Commission, decimal TradeBuy, decimal TradeSell, decimal Interest);

        private sealed record IBKRReportAttachmentInfo(string ReportType, string ReportId, DateTime ReportDate);

        private sealed record InMemoryIBKRReportAttachment(
            string FileName,
            string ReportType,
            string ReportId,
            DateTime ReportDate,
            byte[] Content);

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
                amount = new Currency(RoundIBKRMoney(amount.v), amount.t);
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
                    Source = LimitIBKRRecordText($"IBKR HTML {sourceName}/{source}"),
                    Reason = LimitIBKRRecordText(reason)
                };
                record.CopyFrom(amount);
                Records.Add(record);

                if (affectsNetAsset)
                    NetAssetChangeTotal += amount.v;
            }
        }

        private sealed class IBKRRowSpanCell
        {
            public IBKRRowSpanCell(string text, int remainingRows)
            {
                Text = text;
                RemainingRows = remainingRows;
            }

            public string Text { get; }
            public int RemainingRows { get; set; }
        }
    }
}
