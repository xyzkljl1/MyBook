using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MimeKit;

namespace MyBook
{
    // ICBC 邮件账单获取与工行信用卡交易解析逻辑。
    partial class MailUtil
    {
        private const string ICBCAccountType = "ICBC";
        private const StatementImportProvider ICBCProvider = StatementImportProvider.ICBCBillMail;

        private Account GetICBCCardAccount(string name)
        {
            return database.GetAccountByTypeAndId(ICBCAccountType, name.Substring(0, 4));
        }

        private Account GetICBCPostingAccount(string name)
        {
            return database.GetPostingAccount(GetICBCCardAccount(name));
        }

        public async Task FetchICBCBills()
        {
            await FetchMonthlyStatements(ICBCProvider, "ICBC bill", FetchICBCBill);
        }

        // 工行对账单，按卡号区分用途。
        public async Task<bool> FetchICBCBill(DateTime date)
        {
            var reportMonth = FirstDayOfMonth(date);
            var message = await SearchBill("webmaster@icbc.com.cn", "中国工商银行客户对账单", reportMonth, IsICBCInlineBillMessage);
            if (message is null)
                return false;

            var billText = message.HtmlBody ?? message.TextBody ?? "";
            if (String.IsNullOrEmpty(billText))
                return false;

            Records records = new();
            try
            {
                var statementEndDate = ParseICBCStatementEndDate(billText);
                var statementKey = statementEndDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                if (database.IsStatementKeyImported(ICBCProvider, statementKey))
                    return true;

                var tables = FormUtil.ReadFromHTML(billText);
                if (!IsICBCInlineBillTables(tables))
                    throw new MailParseException("Parse ICBC Bill Fail, Invalid Tables");
                var beginningAccountBalances = ParseICBCAccountBalances(tables[1], 1);
                var accountBalances = ParseICBCAccountBalances(tables[1], 4);
                // 只从汇总表登记卡号；交易明细里的卡号只用于该交易本身，不能作为卡号来源。
                var internalCardNos = ParseICBCInternalCardNos(tables[1]);

                records.AddRange(ParseICBCTransactionRecords(tables[2])); // 人民币交易明细。
                records.AddRange(ParseICBCTransactionRecords(tables[3])); // 外币交易明细。
                database.SaveStatementRecordsOnce(
                    ICBCProvider,
                    GetMailDate(message),
                    records,
                    accountBalances,
                    statementKey,
                    beginningAccountBalances,
                    internalCardNos: internalCardNos,
                    afterSaveInTransaction: statementImportId =>
                    {
                        OffsetMatchedICBCRefundRecords(statementImportId);
                    });
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine($"parse mail fail :{e.Message}");
                throw;
            }
        }

        private static DateTime ParseICBCStatementEndDate(string billText)
        {
            if (TryParseICBCStatementEndDate(billText, out var statementEndDate))
                return statementEndDate;

            throw new MailParseException("Parse ICBC Bill Fail, Missing Statement Period");
        }

        private static bool TryParseICBCStatementEndDate(string billText, out DateTime statementEndDate)
        {
            statementEndDate = default;
            var text = NormalizeMailText(Regex.Replace(System.Net.WebUtility.HtmlDecode(billText), "<[^>]+>", " "));
            var match = Regex.Match(text, @"账单周期\s*\d{4}年\d{1,2}月\d{1,2}日\s*[—\-－~至到]\s*(?<end>\d{4}年\d{1,2}月\d{1,2}日)");
            if (!match.Success)
                return false;

            statementEndDate = DateTime.ParseExact(match.Groups["end"].Value, "yyyy年M月d日", CultureInfo.InvariantCulture);
            return true;
        }

        private static bool IsICBCInlineBillMessage(MimeMessage message)
        {
            var billText = message.HtmlBody ?? message.TextBody ?? "";
            if (String.IsNullOrWhiteSpace(billText) || !TryParseICBCStatementEndDate(billText, out _))
                return false;

            try
            {
                return IsICBCInlineBillTables(FormUtil.ReadFromHTML(billText));
            }
            catch
            {
                return false;
            }
        }

        private static bool IsICBCInlineBillTables(List<FormUtil.FormTable> tables)
        {
            return tables.Count >= 4
                && tables[0].Title == "需 还 款 明 细"
                && tables[1].Title == "本 期 交 易 汇 总"
                && tables[2].Title == "人民币(本位币) 交 易 明 细"
                && tables[3].Title == "外 币 交 易 明 细";
        }

        private List<AccountBalance> ParseICBCAccountBalances(FormUtil.FormTable table, int balanceColumn)
        {
            var accountBalances = new List<AccountBalance>();
            foreach (var line in table.Rows)
            {
                if (line.Count <= balanceColumn || line[0] == "合计")
                    continue;

                var balance = Currency.Parse(line[balanceColumn]);
                var account = GetICBCPostingAccount(line[0]);
                accountBalances.Add(new AccountBalance(account, balance));
            }

            return accountBalances;
        }

        private List<AccountInternalId> ParseICBCInternalCardNos(params FormUtil.FormTable[] tables)
        {
            var result = new List<AccountInternalId>();
            foreach (var table in tables)
            {
                foreach (var line in table.Rows)
                {
                if (line.Count == 0 || line[0] == "鍚堣")
                    continue;

                    var cardNo = Regex.Replace(line[0], @"\D", "");
                    if (String.IsNullOrWhiteSpace(cardNo))
                        continue;

                    result.Add(new AccountInternalId
                    {
                        Account = GetICBCCardAccount(cardNo),
                        cardNo = cardNo,
                        sourceText = $"ICBC table={table.Title}; row={String.Join(" | ", line)}"
                    });
                }
            }

            return result;
        }

        private Records ParseICBCTransactionRecords(FormUtil.FormTable table)
        {
            if (table.Headers.Count != 7
                || table.Headers[0] != "卡号后四位"
                || table.Headers[1] != "交易日"
                || table.Headers[3] != "交易类型"
                || table.Headers[4] != "商户名称/城市"
                || table.Headers[6] != "记账金额/币种")
                throw new MailParseException("Parse ICBC Bill Fail, Invalid Headers");

            Records records = new();
            foreach (var line in table.Rows)
            {
                var record = new Record();
                var postingCurrency = Currency.Parse(line[6]);
                var cardAccount = GetICBCCardAccount(line[0]);
                record.updateTime = DateTime.Now;
                record.Account = database.GetPostingAccount(cardAccount);
                record.date = DateTime.ParseExact(line[1], "yyyy-MM-dd", CultureInfo.InvariantCulture);
                record.postingDate = DateTime.ParseExact(line[2], "yyyy-MM-dd", CultureInfo.InvariantCulture);
                record.Source = $"ICBC对账单邮件{table.Title}/{DateTime.Now}/{string.Join(",", line)}";
                record.DestAccount = line[4];
                record.CopyFrom(postingCurrency);
                record.DescCurrency = ApplyICBCOriginalCurrencySign(Currency.Parse(line[5]), record);
                var internalCounterparty = database.FindAccountByInternalCardNoText(
                    null,
                    $"ICBC import table={table.Title}; transactionDate={line[1]}; postingDate={line[2]}; card={line[0]}; type={line[3]}; row={String.Join(" | ", line)}",
                    record.DestAccount);
                if (internalCounterparty is not null && !IsSameAccount(database.GetPostingAccount(internalCounterparty), record.Account))
                {
                    record.DestAccount = internalCounterparty.name;
                    record.isInternal = true;
                }

                var transactionType = line[3];
                if (IsICBCExpenseTransactionType(transactionType))
                {
                    if (record.v >= 0)
                        throw new MailParseException($"Parse ICBC Bill Fail, Invalid Expense: {transactionType}");
                    record.Reason = GetICBCExpenseReason(transactionType, cardAccount);
                    records.Add(record);
                }
                else if (IsICBCRefundTransactionType(transactionType))
                {
                    if (record.v <= 0)
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid In");
                    record.Reason = cardAccount.desc; // 工行按交易明细中的卡区分用途，副卡记录仍入主卡账。
                    records.Add(record);
                }
                else if (IsICBCPaymentTransactionType(transactionType))
                {
                    if (transactionType == "转账" && record.v <= 0)
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid Transfer");
                    record.isInternal = true;
                    record.Reason = "信用卡还款";
                    records.Add(record);
                }
                else if (transactionType == "年费减免")
                {
                    if (record.v != 0)
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid Annual Fee Waiver");
                }
                else
                {
                    throw new MailParseException($"Parse ICBC Bill Fail, Unknown Transaction Type: {transactionType}");
                }
            }

            return records;
        }

        private static bool IsICBCExpenseTransactionType(string transactionType)
        {
            return transactionType is "消费"
                or "跨行消费"
                or "境外消费"
                or "缴费"
                or "直接分期扣款"
                or "分期付款到期扣收"
                or "境外取现"
                or "跨境手续费"
                or "透支利息";
        }

        private static string GetICBCExpenseReason(string transactionType, Account cardAccount)
        {
            return transactionType switch
            {
                "跨境手续费" => "手续费",
                "透支利息" => "利息",
                _ => cardAccount.desc // 工行按交易明细中的卡区分用途，副卡记录仍入主卡账。
            };
        }

        private static bool IsICBCRefundTransactionType(string transactionType)
        {
            return transactionType is "退款"
                or "境外退货"
                or "消费返利";
        }

        private static bool IsICBCPaymentTransactionType(string transactionType)
        {
            return transactionType is "人民币自动转账还款"
                or "自动购汇还款"
                or "转账";
        }

        private int OffsetMatchedICBCRefundRecords(int targetStatementImportId)
        {
            var refunds = database.GetRecordsByStatementImport(targetStatementImportId)
                .Where(record => !record.isRefundMatched)
                .Where(IsICBCRefundRecord)
                .OrderBy(record => record.date)
                .ThenBy(record => record.Id)
                .ToList();
            if (refunds.Count == 0)
                return 0;

            var minDate = refunds.Min(record => record.date.Date.AddMonths(-2));
            var maxDate = refunds.Max(record => record.date);
            var expenses = database.GetStatementRecords(ICBCProvider, minDate, maxDate)
                .Where(record => !DatabaseUtil.IsInitializationRecord(record)
                    && !record.isRefundMatched
                    && record.v < 0)
                .Where(IsICBCExpenseRecord)
                .OrderBy(record => record.date)
                .ThenBy(record => record.Id)
                .ToList();
            var matchedRecordIds = new HashSet<int>();
            var pairCount = 0;

            foreach (var refund in refunds)
            {
                if (matchedRecordIds.Contains(refund.Id))
                    continue;

                var matchStart = refund.date.Date.AddMonths(-2);
                var expense = expenses
                    .Where(record => !matchedRecordIds.Contains(record.Id)
                        && record.date >= matchStart
                        && (record.date < refund.date
                            || (record.date == refund.date && record.Id < refund.Id))
                        && IsICBCRefundExpenseMatch(refund, record))
                    .OrderByDescending(record => record.date)
                    .ThenByDescending(record => record.Id)
                    .FirstOrDefault();
                if (expense is null)
                    continue;

                matchedRecordIds.Add(refund.Id);
                matchedRecordIds.Add(expense.Id);
                pairCount++;
            }

            if (matchedRecordIds.Count == 0)
                return 0;

            database.MarkRecordsAsRefundMatched(refunds
                .Concat(expenses)
                .Where(record => matchedRecordIds.Contains(record.Id)));
            Console.WriteLine($"matched ICBC refund records: {pairCount} pairs");
            return pairCount;
        }

        private static bool IsICBCRefundRecord(Record record)
        {
            return record.v > 0
                && (record.Source.Contains("退款", StringComparison.Ordinal)
                    || record.Source.Contains("境外退货", StringComparison.Ordinal));
        }

        private static bool IsICBCExpenseRecord(Record record)
        {
            return record.v < 0
                && (record.Source.Contains("消费", StringComparison.Ordinal)
                    || record.Source.Contains("缴费", StringComparison.Ordinal));
        }

        private static bool IsICBCRefundExpenseMatch(Record refund, Record expense)
        {
            return refund._account_Id == expense._account_Id
                && refund.DestAccount == expense.DestAccount
                && refund.DescCurrency is not null
                && IsOppositeICBCOriginalCurrency(refund.DescCurrency, expense.DescCurrency);
        }

        private static Currency ApplyICBCOriginalCurrencySign(Currency originalCurrency, Currency postingCurrency)
        {
            if (postingCurrency.v == 0 || originalCurrency.v == 0)
                return originalCurrency;

            return new Currency(Math.Abs(originalCurrency.v) * Math.Sign(postingCurrency.v), originalCurrency.t);
        }

        private static bool IsOppositeICBCOriginalCurrency(Currency? left, Currency? right)
        {
            return left is not null
                && right is not null
                && left.t == right.t
                && left.v + right.v == 0;
        }
    }
}
