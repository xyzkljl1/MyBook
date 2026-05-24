using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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
            var message = await SearchBill("webmaster@icbc.com.cn", "中国工商银行客户对账单", reportMonth, null);
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
                if (tables.Count < 4
                    || tables[0].Title != "需 还 款 明 细"
                    || tables[1].Title != "本 期 交 易 汇 总"
                    || tables[2].Title != "人民币(本位币) 交 易 明 细"
                    || tables[3].Title != "外 币 交 易 明 细")
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
                    afterSaveInTransaction: statementImportId =>
                    {
                        database.EnsureAccountInternalCardNos(internalCardNos);
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
            var text = NormalizeMailText(Regex.Replace(System.Net.WebUtility.HtmlDecode(billText), "<[^>]+>", " "));
            var match = Regex.Match(text, @"账单周期\s*\d{4}年\d{1,2}月\d{1,2}日\s*[—\-－~至到]\s*(?<end>\d{4}年\d{1,2}月\d{1,2}日)");
            if (!match.Success)
                throw new MailParseException("Parse ICBC Bill Fail, Missing Statement Period");

            return DateTime.ParseExact(match.Groups["end"].Value, "yyyy年M月d日", CultureInfo.InvariantCulture);
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
                record.Source = $"ICBC对账单邮件{table.Title}/{DateTime.Now}/{string.Join(",", line)}";
                record.DestAccount = line[4];
                record.DescCurrency = Currency.Parse(line[5]);
                record.CopyFrom(postingCurrency);
                var internalCounterparty = database.FindAccountByInternalCardNoText(
                    null,
                    $"ICBC import table={table.Title}; date={line[1]}; card={line[0]}; type={line[3]}; row={String.Join(" | ", line)}",
                    record.DestAccount);
                if (internalCounterparty is not null && !IsSameAccount(database.GetPostingAccount(internalCounterparty), record.Account))
                {
                    record.DestAccount = internalCounterparty.name;
                    record.isInternal = true;
                }

                if (line[3] == "消费" || line[3] == "跨行消费" || line[3] == "境外消费" || line[3] == "缴费")
                {
                    record.Reason = cardAccount.desc; // 工行按交易明细中的卡区分用途，副卡记录仍入主卡账。
                    records.Add(record);
                }
                else if (line[3] == "退款" || line[3] == "境外退货")
                {
                    if (record.v <= 0)
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid In");
                    // 副卡消费产生的退款仍会显示在副卡卡号下；在同一个月内向前搜索对应的消费，尝试消除。
                    // 比较 DescCurrency 因为退款是按交易金额退的，IsSameAccount 则保证不会跨入账户匹配。
                    Record? destRecord = records.FindLast(destRecord =>
                        destRecord.DestAccount == record.DestAccount && destRecord.v < 0
                        && IsSameAccount(destRecord.Account, record.Account) && destRecord.DescCurrency == record.DescCurrency);
                    if (destRecord is not null)
                    {
                        records.Remove(destRecord);
                        if (destRecord.t != record.t)
                            throw new MailParseException("Parse ICBC Bill Fail, Refund Currency Mismatch");

                        record.v += destRecord.v;
                        if (record.DescCurrency is not null)
                            record.DescCurrency = new Currency(0, record.DescCurrency.t);
                        if (record.v != 0)
                        {
                            record.Reason = "退款汇率差异";
                            records.Add(record);
                        }
                    }
                    else
                    {
                        record.Reason = cardAccount.desc; // 工行按交易明细中的卡区分用途，副卡记录仍入主卡账。
                        records.Add(record);
                    }
                }
                else if (line[3] == "人民币自动转账还款" || line[3] == "自动购汇还款" || line[3] == "转账")
                {
                    if (line[3] == "转账" && record.v <= 0)
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid Transfer");
                    record.isInternal = true;
                    record.Reason = "信用卡还款";
                    records.Add(record);
                }
                else if (line[3] == "年费减免")
                {
                    if (record.v != 0)
                        throw new MailParseException("Parse ICBC Bill Fail, Invalid Annual Fee Waiver");
                }
                else
                {
                    throw new MailParseException($"Parse ICBC Bill Fail, Unknown Transaction Type: {line[3]}");
                }
            }

            return records;
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
                .Where(record => !record.isRefundMatched
                    && record._statementImport_Id != targetStatementImportId
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
                        && record.date < refund.date
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
                && refund.DescCurrency == expense.DescCurrency;
        }
    }
}
