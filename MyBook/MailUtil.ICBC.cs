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
            var month = GetNextMonthlyStatementDate(ICBCProvider);
            var currentMonth = FirstDayOfMonth(DateTime.Today);
            while (month <= currentMonth)
            {
                var imported = await FetchICBCBill(month);
                if (!imported)
                {
                    if (DateTime.Today >= month.AddMonths(1))
                        throw new InvalidOperationException($"Missing ICBC bill for {month:yyyy-MM}");
                    return;
                }

                month = month.AddMonths(1);
            }
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

                records.AddRange(ParseICBCTransactionRecords(tables[2])); // 人民币交易明细。
                records.AddRange(ParseICBCTransactionRecords(tables[3])); // 外币交易明细。
                database.SaveStatementRecordsOnce(
                    ICBCProvider,
                    GetMailDate(message),
                    records,
                    accountBalances,
                    statementKey,
                    beginningAccountBalances);
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
                        records.Remove(destRecord);
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
                else
                {
                    throw new MailParseException($"Parse ICBC Bill Fail, Unknown Transaction Type: {line[3]}");
                }
            }

            return records;
        }
    }
}
