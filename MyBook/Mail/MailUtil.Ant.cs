using System.Globalization;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    private const string AntSender = "hk_antbank_service@notify.antbank.hk";

    public Task FetchAntMessages() => FetchHKBankMessages("Ant", AntSender, StatementImportProvider.AntMail,
        IsAntTransactionSubject, ParseAntMessage);

    private static bool IsAntTransactionSubject(string subject)
    {
        subject = subject.Trim();
        if (subject is "款項已存入您的賬戶" or "支付成功") return true;
        if (subject is "綁定成功" or "授權成功" or "您的位置已更改" or "螞蟻銀行通知"
            or "註冊成功-請激活賬戶（Application Successful-Account Activation Needed）"
            or "【螞蟻銀行香港】電郵地址驗證碼"
            or "eM+ HKD Savings Account Interest Rate Adjustment 餘額+港幣活期賬戶年利率調整") return false;
        throw new MailParseException("Unsupported Ant mail subject.");
    }

    private static Record ParseAntMessage(MimeMessage message)
    {
        if (!IsFrom(message, AntSender) || !IsAntTransactionSubject(message.Subject))
            throw new MailParseException("Not an Ant transaction notification.");
        var text = HKBankMailText(message);
        var receipt = message.Subject.Trim() == "款項已存入您的賬戶";
        var english = MatchHKBankMail("Ant", text, receipt
            ? @"A transfer of " + HKBankMoneyPattern + @" has been credited to your account\."
            : @"We have debited " + HKBankMoneyPattern + @" on (?<date>\d{2}/\d{2} \d{2}:\d{2}) from your Savings account according to your instruction\.");
        var chinese = MatchHKBankMail("Ant", text, receipt
            ? @"您收到一筆" + HKBankMoneyPattern + "的款項，已存入您的賬戶。"
            : @"本行已按您於(?<date>\d{2}月\d{2}日\d{2}:\d{2})的指示，成功從活期賬戶扣取" + HKBankMoneyPattern + "。");
        var amount = HKBankAmount(english);
        if (!amount.Equals(HKBankAmount(chinese))) throw new MailParseException("Ant bilingual amounts disagree.");
        var date = GetMailDateTime(message);
        if (!receipt)
        {
            // The payment time omits the year; use the latest occurrence preceding the mail (including New Year).
            var bankMailDate = message.Date.ToOffset(TimeSpan.FromHours(8)).DateTime;
            var bankDate = DateTime.ParseExact($"{bankMailDate.Year}/{english.Groups["date"].Value}",
                "yyyy/dd/MM HH:mm", CultureInfo.InvariantCulture);
            if (bankDate > bankMailDate) bankDate = bankDate.AddYears(-1);
            if (bankMailDate - bankDate > TimeSpan.FromDays(2)
                || bankDate.ToString("MM月dd日HH:mm", CultureInfo.InvariantCulture) != chinese.Groups["date"].Value)
                throw new MailParseException("Ant payment date is inconsistent with the mail.");
            date = new DateTimeOffset(bankDate, TimeSpan.FromHours(8)).LocalDateTime;
            amount.v = -amount.v;
        }
        var record = HKBankRecord("Ant", message, amount, date, receipt ? "转入" : "支付");
        // Deposit notices contain neither a transaction time nor a counterparty.
        if (receipt) record.Source += "; timeSource=mail-date";
        return record;
    }
}
