using System.Text.RegularExpressions;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    private const string EleSender = "appsendmail@notification.elebank.com";

    public Task FetchEleMessages() => FetchHKBankMessages("ELE", EleSender, StatementImportProvider.EleMail,
        IsEleTransactionSubject, ParseEleMessage);

    private static bool IsEleTransactionSubject(string subject)
    {
        subject = subject.Trim();
        if (subject is "大象银行EleBank - 收款成功通知" or "大象银行交易通知") return true;
        if (subject is "大象银行EleBank - 开户通知" or "大象银行EleBank - 开通转数快转账功能电邮验证码验证"
            or "大象银行EleBank - 电邮验证成功" or "大象银行EleBank - 电邮验证码"
            or "大象银行实体卡申领通知" or "大象银行开卡通知" or "大象银行绑卡通知") return false;
        throw new MailParseException("Unsupported ELE mail subject.");
    }

    private static Record ParseEleMessage(MimeMessage message)
    {
        if (!IsFrom(message, EleSender) || !IsEleTransactionSubject(message.Subject))
            throw new MailParseException("Not an ELE transaction notification.");
        var text = HKBankMailText(message);
        if (Regex.IsMatch(text, @"交易已取消|退款|退還|退还|\brefund(?:ed)?\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            throw new MailParseException("ELE refunds and cancelled transactions are not supported.");
        const string datePattern = @"(?<date>\d{4}/\d{2}/\d{2} \d{2}:\d{2})";
        Match detail;
        Currency amount;
        string reason;
        var receipt = message.Subject.Trim() == "大象银行EleBank - 收款成功通知";
        if (receipt)
        {
            detail = MatchHKBankMail("ELE", text, @"来自(?<party>.+?)\s+" + HKBankMoneyPattern
                + "的转账已于" + datePattern + @"存入你的账户（尾数(?<suffix>\d+)）。");
            amount = HKBankAmount(detail);
            reason = "转入";
        }
        else
        {
            detail = MatchHKBankMail("ELE", text, @"你的借记卡（卡结尾(?<suffix>\d{4})）于\s*" + datePattern
                + @"，在\s+(?<party>.+?)\s+消费\s*" + HKBankMoneyPattern + @"。如有疑问");
            reason = "消费";
            amount = HKBankAmount(detail);
            amount.v = -amount.v;
        }
        var record = HKBankRecord("ELE", message, amount,
            HKBankLocalDate(detail.Groups["date"].Value, "yyyy/MM/dd HH:mm"), reason, detail.Groups["party"].Value);
        var suffix = detail.Groups["suffix"].Value;
        record.Source += receipt ? $"; ownAccountSuffix={suffix}" : $"; ownCardSuffix={suffix}";
        return record;
    }

}
