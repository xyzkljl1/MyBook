using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MailKit.Search;
using MimeKit;

namespace MyBook;

partial class MailUtil
{
    private const string ZASender = "notification@service.bank.za.group";
    private const StatementImportProvider ZAProvider = StatementImportProvider.ZAMail;
    private const string ZAMoneyPattern = @"(?<currency>HKD|USD|CNY|RMB)\s+(?<amount>(?:\d+|\d{1,3}(?:,\d{3})+)\.\d{2})(?![\d.])";
    private const string ZADatePattern = @"(?<date>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})";

    public async Task FetchZAMessages()
    {
        var imports = database.GetStatementImports(ZAProvider);
        var checkpoint = imports.Where(item => item.statementKey == "")
            .Select(item => (DateTime?)item.time).SingleOrDefault()
            ?? throw new InvalidOperationException("Missing ZA mail checkpoint.");
        var since = imports.Where(item => item.statementKey.StartsWith("ZA-mail-", StringComparison.Ordinal))
            .Select(item => item.time).Append(checkpoint).Max();
        var searchDate = since.ToUniversalTime().AddHours(-14).Date;
        await RunWithMailSessionScope(async () =>
        {
            var messages = await SearchMessagesFromMailbox(CreateYahooMailbox() with { Proxy = null }, "ZA transactions",
                SearchQuery.FromContains(ZASender).And(SearchQuery.SentSince(searchDate)),
                summary => SummaryIsFrom(summary, ZASender) && GetSummaryDateTime(summary) >= since
                    && IsZATransactionSubject(summary.Envelope?.Subject ?? "")
                    && (String.IsNullOrWhiteSpace(summary.Envelope?.MessageId)
                        || !database.IsStatementKeyImported(ZAProvider, ZAMessageKey(summary.Envelope.MessageId))),
                message => IsFrom(message, ZASender), GetMailDateTime).ConfigureAwait(false);
            foreach (var message in messages) ImportZAMessage(message);
        }).ConfigureAwait(false);
    }

    private static bool IsZATransactionSubject(string subject)
    {
        subject = subject.Trim();
        if (Regex.IsMatch(subject, @"^你已收到\s*" + ZAMoneyPattern + "$")) return true;
        if (Regex.IsMatch(subject, @"^(?:\u2705\uFE0F?\s*)?你已转出\s*" + ZAMoneyPattern + "$")) return true;
        if (Regex.IsMatch(subject, @"^你有一笔\s+" + ZAMoneyPattern + @"\s+的签账$")) return true;
        if (Regex.IsMatch(subject, @"^(?:\S+\s+)?交易失败$")
            || subject is "成功注册转数快及设置为预设收款银行" or "更新转数快预设收款账户"
                or "修改转出限额" or "更改转出限额" or "新增登记收款人") return false;
        if (Regex.IsMatch(subject, @"转入|转出|收到|签账|消费|退款|扣款|入账|利息|交易|payment|transfer|refund|interest|transaction|purchase",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            throw new MailParseException("Unsupported ZA transaction notification subject.");
        return false;
    }

    private void ImportZAMessage(MimeMessage message)
    {
        if (!IsFrom(message, ZASender)) throw new MailParseException("Invalid ZA sender.");
        if (!IsZATransactionSubject(message.Subject)) return;
        var key = ZAMessageKey(message.MessageId);
        if (database.IsStatementKeyImported(ZAProvider, key)) return;
        var record = ParseZAMessage(message, out var cardTail);
        var account = GetZAAccount(cardTail);
        if (account.isCredit || !account.relativeBalance)
            throw new InvalidOperationException("ZA mail requires a non-credit relative-balance account.");
        record.Account = account;
        record._account_Id = account.Id;
        if (record.Reason is "转入" or "转出")
            DatabaseUtil.ApplyTransferCounterparty(record,
                database.ResolveTransferCounterparty(null, [record.DestAccount], record.DestAccount));
        database.SaveStatementRecordsOnce(ZAProvider, GetMailDateTime(message), [record], statementKey: key);
        Console.WriteLine("Imported ZA transaction: records=1");
    }

    private Account GetZAAccount(string cardTail)
    {
        if (cardTail.Length > 0) return database.GetAccountByTypeAndId("ZA", cardTail);
        var accounts = database.GetAccountsByNamePrefix("ZA_");
        if (accounts.Count != 1)
            throw new InvalidOperationException("ZA transfer mail has no card identifier; exactly one ZA account is required.");
        return accounts[0];
    }

    private static Record ParseZAMessage(MimeMessage message, out string cardTail)
    {
        if (!IsFrom(message, ZASender)) throw new MailParseException("Invalid ZA sender.");
        if (!IsZATransactionSubject(message.Subject)) throw new MailParseException("Not a ZA transaction notification.");
        var document = new HtmlDocument();
        document.LoadHtml(message.HtmlBody ?? "");
        foreach (var node in document.DocumentNode.SelectNodes("//style|//script")?.ToList() ?? [])
            node.Remove();
        var text = NormalizeMailText(WebUtility.HtmlDecode(message.TextBody ?? document.DocumentNode.InnerText));
        Match detail;
        string reason;
        var sign = -1;
        if (message.Subject.Trim().StartsWith("你已收到", StringComparison.Ordinal))
        {
            detail = MatchZAMail(text, @"^收到一笔转账\s+你好.*?你已于\s+" + ZADatePattern
                + @"\s+收到以下转账。\s+金额：\s*" + ZAMoneyPattern
                + @"\s+付款人：(?<party>.+?)\s+交易类型：转入\s+你可到 ZA Bank App");
            reason = "转入";
            sign = 1;
        }
        else if (message.Subject.Contains("你已转出", StringComparison.Ordinal))
        {
            detail = MatchZAMail(text, @"^完成一笔转出\s+你好.*?你已于\s+" + ZADatePattern
                + @"\s+完成以下交易。\s+转出金额：\s*" + ZAMoneyPattern
                + @"\s+收款人：(?<party>.+?)\s+交易类型：(?:转出|手机号转出|Email转出)\s+你可到 ZA Bank App");
            reason = "转出";
        }
        else
        {
            detail = MatchZAMail(text, @"^完成一笔消费\s+你好.*?你于\s+" + ZADatePattern
                + @"\s+使用 ZA Card \((?<card>\d{4})\) 完成以下一笔消费\s+消费金额：\s*" + ZAMoneyPattern
                + @"\s+消费商户：(?<party>.+?)\s+请到 ZA Bank App");
            reason = "消费";
        }
        cardTail = detail.Groups["card"].Value;
        var amount = ParseZAMoney(detail);
        var subjectAmount = ParseZAMoney(MatchZAMail(message.Subject, ZAMoneyPattern));
        if (amount.t != subjectAmount.t || amount.v != subjectAmount.v)
            throw new MailParseException("ZA subject/body amount mismatch.");
        amount.v *= sign;
        var bankDate = DateTime.ParseExact(detail.Groups["date"].Value, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        var date = new DateTimeOffset(DateTime.SpecifyKind(bankDate, DateTimeKind.Unspecified), TimeSpan.FromHours(8)).LocalDateTime;
        var party = detail.Groups["party"].Value.Trim();
        if (String.IsNullOrWhiteSpace(party) || party.Length > 255) throw new MailParseException("Invalid ZA counterparty.");
        var record = new Record {
            date = date, postingDate = date, updateTime = DateTime.Now,
            Reason = reason, DestAccount = party, DescCurrency = new Currency(amount.v, amount.t),
            Source = $"code={ZAMessageKey(message.MessageId)}; ZA mail"
        };
        record.CopyFrom(amount);
        return record;
    }

    private static Currency ParseZAMoney(Match match)
    {
        var value = Decimal.Parse(match.Groups["amount"].Value, NumberStyles.Number, CultureInfo.InvariantCulture);
        if (value <= 0) throw new MailParseException("ZA transaction amount must be positive.");
        return new Currency(value, match.Groups["currency"].Value);
    }

    private static Match MatchZAMail(string text, string pattern)
    {
        var matches = Regex.Matches(text, pattern, RegexOptions.CultureInvariant);
        if (matches.Count != 1) throw new MailParseException("ZA transaction field missing or ambiguous.");
        return matches[0];
    }

    private static string ZAMessageKey(string? messageId)
    {
        if (String.IsNullOrWhiteSpace(messageId)) throw new MailParseException("ZA transaction mail has no Message-ID.");
        return "ZA-mail-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(messageId)));
    }
}
