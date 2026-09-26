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
    private const string HKBankMoneyPattern = @"(?<currency>HKD|USD|CNY|RMB)\s*(?<amount>(?:\d+|\d{1,3}(?:,\d{3})+)\.\d{2})(?![\d.])";

    private Task FetchHKBankMessages(string bank, string sender, StatementImportProvider provider,
        Func<string, bool> isTransaction, Func<MimeMessage, Record> parse)
    {
        var accounts = database.GetAccountsByNamePrefix(bank + "_");
        if (accounts.Count != 1)
            throw new InvalidOperationException($"{bank} requires exactly one configured account.");
        var account = accounts[0];
        if (account.isCredit || !account.relativeBalance)
            throw new InvalidOperationException($"{bank} mail requires a non-credit relative-balance account.");
        return FetchHKBankMessages(bank, sender, provider, isTransaction, message => (parse(message), account));
    }

    private async Task FetchHKBankMessages(string bank, string sender, StatementImportProvider provider,
        Func<string, bool> isTransaction, Func<MimeMessage, (Record Record, Account Account)> parse)
    {
        var imports = database.GetStatementImports(provider);
        var checkpoint = imports.Where(item => item.statementKey == "")
            .Select(item => (DateTime?)item.time).SingleOrDefault()
            ?? throw new InvalidOperationException($"Missing {bank} mail checkpoint.");
        var prefix = $"{bank}-mail-";
        var since = imports.Where(item => item.statementKey.StartsWith(prefix, StringComparison.Ordinal))
            .Select(item => item.time).Append(checkpoint).Max();
        await RunWithMailSessionScope(async () =>
        {
            var messages = await SearchMessagesFromMailbox(CreateYahooMailbox() with { Proxy = null }, $"{bank} transactions",
                SearchQuery.FromContains(sender).And(SearchQuery.SentSince(since.ToUniversalTime().AddHours(-14).Date)),
                summary => SummaryIsFrom(summary, sender) && GetSummaryDateTime(summary) >= since
                    && isTransaction(summary.Envelope?.Subject ?? "")
                    && (String.IsNullOrWhiteSpace(summary.Envelope?.MessageId)
                        || !database.IsStatementKeyImported(provider, HKBankMessageKey(bank, summary.Envelope.MessageId))),
                message => IsFrom(message, sender), GetMailDateTime).ConfigureAwait(false);
            foreach (var message in messages)
            {
                var key = HKBankMessageKey(bank, message.MessageId);
                if (database.IsStatementKeyImported(provider, key)) continue;
                var (record, account) = parse(message);
                if (account.isCredit || !account.relativeBalance)
                    throw new InvalidOperationException($"{bank} mail requires a non-credit relative-balance account.");
                record.Account = account;
                record._account_Id = account.Id;
                if (record.Reason is "转入" or "转出")
                    DatabaseUtil.ApplyTransferCounterparty(record,
                        database.ResolveTransferCounterparty(null, [record.DestAccount], record.DestAccount));
                record.DescCurrency = new Currency(record.v, record.t);
                database.SaveStatementRecordsOnce(provider, GetMailDateTime(message), [record], statementKey: key);
            }
        }).ConfigureAwait(false);
    }

    private static string HKBankMessageKey(string bank, string? messageId)
    {
        if (String.IsNullOrWhiteSpace(messageId)) throw new MailParseException($"{bank} mail has no Message-ID.");
        return $"{bank}-mail-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(messageId)));
    }

    private static string HKBankMailText(MimeMessage message)
    {
        var document = new HtmlDocument();
        document.LoadHtml(message.HtmlBody ?? "");
        foreach (var node in document.DocumentNode.SelectNodes("//style|//script")?.ToList() ?? []) node.Remove();
        return NormalizeMailText(WebUtility.HtmlDecode(message.TextBody ?? document.DocumentNode.InnerText));
    }

    private static Match MatchHKBankMail(string bank, string text, string pattern)
    {
        var matches = Regex.Matches(text, pattern, RegexOptions.CultureInvariant);
        if (matches.Count != 1) throw new MailParseException($"{bank} transaction field missing or ambiguous.");
        return matches[0];
    }

    private static Currency HKBankAmount(Match match)
    {
        var amount = Decimal.Parse(match.Groups["amount"].Value, NumberStyles.Number, CultureInfo.InvariantCulture);
        if (amount <= 0) throw new MailParseException("Bank transaction amount must be positive.");
        return new Currency(amount, match.Groups["currency"].Value);
    }

    private static DateTime HKBankLocalDate(string value, string format) => new DateTimeOffset(
        DateTime.SpecifyKind(DateTime.ParseExact(value, format, CultureInfo.InvariantCulture), DateTimeKind.Unspecified),
        TimeSpan.FromHours(8)).LocalDateTime;

    private static Record HKBankRecord(string bank, MimeMessage message, Currency amount, DateTime date,
        string reason, string party = "")
    {
        var record = new Record { date = date, postingDate = date, updateTime = DateTime.Now,
            Reason = reason, DestAccount = party.Trim(), Source = $"code={HKBankMessageKey(bank, message.MessageId)}; {bank} mail" };
        record.CopyFrom(amount);
        return record;
    }
}
