using Newtonsoft.Json.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace MyBook;

internal sealed partial class WiseUtil
{
    internal static ReceiptData ParseReceipt(byte[] bytes, string transferId)
    {
        if (bytes.Length < 5 || System.Text.Encoding.ASCII.GetString(bytes, 0, 5) != "%PDF-")
            throw Error("receipt: expected PDF");
        using var document = PdfDocument.Open(bytes);
        if (document.NumberOfPages is < 1 or > 10) throw Error("receipt: unexpected page count");
        // Preserve the PDF's content order: sorting by Y merges the account and address columns.
        var text = String.Join("\n", document.GetPages().Select(page => ContentOrderTextExtractor.GetText(page)));
        if (!Regex.IsMatch(text, @"(?<!\d)" + Regex.Escape(transferId) + @"(?!\d)"))
            throw Error("receipt: transfer identity missing");
        return new(200, text, Convert.ToHexString(SHA256.HashData(bytes)));
    }

    private AccountMatch FindCounterparty(EventData data)
    {
        var match = ResolveCounterparty(data,
            value => database.FindAccountByInternalCardNo(value),
            texts => database.FindAccountByInternalCardNoText(null, "Wise API counterparty", false, texts),
            account => database.GetPostingAccount(account));
        if (Text(data.Activity, "type") is not ("TRANSFER" or "BALANCE_DEPOSIT")) return match;
        // Historical beneficiary names are authoritative; do not scan references or intermediary banks.
        var names = new List<string> { Plain(Text(data.Activity, "title")) };
        if (ParseAmount(Text(data.Activity, "primaryAmount")).Sign != "+" && data.Receipt?.Text is string receipt)
        {
            var lines = receipt.Split('\n').Select(line => line.Trim()).ToArray();
            var start = System.Array.FindIndex(lines, line => line == "Sent to");
            var end = start < 0 ? -1 : System.Array.FindIndex(lines, start + 1, line => line == "Account details");
            if (start >= 0 && end > start) names.Add(String.Join(" ", lines[(start + 1)..end]));
        }
        var exact = match.AccountName is null ? null : database.GetAccountByName(match.AccountName);
        var account = database.FindTransferAccountByInstitution(exact, names.ToArray());
        return account is null ? match : match with { Status = "Matched", AccountName = database.GetPostingAccount(account).name };
    }

    internal static AccountMatch ResolveCounterparty(EventData data, Func<string, Account?> exact,
        Func<string[], Account?> textMatch, Func<Account, Account> postingAccount)
    {
        var type = Text(data.Activity, "type");
        if (type == "INTERBALANCE") return new("ConversionPendingFees", [], [], null);
        if (type is not ("TRANSFER" or "BALANCE_DEPOSIT" or "DIRECT_DEBIT_TRANSACTION"))
            return new("NotApplicable", [], [], null);
        var references = new[] { Optional(data.Transfer?["details"], "reference"), Optional(data.Activity, "title"), Optional(data.Activity, "description") }
            .Where(s => !String.IsNullOrWhiteSpace(s)).Select(s => Plain(s!)).Distinct(StringComparer.Ordinal).ToArray();
        var identifiers = new List<string>();
        var currentOnly = false;
        var incoming = ParseAmount(Text(data.Activity, "primaryAmount")).Sign == "+";
        if (!incoming && data.Recipient is not null)
        {
            // A recipient belonging to Wise's balance is the receiving Wise side, not an external payer.
            if (Text(data.Recipient, "type") != "balance")
            {
                var current = new[] { "accountNumber", "iban", "IBAN" }
                    .Select(field => Optional(data.Recipient["details"], field)).Where(s => s is not null).Cast<string>()
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                var historical = data.Receipt?.Text is string receipt ? ReceiptAccountIdentifiers(receipt) : [];
                var mt103 = Optional(data.Payout, "mt103");
                var beneficiary = mt103 is null ? "" : Regex.Match(mt103, @"(?:^|\n):59(?:A|F)?:/?([^\r\n]+)").Groups[1].Value.Trim();
                if (beneficiary.Length > 0)
                {
                    if (historical.Count > 0 && !historical.Any(value => SameAccountIdentifier(value, beneficiary)))
                        throw Error("counterparty: receipt and MT103 account conflict");
                    historical.Add(beneficiary);
                }
                if (historical.Count > 0 && current.Count > 0 && current.Any(value => !historical.Any(old => SameAccountIdentifier(value, old))))
                    throw Error("counterparty: current recipient differs from historical receipt");
                identifiers.AddRange(historical.Count > 0 ? historical : current);
                currentOnly = historical.Count == 0 && current.Count > 0;
            }
        }
        var ids = identifiers.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var candidates = new List<Account>();
        if (!currentOnly)
            foreach (var id in ids)
                if (exact(id) is Account account) candidates.Add(postingAccount(account));
        foreach (var reference in references)
            foreach (var (value, allowSuffix) in ReferenceAccountIdentifiers(reference))
            {
                var referenced = exact(value) ?? (allowSuffix ? textMatch([value]) : null);
                if (referenced is not null) candidates.Add(postingAccount(referenced));
            }
        var matches = candidates.DistinctBy(a => a.Id).ToList();
        if (matches.Count > 1) throw Error("counterparty: conflicting account matches");
        return new(matches.Count == 1 ? "Matched" : currentOnly ? "CurrentRecipientOnly" : "Unmatched",
            ids, references, matches.SingleOrDefault()?.name);
    }

    internal static IEnumerable<(string Value, bool AllowSuffix)> ReferenceAccountIdentifiers(string text)
    {
        // Bare numeric fields are allowed by policy, but numbers embedded in arbitrary prose are not.
        if (Regex.IsMatch(text.Trim(), @"^\d{4,34}$"))
        {
            yield return (text.Trim(), true);
            yield break;
        }
        const string pattern = @"(?:(?<tail>尾号|末四位|末4位|\bending(?:\s+in)?\b|\blast\s+4(?:\s+digits)?\b)"
            + @"|(?:账号|帐号|卡号|\b(?:account|acct|card|iban)\b)(?:\s+(?:number|no\.?))?)"
            + @"\s*[:：#]?\s*(?<mask>[*xX]{2,}\s*)?(?<id>[A-Z]{0,4}\d[A-Z0-9-]{3,33})(?![A-Z0-9-])";
        foreach (Match match in Regex.Matches(text, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            var value = match.Groups["id"].Value;
            var suffix = match.Groups["tail"].Success || match.Groups["mask"].Success;
            if (suffix && !Regex.IsMatch(value, @"^\d{4,34}$")) continue;
            yield return (value, suffix);
        }
    }

    internal static List<string> ReceiptAccountIdentifiers(string text)
    {
        var lines = text.Split('\n').Select(s => s.Trim()).ToArray();
        var start = System.Array.FindIndex(lines, line => line == "Sent to");
        var end = start < 0 ? -1 : System.Array.FindIndex(lines, start + 1, line => line == "Paid out from");
        if (start < 0 || end < 0) throw Error("receipt: recipient section missing");
        var heading = System.Array.FindIndex(lines, start + 1, end - start - 1, line => line == "Account details");
        if (heading < 0) throw Error("receipt: recipient account details missing");
        var result = new List<string>();
        for (var i = heading + 1; i < end && lines[i] != "Address"; i++)
        {
            var value = Regex.Replace(lines[i], @"\s", "");
            if (Regex.IsMatch(value, @"^(?:\+?\d{7,34}|[A-Z]{2}\d{2}[A-Z0-9]{10,30}|[^\s@]+@[^\s@]+\.[^\s@]+)$", RegexOptions.IgnoreCase))
                result.Add(value);
        }
        if (result.Count == 0) throw Error("receipt: no recognized recipient account identifier");
        return result;
    }

    private static bool SameAccountIdentifier(string left, string right)
    {
        string Normalize(string value) => Regex.Replace(value, @"\s|-", "").ToUpperInvariant();
        left = Normalize(left); right = Normalize(right);
        if (left.Contains('@') || right.Contains('@')) return left == right;
        return left == right || Regex.IsMatch(left, @"^[A-Z]{2}\d{2}") && right.Length >= 7 && left.EndsWith(right, StringComparison.Ordinal)
            || Regex.IsMatch(right, @"^[A-Z]{2}\d{2}") && left.Length >= 7 && right.EndsWith(left, StringComparison.Ordinal);
    }

    internal static void ApplyCounterparty(List<Record> records, Account current, AccountMatch match)
    {
        if (match.AccountName is null || match.AccountName == current.name) return;
        foreach (var record in records.Where(r => r.Reason != "手续费"))
        {
            record.DestAccount = match.AccountName;
            // Gross amounts with unknown fees stay visible even when the counterparty is known.
            record.isInternal = record.Reason == "转账";
        }
    }
}
