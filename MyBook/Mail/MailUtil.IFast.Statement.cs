using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace MyBook;

partial class MailUtil
{
    private static readonly CurrencyType[] IFastCurrencies =
        [CurrencyType.GBP, CurrencyType.USD, CurrencyType.EUR, CurrencyType.HKD, CurrencyType.SGD, CurrencyType.RMB];
    private static readonly TimeZoneInfo IFastTimeZone = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
    private const string IFastInitializationPrefix = "IFast-initial-range:";
    private sealed record IFastSummary(decimal Opening, decimal Debit, decimal Credit, decimal Closing);
    private sealed record IFastStatement(DateTime Month, List<Record> Records, Dictionary<CurrencyType, IFastSummary> Summaries);
    private sealed record IFastPdfWord(string Text, double Left, double Right);
    private sealed record IFastPdfLine(string Text, List<IFastPdfWord> Words);

    private static DateTime IFastLocalTime(DateTime bankTime) => TimeZoneInfo.ConvertTime(
        DateTime.SpecifyKind(bankTime, DateTimeKind.Unspecified), IFastTimeZone, TimeZoneInfo.Local);

    private static DateTime IFastBankDate(Record record)
    {
        var explicitDate = Regex.Match(record.Source, @"(?:^|;)\s*bankDate=(\d{4}-\d{2}-\d{2})(?:;|$)");
        if (explicitDate.Success)
            return DateTime.ParseExact(explicitDate.Groups[1].Value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
        // Receipt notices contain a bank date, not a timestamp.
        if (record.Source.StartsWith("code=IFast-receipt-", StringComparison.Ordinal))
            return record.date.Date;
        return TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(record.date, DateTimeKind.Local), IFastTimeZone).Date;
    }

    private List<Record> GetIFastAccountRecords() => IFastCurrencies
        .SelectMany(currency => database.GetAccountRecords(database.GetAccountByName(IFastAccountName(currency)))).ToList();

    private List<IFastStatement> ReadIFastStatements()
    {
        var files = EnumerateInitialReportSearchRoots()
            .Select(root => Path.Combine(root, InitialReportDirectoryName))
            .Where(Directory.Exists)
            .SelectMany(directory => Directory.EnumerateFiles(directory, "Monthly Statement*.pdf"))
            .Distinct(StringComparer.OrdinalIgnoreCase);
        var statements = files.Select(ParseIFastStatement).OrderBy(statement => statement.Month).ToList();
        if (statements.GroupBy(statement => statement.Month).Any(group => group.Count() != 1))
            throw new MailParseException("Duplicate IFast monthly statement files.");
        return statements;
    }

    private (DateTime Start, DateTime End) GetIFastInitializationRange()
    {
        var key = database.GetStatementImports(IFastProvider)
            .SingleOrDefault(item => item.statementKey.StartsWith(IFastInitializationPrefix, StringComparison.Ordinal))?.statementKey
            ?? throw new InvalidOperationException("IFast requires initial monthly statements in initialReports.");
        var range = key[IFastInitializationPrefix.Length..].Split(':');
        return (DateTime.ParseExact(range[0], "yyyy-MM", CultureInfo.InvariantCulture),
            DateTime.ParseExact(range[1], "yyyy-MM", CultureInfo.InvariantCulture));
    }

    private void ImportIFastInitialStatements()
    {
        var statements = ReadIFastStatements();
        if (!database.GetStatementImports(IFastProvider).Any(item => item.statementKey.StartsWith(IFastInitializationPrefix, StringComparison.Ordinal)))
        {
            if (statements.Count == 0)
                throw new InvalidOperationException("IFast requires initial monthly statements in initialReports.");
            for (var index = 1; index < statements.Count; index++)
                if (statements[index].Month != statements[index - 1].Month.AddMonths(1))
                    throw new MailParseException("IFast initialization statements must be consecutive.");
            if (statements[0].Summaries.Values.Any(summary => summary.Opening != 0))
                throw new MailParseException("IFast initialization requires statements starting with zero balances.");
            var end = statements[^1].Month.AddMonths(1);
            database.MarkStatementProcessedOnce(IFastProvider, IFastLocalTime(end).AddTicks(-1),
                $"{IFastInitializationPrefix}{statements[0].Month:yyyy-MM}:{end:yyyy-MM}");
        }
        var range = GetIFastInitializationRange();
        for (var month = range.Start; month < range.End; month = month.AddMonths(1))
        {
            var key = $"IFast-statement-{month:yyyy-MM}";
            var statement = statements.SingleOrDefault(item => item.Month == month);
            if (statement is null)
            {
                if (!database.IsStatementKeyImported(IFastProvider, key))
                    throw new MailParseException($"Missing IFast initialization statement: {month:yyyy-MM}.");
                continue;
            }
            var actual = GetIFastAccountRecords();
            if (actual.Any(record => IFastBankDate(record) < range.Start))
                throw new MailParseException("IFast records precede the zero-balance initialization period.");
            var missing = statement.Records.Where(record => FindIFastMatch(record, actual, false) is null).ToList();
            if (database.IsStatementKeyImported(IFastProvider, key) && missing.Count != 0)
                throw new MailParseException($"IFast initialized statement records changed: {month:yyyy-MM}.");
            ValidateIFastStatement(statement, actual.Concat(missing).ToList());
            database.SaveStatementRecordsOnce(IFastProvider, IFastLocalTime(month.AddMonths(1)).AddTicks(-1), missing,
                statementKey: key, afterSaveInTransaction: _ => ValidateIFastStatement(statement, GetIFastAccountRecords()));
        }
    }

    private void ValidateIFastStatements()
    {
        foreach (var statement in ReadIFastStatements())
            ImportIFastStatementInterest(statement);
    }

    private void ImportIFastStatementInterest(IFastStatement statement)
    {
        var key = $"IFast-statement-{statement.Month:yyyy-MM}";
        var actual = GetIFastAccountRecords();
        var missing = statement.Records.Where(record => record.Reason == "利息"
            && FindIFastMatch(record, actual, false) is null).ToList();
        if (database.IsStatementKeyImported(IFastProvider, key) && missing.Count != 0)
            throw new MailParseException($"IFast validated interest records changed: {statement.Month:yyyy-MM}.");
        if (missing.Count != 0 && database.IsStatementKeyImported(IFastProvider, $"IFast-interest-{statement.Month.AddMonths(-1):yyyy-MM}"))
            throw new MailParseException($"IFast calculated interest differs from statement: {statement.Month:yyyy-MM}.");
        // An incorrect existing interest remains an extra record and fails validation before writing.
        ValidateIFastStatement(statement, actual.Concat(missing).ToList());
        database.SaveStatementRecordsOnce(IFastProvider, IFastLocalTime(statement.Month.AddMonths(1)).AddTicks(-1), missing,
            statementKey: key, afterSaveInTransaction: _ => ValidateIFastStatement(statement, GetIFastAccountRecords()));
    }

    private static Record? FindIFastMatch(Record expected, List<Record> actual, bool required)
    {
        var code = Regex.Match(expected.Source, @"^code=([^;]+)").Groups[1].Value;
        var matches = actual.Where(record => record._account_Id == expected._account_Id
            && record.t == expected.t && record.v == expected.v && IFastBankDate(record) == IFastBankDate(expected)
            && (code.Length > 0 && record.Source.StartsWith($"code={code};", StringComparison.Ordinal)
                || expected.Reason == "利息" && record.Reason == "利息"
                || expected.Reason == "消费" && record.Reason == "消费"
                    && (expected.DestAccount.Contains(record.DestAccount, StringComparison.Ordinal)
                        || record.DestAccount.Contains(expected.DestAccount, StringComparison.Ordinal))
                    && !String.IsNullOrWhiteSpace(expected.DestAccount)
                    && !String.IsNullOrWhiteSpace(record.DestAccount))).ToList();
        if (matches.Count > 1 || required && matches.Count != 1)
            throw new MailParseException($"IFast record mismatch: {IFastBankDate(expected):yyyy-MM-dd}, {expected.t}, {expected.v}, candidates={matches.Count}.");
        return matches.SingleOrDefault();
    }

    private static void ValidateIFastStatement(IFastStatement statement, List<Record> all)
    {
        var end = statement.Month.AddMonths(1);
        var unmatched = all.Where(record => IFastBankDate(record) >= statement.Month && IFastBankDate(record) < end).ToList();
        foreach (var expected in statement.Records)
            unmatched.Remove(FindIFastMatch(expected, unmatched, true)!);
        if (unmatched.Count != 0)
            throw new MailParseException($"IFast statement {statement.Month:yyyy-MM}: {unmatched.Count} extra database records.");
        foreach (var currency in IFastCurrencies)
        {
            var summary = statement.Summaries.GetValueOrDefault(currency) ?? new IFastSummary(0, 0, 0, 0);
            var records = all.Where(record => record.t == currency).ToList();
            var opening = records.Where(record => IFastBankDate(record) < statement.Month).Sum(record => record.v);
            var closing = records.Where(record => IFastBankDate(record) < end).Sum(record => record.v);
            if (opening != summary.Opening || closing != summary.Closing)
                throw new MailParseException($"IFast {statement.Month:yyyy-MM} {currency} balance mismatch: opening={opening}/{summary.Opening}, closing={closing}/{summary.Closing}.");
        }
    }

    private IFastStatement ParseIFastStatement(string path)
    {
        using var document = PdfDocument.Open(path);
        var pages = document.GetPages().Select(page => page.GetWords()
            .GroupBy(word => Math.Round(word.BoundingBox.Bottom / 2) * 2)
            .OrderByDescending(group => group.Key)
            .Select(group => group.OrderBy(word => word.BoundingBox.Left)
                .Select(word => new IFastPdfWord(word.Text, word.BoundingBox.Left, word.BoundingBox.Right)).ToList())
            .Select(words => new IFastPdfLine(String.Join(" ", words.Select(word => word.Text)), words)).ToList()).ToList();
        var title = MatchIFast(String.Join(" ", pages[0].Select(line => line.Text)), @"(?<month>[A-Za-z]{3} \d{4}) Monthly Statement");
        if (!pages[0].Any(line => line.Text.Contains("IFAST GLOBAL BANK LIMITED", StringComparison.OrdinalIgnoreCase)))
            throw new MailParseException("Not an IFast statement.");
        var month = DateTime.ParseExact(title.Groups["month"].Value, "MMM yyyy", CultureInfo.InvariantCulture);
        var summaries = new Dictionary<CurrencyType, IFastSummary>();
        foreach (var line in pages[0])
        {
            var match = Regex.Match(line.Text, @"^(GBP|USD|EUR|HKD|SGD|CNY)\s");
            if (!match.Success) continue;
            var amounts = IFastPdfAmounts(line);
            if (amounts.Count != 4)
                throw new MailParseException("Invalid IFast account summary.");
            var values = amounts.Select(word => ParseIFastPdfAmount(word.Text)).ToList();
            if (values[0] - values[1] + values[2] != values[3])
                throw new MailParseException("IFast summary balance equation failed.");
            summaries.Add(new Currency(0, match.Groups[1].Value).t, new(values[0], values[1], values[2], values[3]));
        }
        if (summaries.Count == 0) throw new MailParseException("Missing IFast currency summaries.");
        var records = new List<Record>();
        var openings = new HashSet<CurrencyType>();
        var closings = new HashSet<CurrencyType>();
        var totals = new HashSet<CurrencyType>();
        CurrencyType? current = null;
        decimal running = 0;
        double debitRight = 0, creditRight = 0;
        Record? pending = null;
        var description = "";
        void FinishRecord()
        {
            if (pending is null) return;
            CompleteIFastPdfRecord(pending, description);
            records.Add(pending);
            pending = null;
            description = "";
        }
        foreach (var page in pages.Skip(1))
        {
            var active = false;
            foreach (var line in page)
            {
                var heading = Regex.Match(line.Text, @"Multi-Currency Current Account (?<currency>GBP|USD|EUR|HKD|SGD|CNY)\(");
                if (heading.Success)
                {
                    FinishRecord();
                    current = new Currency(0, heading.Groups["currency"].Value).t;
                    active = true;
                    continue;
                }
                if (!active || current is null) continue;
                if (line.Text.StartsWith("©") || line.Text.StartsWith("Anything wrong?"))
                { FinishRecord(); active = false; continue; }
                if (line.Text.StartsWith("Transaction Date Description"))
                {
                    debitRight = line.Words.Single(word => word.Text == "Debit").Right;
                    creditRight = line.Words.Single(word => word.Text == "Credit").Right;
                    continue;
                }
                var amounts = IFastPdfAmounts(line);
                if (line.Text.StartsWith("Opening Balance"))
                {
                    if (amounts.Count != 1 || !openings.Add(current.Value)) throw new MailParseException("Invalid IFast opening balance.");
                    running = ParseIFastPdfAmount(amounts[0].Text);
                    if (running != summaries[current.Value].Opening) throw new MailParseException("IFast opening balance differs from summary.");
                    continue;
                }
                if (line.Text.StartsWith("Balance Carried Forward"))
                {
                    FinishRecord();
                    if (amounts.Count != 1 || ParseIFastPdfAmount(amounts[0].Text) != running
                        || running != summaries[current.Value].Closing || !closings.Add(current.Value))
                        throw new MailParseException("IFast closing balance mismatch.");
                    continue;
                }
                if (line.Text.StartsWith("Total Movements"))
                {
                    var items = records.Where(record => record.t == current.Value).ToList();
                    if (amounts.Count != 2 || !totals.Add(current.Value)
                        || ParseIFastPdfAmount(amounts[0].Text) != -items.Where(record => record.v < 0).Sum(record => record.v)
                        || ParseIFastPdfAmount(amounts[1].Text) != items.Where(record => record.v > 0).Sum(record => record.v)
                        || ParseIFastPdfAmount(amounts[0].Text) != summaries[current.Value].Debit
                        || ParseIFastPdfAmount(amounts[1].Text) != summaries[current.Value].Credit)
                        throw new MailParseException("IFast debit/credit totals mismatch.");
                    active = false;
                    continue;
                }
                var date = Regex.Match(line.Text, @"^(\d{2}/\d{2}/\d{4})\s");
                if (date.Success)
                {
                    FinishRecord();
                    var bankDate = DateTime.ParseExact(date.Groups[1].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    if (bankDate < month || bankDate >= month.AddMonths(1) || amounts.Count != 2 || creditRight == 0)
                        throw new MailParseException("Invalid IFast transaction row.");
                    var movement = amounts[0];
                    if (movement.Right > creditRight + 2) throw new MailParseException("Missing IFast debit/credit amount.");
                    var value = ParseIFastPdfAmount(movement.Text) * (movement.Right <= debitRight + 2 ? -1 : 1);
                    running += value;
                    if (running != ParseIFastPdfAmount(amounts[1].Text)) throw new MailParseException("IFast transaction running balance mismatch.");
                    pending = BuildIFastRecord(new Currency(value, current.Value), IFastLocalTime(bankDate), "", "", "");
                    pending.Source = $"IFast statement; bankDate={bankDate:yyyy-MM-dd}";
                    description = String.Join(" ", line.Words.Where(word => word.Left > 100 && word.Right < debitRight - 20).Select(word => word.Text));
                }
                else if (pending is not null)
                    description += " " + String.Join(" ", line.Words.Where(word => word.Left > 100 && word.Right < debitRight - 20).Select(word => word.Text));
            }
            FinishRecord();
        }
        if (!openings.SetEquals(summaries.Keys) || !closings.SetEquals(summaries.Keys) || !totals.SetEquals(summaries.Keys))
            throw new MailParseException("IFast statement currency sections incomplete.");
        return new IFastStatement(month, records, summaries);
    }

    private static List<IFastPdfWord> IFastPdfAmounts(IFastPdfLine line) => line.Words
        .Where(word => Regex.IsMatch(word.Text, @"^(?:\d+|\d{1,3}(?:,\d{3})+)\.\d{2}$")).ToList();
    private static decimal ParseIFastPdfAmount(string text) => Decimal.Parse(text, NumberStyles.Number, CultureInfo.InvariantCulture);

    private static void CompleteIFastPdfRecord(Record record, string description)
    {
        var reference = Regex.Match(description, @"\b\d{16}\b").Value;
        var interest = Regex.Match(description, @"^Credit interest for ([A-Za-z]{3} \d{4})\b");
        string code;
        if (interest.Success && record.v > 0)
        {
            var period = DateTime.ParseExact(interest.Groups[1].Value, "MMM yyyy", CultureInfo.InvariantCulture);
            if (period.AddMonths(1) != IFastBankDate(record)) throw new MailParseException("Unexpected IFast interest posting date.");
            record.Reason = "利息";
            code = $"IFast-interest-{period:yyyy-MM}-{record.t}";
        }
        else if (description.StartsWith("Currency Conversion ") && reference.Length > 0)
        {
            record.Reason = "换汇";
            record.isInternal = true;
            code = $"BALANCE-IFAST-{reference}";
        }
        else if (description.StartsWith("Inbound domestic payment ") && record.v > 0 && reference.Length > 0)
        {
            record.Reason = "转入";
            code = $"IFast-receipt-{reference}";
        }
        else if (record.v < 0 && Regex.IsMatch(description, @"QR|Scan.*Pay", RegexOptions.IgnoreCase))
        {
            record.Reason = "消费";
            code = "";
        }
        else throw new MailParseException("Unsupported IFast statement transaction description.");
        record.DestAccount = description.Length > 200 ? description[..200] : description;
        record.Source = $"code={code}; {record.Source}";
    }
}
