using Newtonsoft.Json.Linq;

namespace MyBook
{
    // Nexus Mods GraphQL module. DP transaction detail methods are intentionally left as stubs for now.
    partial class GraphQLUtil
    {
        private const StatementImportProvider NexusDpProvider = StatementImportProvider.NexusDpMonthlyReport;
        private const string DefaultNexusAccountName = "NEXUS";
        private const int DefaultNexusMonthlyBackfillMonths = 6;
        private const decimal NexusDpPerUsd = 1000m;

        public async Task<int> FetchNexusAccountId()
        {
            const string query = """
                query NexusPersonalApiKey {
                  personalApiKey {
                    userId
                  }
                }
                """;

            var data = await ExecuteNexusQuery(query);
            return data["personalApiKey"]?["userId"]?.Value<int>()
                ?? throw new InvalidOperationException("Nexus personalApiKey.userId is missing");
        }

        public Task<NexusDpTransactionPage> FetchNexusDpTransactions(
            int? accountId = null,
            int start = 0,
            int perPage = 100,
            string orderDir = "desc",
            string orderColumn = "created_at")
        {
            if (start < 0)
                throw new ArgumentOutOfRangeException(nameof(start), "start must be non-negative");
            if (perPage <= 0)
                throw new ArgumentOutOfRangeException(nameof(perPage), "perPage must be positive");
            if (!String.Equals(orderDir, "asc", StringComparison.OrdinalIgnoreCase)
                && !String.Equals(orderDir, "desc", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("orderDir must be asc or desc", nameof(orderDir));
            }

            return Task.FromResult(new NexusDpTransactionPage(accountId ?? 0, 0, 0, []));
        }

        public Task<List<NexusDpTransaction>> FetchAllNexusDpTransactions(
            int? accountId = null,
            int pageSize = 100)
        {
            if (pageSize <= 0)
                throw new ArgumentOutOfRangeException(nameof(pageSize), "pageSize must be positive");

            return Task.FromResult(new List<NexusDpTransaction>());
        }

        public async Task<NexusDpMonthlyReport> FetchNexusDpMonthlyReport(int year, int month, int? accountId = null)
        {
            if (month is < 1 or > 12)
                throw new ArgumentOutOfRangeException(nameof(month), "month must be in 1..12");

            var resolvedAccountId = await ResolveNexusAccountId(accountId);
            const string query = """
                query NexusDpMonthlyReport($accountId: Int!, $year: Int!, $month: Int!) {
                  userMonthlyReport(accountId: $accountId, year: $year, month: $month) {
                    userId
                    reportType
                    entries {
                      year
                      month
                      value
                      modValue
                      modCount
                      reportId
                    }
                  }
                }
                """;

            var data = await ExecuteNexusQuery(query, new
            {
                accountId = resolvedAccountId,
                year,
                month
            });
            var report = data["userMonthlyReport"] as JObject
                ?? throw new InvalidOperationException("Nexus userMonthlyReport response is missing");
            var entries = report["entries"]?
                .Select(token => ParseNexusDpMonthlyReportEntry(token, year, month))
                .ToList() ?? [];
            return new NexusDpMonthlyReport(
                resolvedAccountId,
                report["userId"]?.Value<int>() ?? resolvedAccountId,
                year,
                month,
                report["reportType"]?.ToString() ?? "",
                entries);
        }

        public async Task<NexusDpMonthlySummary> FetchNexusDpMonthlySummary(int? accountId = null)
        {
            var resolvedAccountId = await ResolveNexusAccountId(accountId);
            const string query = """
                query NexusDpMonthlySummary($accountId: Int!) {
                  userMonthlySummary(accountId: $accountId) {
                    userId
                    entries {
                      year
                      month
                      value
                      modValue
                      modCount
                      reportType
                    }
                  }
                }
                """;

            var data = await ExecuteNexusQuery(query, new { accountId = resolvedAccountId });
            var summary = data["userMonthlySummary"] as JObject
                ?? throw new InvalidOperationException("Nexus userMonthlySummary response is missing");
            var entries = summary["entries"]?
                .Select(ParseNexusDpMonthlySummaryEntry)
                .ToList() ?? [];
            return new NexusDpMonthlySummary(
                resolvedAccountId,
                summary["userId"]?.Value<int>() ?? resolvedAccountId,
                entries);
        }

        public async Task FetchNexusDpMonthlyReports()
        {
            var db = database ?? throw new InvalidOperationException("FetchNexusDpMonthlyReports requires a database.");
            var account = db.GetAccountByName(GetNexusAccountName());
            var firstMonth = GetFirstNexusMonthlyReportMonth(db);
            var lastMonth = FirstDayOfMonth(DateTime.Today).AddMonths(-1);
            if (firstMonth > lastMonth)
                return;

            var accountId = await ResolveNexusAccountId(null);
            var savedCount = 0;
            var skippedCount = 0;
            for (var month = firstMonth; month <= lastMonth; month = month.AddMonths(1))
            {
                var statementDate = LastDayOfMonth(month);
                var statementKey = BuildNexusDpStatementKey(account, month);
                if (db.IsStatementImported(NexusDpProvider, statementDate, statementKey))
                {
                    skippedCount++;
                    continue;
                }

                var report = await FetchNexusDpMonthlyReport(month.Year, month.Month, accountId);
                if (SaveNexusDpMonthlyReport(db, account, report))
                    savedCount++;
                else
                    skippedCount++;
            }

            Console.WriteLine($"Fetch Nexus DP monthly reports done: saved={savedCount}, skipped={skippedCount}");
        }

        private bool SaveNexusDpMonthlyReport(DatabaseUtil db, Account account, NexusDpMonthlyReport report)
        {
            var statementMonth = new DateTime(report.Year, report.Month, 1);
            var statementDate = LastDayOfMonth(statementMonth);
            var statementKey = BuildNexusDpStatementKey(account, statementMonth);
            var usdIncome = report.TotalDp / NexusDpPerUsd;
            var beginningBalance = db.GetAccountBalance(account, CurrencyType.USD);
            var endingBalance = new Currency(beginningBalance.v + usdIncome, CurrencyType.USD);
            var records = usdIncome == 0
                ? new List<Record>()
                : [
                    new Record
                    {
                        Account = account,
                        v = usdIncome,
                        t = CurrencyType.USD,
                        date = statementDate,
                        updateTime = DateTime.Now,
                        DestAccount = "Donation Points",
                        Source = $"NexusDpMonthlyReport/{report.Year:D4}-{report.Month:D2}/{report.TotalDp}DP",
                        Reason = "创作收入"
                    }
                ];

            return db.SaveStatementRecordsOnce(
                NexusDpProvider,
                statementDate,
                records,
                [new AccountBalance(account, endingBalance)],
                statementKey,
                [new AccountBalance(account, beginningBalance)]);
        }

        private string GetNexusAccountName()
        {
            return String.IsNullOrWhiteSpace(config["nexus_account_name"])
                ? DefaultNexusAccountName
                : config["nexus_account_name"]!.Trim();
        }

        private DateTime GetFirstNexusMonthlyReportMonth(DatabaseUtil db)
        {
            var latestImport = db.GetLatestStatementImportTime(NexusDpProvider);
            if (latestImport.HasValue)
                return FirstDayOfMonth(latestImport.Value).AddMonths(1);

            var currentMonth = FirstDayOfMonth(DateTime.Today);
            return currentMonth.AddMonths(-DefaultNexusMonthlyBackfillMonths);
        }

        private async Task<int> ResolveNexusAccountId(int? accountId)
        {
            if (accountId.HasValue)
                return accountId.Value;

            if (Int32.TryParse(config["nexus_account_id"], out var configuredAccountId))
                return configuredAccountId;

            return await FetchNexusAccountId();
        }

        private static string BuildNexusDpStatementKey(Account account, DateTime month)
        {
            return $"{account.name}_{month:yyyy-MM}";
        }

        private static DateTime FirstDayOfMonth(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        private static DateTime LastDayOfMonth(DateTime month)
        {
            return new DateTime(month.Year, month.Month, DateTime.DaysInMonth(month.Year, month.Month));
        }

        private static NexusDpMonthlyReportEntry ParseNexusDpMonthlyReportEntry(JToken token, int expectedYear, int expectedMonth)
        {
            var year = token["year"]?.Value<int>()
                ?? throw new InvalidOperationException($"Nexus monthly report year is missing: {token}");
            var month = token["month"]?.Value<int>()
                ?? throw new InvalidOperationException($"Nexus monthly report month is missing: {token}");
            if (year != expectedYear || month != expectedMonth)
                throw new InvalidOperationException($"Nexus monthly report entry has unexpected period: {token}");

            return new NexusDpMonthlyReportEntry(
                year,
                month,
                token["value"]?.Value<int>() ?? 0,
                token["modValue"]?.Value<int>() ?? 0,
                token["modCount"]?.Value<int>() ?? 0,
                token["reportId"]?.Value<int>() ?? 0);
        }

        private static NexusDpMonthlySummaryEntry ParseNexusDpMonthlySummaryEntry(JToken token)
        {
            return new NexusDpMonthlySummaryEntry(
                token["year"]?.Value<int>()
                    ?? throw new InvalidOperationException($"Nexus monthly summary year is missing: {token}"),
                token["month"]?.Value<int>()
                    ?? throw new InvalidOperationException($"Nexus monthly summary month is missing: {token}"),
                token["value"]?.Value<int>() ?? 0,
                token["modValue"]?.Value<int>() ?? 0,
                token["modCount"]?.Value<int>() ?? 0,
                token["reportType"]?.ToString() ?? "");
        }
    }

    public record NexusDpTransactionPage(
        int AccountId,
        int FilteredCount,
        int TotalCount,
        List<NexusDpTransaction> Transactions);

    public record NexusDpTransaction(
        int Id,
        int Amount,
        int SignedAmount,
        NexusDpTransactionDirection Direction,
        DateTimeOffset? CreatedAt,
        string Label,
        string Type,
        string Creditor,
        string Debitor,
        NexusPaymentEntity? CreditorEntity,
        NexusPaymentEntity? DebitorEntity);

    public enum NexusDpTransactionDirection
    {
        Unknown,
        In,
        Out,
    }

    public record NexusPaymentEntity(
        int Id,
        string Label,
        string Type);

    public record NexusDpMonthlyReport(
        int AccountId,
        int UserId,
        int Year,
        int Month,
        string ReportType,
        List<NexusDpMonthlyReportEntry> Entries)
    {
        public int TotalDp => Entries.Sum(entry => entry.Value);
    }

    public record NexusDpMonthlyReportEntry(
        int Year,
        int Month,
        int Value,
        int ModValue,
        int ModCount,
        int ReportId);

    public record NexusDpMonthlySummary(
        int AccountId,
        int UserId,
        List<NexusDpMonthlySummaryEntry> Entries);

    public record NexusDpMonthlySummaryEntry(
        int Year,
        int Month,
        int Value,
        int ModValue,
        int ModCount,
        string ReportType);
}
