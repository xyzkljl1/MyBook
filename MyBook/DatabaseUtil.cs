using Microsoft.Extensions.Configuration;
using SqlSugar;
using System.Reflection;

namespace MyBook
{
    class DatabaseUtil
    {
        private const string DefaultConnectionString = "server=localhost;port=3306;database=mybook;uid=root;pwd=;charset=utf8mb4;";
        private readonly SqlSugarClient db;
        private static readonly Type[] SchemaTypes = [typeof(Account), typeof(AccountBalance), typeof(Record), typeof(Holding), typeof(Finance), typeof(StatementImport)];
        private static readonly ForeignKeyDefinition[] ForeignKeys =
        [
            new("fk_Accounts_primaryAccount", "Accounts", "_primaryAccount_Id", "Accounts", "Id"),
            new("fk_AccountBalances_account", "AccountBalances", "_account_Id", "Accounts", "Id"),
            new("fk_Holdings_account", "Holdings", "_account_Id", "Accounts", "Id"),
            new("fk_Records_account", "Records", "_account_Id", "Accounts", "Id"),
            new("fk_Records_statementImport", "Records", "_statementImport_Id", "StatementImports", "Id")
        ];

        public DatabaseUtil(IConfigurationRoot config)
        {
            var connectionString = config["database_connection"]
                ?? config.GetConnectionString("Default")
                ?? DefaultConnectionString;

            db = new SqlSugarClient(new ConnectionConfig
            {
                ConnectionString = connectionString,
                DbType = DbType.MySql,
                InitKeyType = InitKeyType.Attribute,
                IsAutoCloseConnection = true
            });
            ValidateSchema();
            ValidateForeignKeys();
            ValidateAccountPrimaryRelations();
        }

        public bool IsStatementImported(StatementImportProvider provider, DateTime time)
        {
            return IsStatementImported(provider, time, "");
        }

        public bool IsStatementImported(StatementImportProvider provider, DateTime time, string statementKey)
        {
            var importTime = NormalizeStatementImportTime(time);
            return db.Queryable<StatementImport>()
                .Any(it => it.provider == provider && it.time == importTime && it.statementKey == statementKey);
        }

        public bool IsStatementKeyImported(StatementImportProvider provider, string statementKey)
        {
            return db.Queryable<StatementImport>()
                .Any(it => it.provider == provider && it.statementKey == statementKey);
        }

        public DateTime? GetLatestStatementImportTime(StatementImportProvider provider)
        {
            var latestImport = db.Queryable<StatementImport>()
                .Where(it => it.provider == provider)
                .OrderByDescending(it => it.time)
                .First();
            return latestImport?.time;
        }

        public string? GetLatestStatementImportKey(StatementImportProvider provider)
        {
            var latestImport = db.Queryable<StatementImport>()
                .Where(it => it.provider == provider && it.statementKey != "")
                .OrderByDescending(it => it.statementKey)
                .First();
            return latestImport?.statementKey;
        }

        public bool SaveStatementRecordsOnce(
            StatementImportProvider provider,
            DateTime time,
            IEnumerable<Record> records,
            IEnumerable<AccountBalance>? accountBalances = null,
            string statementKey = "",
            IEnumerable<AccountBalance>? beginningAccountBalances = null)
        {
            var recordList = records.ToList();
            var accountBalanceList = accountBalances?.ToList() ?? [];
            var beginningAccountBalanceList = beginningAccountBalances?.ToList() ?? [];
            db.Ado.BeginTran();
            try
            {
                if (IsStatementImported(provider, time, statementKey))
                {
                    db.Ado.CommitTran();
                    return false;
                }

                ValidateBeginningAccountBalances(
                    provider,
                    beginningAccountBalanceList,
                    ShouldValidateBeginningAccountBalances(provider));
                ValidateRecordBalanceChanges(provider, recordList, beginningAccountBalanceList, accountBalanceList);
                var statementImportId = InsertStatementImport(provider, time, statementKey);
                SaveAccountBalancesCore(accountBalanceList);
                SaveRecordsCore(recordList, statementImportId);
                db.Ado.CommitTran();
                return true;
            }
            catch (Exception e)
            {
                db.Ado.RollbackTran();
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time, statementKey))
                    return false;
                throw;
            }
        }

        public List<bool> SaveStatementRecordsAndHoldingsOnce(IEnumerable<StatementRecordHoldingImport> imports)
        {
            var importList = imports.ToList();
            var saved = new List<bool>();
            var shouldValidateBeginningBalances = importList
                .Select(import => import.Provider)
                .Distinct()
                .ToDictionary(provider => provider, ShouldValidateBeginningAccountBalances);
            db.Ado.BeginTran();
            try
            {
                foreach (var import in importList)
                {
                    if (IsStatementImported(import.Provider, import.Time, import.StatementKey))
                    {
                        saved.Add(false);
                        continue;
                    }

                    ValidateBeginningAccountBalances(
                        import.Provider,
                        import.BeginningAccountBalances,
                        shouldValidateBeginningBalances[import.Provider]);
                    ValidateRecordBalanceChanges(import.Provider, import.Records, import.BeginningAccountBalances, import.AccountBalances);
                    var statementImportId = InsertStatementImport(import.Provider, import.Time, import.StatementKey);
                    SaveAccountBalancesCore(import.AccountBalances);
                    SaveAccountHoldingsCore(import.HoldingAccount, import.Holdings);
                    SaveRecordsCore(import.Records, statementImportId);
                    saved.Add(true);
                }

                db.Ado.CommitTran();
                return saved;
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
        }

        public bool SaveStatementRecordsWithoutBalanceOnce(
            StatementImportProvider provider,
            DateTime time,
            string statementKey,
            IEnumerable<Record> records)
        {
            var recordList = records.ToList();
            db.Ado.BeginTran();
            try
            {
                if (IsStatementImported(provider, time, statementKey))
                {
                    db.Ado.CommitTran();
                    return false;
                }

                var statementImportId = InsertStatementImport(provider, time, statementKey);
                SaveRecordsCore(recordList, statementImportId);
                db.Ado.CommitTran();
                return true;
            }
            catch (Exception e)
            {
                db.Ado.RollbackTran();
                if (IsDuplicateKeyException(e) && IsStatementImported(provider, time, statementKey))
                    return false;
                throw;
            }
        }

        public Account GetAccountByTypeAndId(string? accountType, string id)
        {
            return GetAccountByName(BuildAccountName(accountType, id));
        }

        public Account GetAccountByName(string accountName)
        {
            var account = FindAccountByName(accountName);
            if (account is null)
                throw new InvalidOperationException($"Account not found: {accountName}");

            return account;
        }

        public List<Account> GetAllAccounts()
        {
            return db.Queryable<Account>().ToList();
        }

        private Account? FindAccountByName(string accountName)
        {
            return db.Queryable<Account>()
                .Where(it => it.name == accountName)
                .First();
        }

        public static string BuildAccountName(string? accountType, string id)
        {
            var parts = new[] { accountType, id }
                .Where(part => !String.IsNullOrWhiteSpace(part))
                .Select(part => part!.Trim());
            return String.Join("_", parts);
        }

        private Account GetExistingAccountByName(Account account)
        {
            var existing = GetAccountByName(account.name);
            account.Id = existing.Id;
            return existing;
        }

        public Account GetPostingAccount(Account account)
        {
            var current = GetExistingAccountByName(account);
            if (current._primaryAccount_Id is null)
                return current;

            if (current._primaryAccount_Id.Value == current.Id)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name} points to itself");

            var primary = db.Queryable<Account>()
                .Where(it => it.Id == current._primaryAccount_Id.Value)
                .First();
            if (primary is null)
                throw new InvalidOperationException($"Invalid account primary relation: {current.name} points to missing account {current._primaryAccount_Id.Value}");
            if (primary._primaryAccount_Id is not null)
                throw new InvalidOperationException($"Invalid account primary relation: primary account {primary.name} is also a supplementary account");

            return primary;
        }

        private int InsertStatementImport(StatementImportProvider provider, DateTime time, string statementKey)
        {
            return db.Insertable(new StatementImport
            {
                provider = provider,
                time = NormalizeStatementImportTime(time),
                statementKey = statementKey
            }).ExecuteReturnIdentity();
        }

        private void SaveRecordsCore(List<Record> recordList, int statementImportId)
        {
            if (statementImportId <= 0)
                throw new InvalidOperationException("Record statement import is required.");

            foreach (var record in recordList)
            {
                if (record.Account is null)
                    throw new InvalidOperationException("Record account is required.");

                var account = GetPostingAccount(record.Account);
                record.Account = account;
                record._account_Id = account.Id;
                record._statementImport_Id = statementImportId;
            }

            if (recordList.Count > 0)
                db.Insertable(recordList).ExecuteCommand();
        }

        private bool ShouldValidateBeginningAccountBalances(StatementImportProvider provider)
        {
            return db.Queryable<StatementImport>()
                .Where(it => it.provider == provider)
                .Count() > 1;
        }

        private void ValidateBeginningAccountBalances(
            StatementImportProvider provider,
            List<AccountBalance> beginningAccountBalances,
            bool shouldValidate)
        {
            if (!shouldValidate || beginningAccountBalances.Count == 0)
                return;

            foreach (var beginningAccountBalance in beginningAccountBalances)
            {
                if (beginningAccountBalance.Account is null)
                    throw new InvalidOperationException("Beginning account balance account is required.");

                var account = GetPostingAccount(beginningAccountBalance.Account);
                var existing = db.Queryable<AccountBalance>()
                    .Where(it => it._account_Id == account.Id && it.t == beginningAccountBalance.t)
                    .First();
                if (existing is null)
                {
                    throw new InvalidOperationException(
                        $"Missing current account balance for {provider}: {account.name}/{beginningAccountBalance.t}");
                }

                if (existing.v != beginningAccountBalance.v)
                {
                    throw new InvalidOperationException(
                        $"Beginning account balance mismatch for {provider}: {account.name}/{beginningAccountBalance.t}, current={existing.v}, beginning={beginningAccountBalance.v}");
                }
            }
        }

        private void ValidateRecordBalanceChanges(
            StatementImportProvider provider,
            List<Record> records,
            List<AccountBalance> beginningAccountBalances,
            List<AccountBalance> endingAccountBalances)
        {
            if (beginningAccountBalances.Count == 0 || endingAccountBalances.Count == 0)
                throw new InvalidOperationException($"Record balance validation requires beginning and ending account balances for {provider}.");

            var beginningBalances = BuildAccountBalanceMap(beginningAccountBalances, "Beginning account balance");
            var endingBalances = BuildAccountBalanceMap(endingAccountBalances, "Ending account balance");
            var recordChanges = SumRecordAmountsByAccountAndCurrency(records);
            var keys = beginningBalances.Keys
                .Union(endingBalances.Keys)
                .Union(recordChanges.Keys)
                .ToList();

            foreach (var key in keys)
            {
                var beginning = beginningBalances.TryGetValue(key, out var beginningValue) ? beginningValue : 0;
                var ending = endingBalances.TryGetValue(key, out var endingValue) ? endingValue : 0;
                var change = recordChanges.TryGetValue(key, out var changeValue) ? changeValue : 0;
                var expectedEnding = beginning + change;
                if (expectedEnding != ending)
                {
                    throw new InvalidOperationException(
                        $"Record balance mismatch for {provider}: accountId={key.AccountId}, currency={key.Currency}, beginning={beginning}, records={change}, ending={ending}");
                }
            }
        }

        public Dictionary<(int AccountId, CurrencyType Currency), decimal> SumRecordAmountsByAccountAndCurrency(IEnumerable<Record> records)
        {
            var sums = new Dictionary<(int AccountId, CurrencyType Currency), decimal>();
            foreach (var record in records)
            {
                if (record.Account is null)
                    throw new InvalidOperationException("Record account is required.");

                var account = GetPostingAccount(record.Account);
                var key = (account.Id, record.t);
                sums[key] = sums.TryGetValue(key, out var current)
                    ? current + record.v
                    : record.v;
            }

            return sums.ToDictionary(item => item.Key, item => item.Value);
        }

        private Dictionary<(int AccountId, CurrencyType Currency), decimal> BuildAccountBalanceMap(
            List<AccountBalance> accountBalances,
            string context)
        {
            var balances = new Dictionary<(int AccountId, CurrencyType Currency), decimal>();
            foreach (var accountBalance in accountBalances)
            {
                if (accountBalance.Account is null)
                    throw new InvalidOperationException($"{context} account is required.");

                var account = GetPostingAccount(accountBalance.Account);
                balances[(account.Id, accountBalance.t)] = accountBalance.v;
            }

            return balances;
        }

        private void SaveAccountBalancesCore(List<AccountBalance> accountBalances)
        {
            foreach (var accountBalance in accountBalances)
            {
                if (accountBalance.Account is null)
                    throw new InvalidOperationException("Account balance account is required.");

                var account = GetPostingAccount(accountBalance.Account);
                accountBalance.Account = account;
                accountBalance._account_Id = account.Id;

                var existing = db.Queryable<AccountBalance>()
                    .Where(it => it._account_Id == account.Id && it.t == accountBalance.t)
                    .First();
                if (existing is null)
                {
                    db.Insertable(accountBalance).ExecuteCommand();
                    continue;
                }

                existing.v = accountBalance.v;
                db.Updateable(existing).ExecuteCommand();
            }
        }

        public void SaveAccountHoldings(Account account, IEnumerable<Holding> holdings)
        {
            var holdingList = holdings.ToList();
            db.Ado.BeginTran();
            try
            {
                SaveAccountHoldingsCore(account, holdingList);
                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }
        }

        public List<Holding> GetHoldings()
        {
            return db.Queryable<Holding>()
                .Includes(it => it.Account)
                .ToList();
        }

        public List<Finance> GetFinances()
        {
            return db.Queryable<Finance>().ToList();
        }

        public DashboardData GetDashboardData(DateTime today)
        {
            var currentMonth = new DateTime(today.Year, today.Month, 1);
            var firstMonth = currentMonth.AddMonths(-11);
            var nextMonth = currentMonth.AddMonths(1);
            var lastMonthStart = currentMonth.AddMonths(-1);
            var lastMonthEnd = currentMonth.AddDays(-1);
            var reasonMonths = Enumerable.Range(0, 12)
                .Select(firstMonth.AddMonths)
                .ToList();
            var records = db.Queryable<Record>()
                .Where(record => !record.isInternal && record.date >= firstMonth && record.date < nextMonth)
                .ToList();
            var balances = db.Queryable<AccountBalance>()
                .Where(balance => balance.v != 0)
                .ToList();
            var holdings = db.Queryable<Holding>().ToList();
            var investmentAccountIds = holdings
                .Select(holding => holding._account_Id)
                .Distinct()
                .ToHashSet();
            var investmentRecords = db.Queryable<Record>()
                .Where(record => !record.isInternal && record.v > 0 && record.date < today.Date.AddDays(1))
                .ToList()
                .Where(record => investmentAccountIds.Contains(record._account_Id))
                .ToList();
            var accounts = db.Queryable<Account>().ToList().ToDictionary(account => account.Id, account => account.name);
            var holdingNames = BuildHoldingNames(holdings, accounts);
            var exchangeRates = GetCurrencyToRmbRates();
            var usedCurrencies = balances
                .Select(balance => balance.t)
                .Union(records.Select(record => record.t))
                .Union(investmentRecords.Select(record => record.t))
                .Distinct()
                .ToList();

            return new DashboardData
            {
                CurrencySummaries = BuildCurrencySummaries(balances),
                MonthlyFlowSeries = BuildMonthlyFlowSeries(records, firstMonth),
                RmbMonthlyFlowSeries = BuildRmbMonthlyFlowSeries(records, firstMonth, exchangeRates),
                RmbReasonFlowSeriesByMonth = reasonMonths
                    .Select(month => BuildRmbReasonFlowSeries(records, month, month.AddMonths(1), exchangeRates))
                    .ToList(),
                DefaultReasonMonthIndex = Math.Max(0, reasonMonths.FindIndex(month => month == lastMonthStart)),
                InvestmentByReason = BuildInvestmentStatistics(
                    investmentRecords,
                    today,
                    exchangeRates,
                    record => String.IsNullOrWhiteSpace(record.Reason) ? "未分类" : record.Reason),
                InvestmentByHolding = BuildInvestmentStatistics(
                    investmentRecords,
                    today,
                    exchangeRates,
                    record => BuildHoldingInvestmentKey(record, holdingNames)),
                TotalAssetsRmb = Currency.RoundMoney(balances
                    .Select(balance => TryConvertToRmb(balance.v, balance.t, exchangeRates))
                    .Where(value => value.HasValue)
                    .Sum(value => value!.Value)),
                MissingExchangeRateCurrencies = usedCurrencies
                    .Where(currency => currency != CurrencyType.RMB && !exchangeRates.ContainsKey(currency))
                    .OrderBy(currency => currency)
                    .ToList(),
                LastMonthStart = lastMonthStart,
                LastMonthEnd = lastMonthEnd
            };
        }

        private Dictionary<CurrencyType, decimal> GetCurrencyToRmbRates()
        {
            var rates = new Dictionary<CurrencyType, decimal>
            {
                [CurrencyType.RMB] = 1
            };
            var finances = db.Queryable<Finance>()
                .Where(finance => finance.holdingType == HoldingType.Cash && finance._currentPrice_t == CurrencyType.RMB)
                .ToList();
            foreach (var finance in finances)
            {
                if (finance._currentPrice_v <= 0 || !Enum.TryParse<CurrencyType>(finance.code, out var currency))
                    continue;

                rates[currency] = finance._currentPrice_v;
            }

            return rates;
        }

        private static List<CurrencyBalanceSummary> BuildCurrencySummaries(List<AccountBalance> balances)
        {
            return balances
                .GroupBy(balance => balance.t)
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    var assets = Currency.RoundMoney(group.Where(balance => balance.v > 0).Sum(balance => balance.v));
                    var liabilities = Currency.RoundMoney(group.Where(balance => balance.v < 0).Sum(balance => balance.v));
                    return new CurrencyBalanceSummary
                    {
                        Currency = group.Key,
                        Assets = assets,
                        Liabilities = liabilities,
                        Net = Currency.RoundMoney(assets + liabilities),
                        AccountCount = group.Select(balance => balance._account_Id).Distinct().Count()
                    };
                })
                .ToList();
        }

        private static List<MonthlyFlowSeries> BuildMonthlyFlowSeries(List<Record> records, DateTime firstMonth)
        {
            var months = Enumerable.Range(0, 12)
                .Select(firstMonth.AddMonths)
                .ToList();
            return records
                .Select(record => record.t)
                .Distinct()
                .OrderBy(currency => currency)
                .Select(currency =>
                {
                    var points = months
                        .Select(month =>
                        {
                            var monthRecords = records
                                .Where(record => record.t == currency
                                    && record.date >= month
                                    && record.date < month.AddMonths(1))
                                .ToList();
                            return new MonthlyFlowPoint
                            {
                                Month = month,
                                MonthLabel = month.ToString("MM月"),
                                Income = Currency.RoundMoney(monthRecords.Where(record => record.v > 0).Sum(record => record.v)),
                                Expense = Currency.RoundMoney(-monthRecords.Where(record => record.v < 0).Sum(record => record.v)),
                                NetChange = Currency.RoundMoney(monthRecords.Sum(record => record.v)),
                                IncomeSegments = BuildSingleCurrencySegments(currency, monthRecords.Where(record => record.v > 0).Sum(record => record.v)),
                                ExpenseSegments = BuildSingleCurrencySegments(currency, -monthRecords.Where(record => record.v < 0).Sum(record => record.v))
                            };
                        })
                        .ToList();
                    return new MonthlyFlowSeries
                    {
                        DisplayName = currency.ToString(),
                        Currency = currency,
                        Points = points,
                        TotalIncome = Currency.RoundMoney(points.Sum(point => point.Income)),
                        TotalExpense = Currency.RoundMoney(points.Sum(point => point.Expense)),
                        NetChange = Currency.RoundMoney(points.Sum(point => point.NetChange))
                    };
                })
                .ToList();
        }

        private static MonthlyFlowSeries BuildRmbMonthlyFlowSeries(
            List<Record> records,
            DateTime firstMonth,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var months = Enumerable.Range(0, 12)
                .Select(firstMonth.AddMonths)
                .ToList();
            var points = months
                .Select(month =>
                {
                    var convertedRecords = records
                        .Where(record => record.date >= month && record.date < month.AddMonths(1))
                        .Select(record => TryConvertToRmb(record.v, record.t, exchangeRates))
                        .Where(value => value.HasValue)
                        .Select(value => value!.Value)
                        .ToList();
                    return new MonthlyFlowPoint
                    {
                        Month = month,
                        MonthLabel = month.ToString("MM月"),
                        Income = Currency.RoundMoney(convertedRecords.Where(value => value > 0).Sum()),
                        Expense = Currency.RoundMoney(-convertedRecords.Where(value => value < 0).Sum()),
                        NetChange = Currency.RoundMoney(convertedRecords.Sum()),
                        IncomeSegments = BuildRmbMonthlySegments(
                            records.Where(record => record.date >= month && record.date < month.AddMonths(1) && record.v > 0),
                            exchangeRates),
                        ExpenseSegments = BuildRmbMonthlySegments(
                            records.Where(record => record.date >= month && record.date < month.AddMonths(1) && record.v < 0),
                            exchangeRates,
                            invertSign: true)
                    };
                })
                .ToList();
            return new MonthlyFlowSeries
            {
                DisplayName = "",
                Currency = CurrencyType.RMB,
                Points = points,
                TotalIncome = Currency.RoundMoney(points.Sum(point => point.Income)),
                TotalExpense = Currency.RoundMoney(points.Sum(point => point.Expense)),
                NetChange = Currency.RoundMoney(points.Sum(point => point.NetChange))
            };
        }

        private static ReasonFlowSeries BuildRmbReasonFlowSeries(
            List<Record> records,
            DateTime start,
            DateTime end,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var lastMonthRecords = records
                .Where(record => record.date >= start && record.date < end)
                .ToList();
            var items = lastMonthRecords
                .GroupBy(record => String.IsNullOrWhiteSpace(record.Reason) ? "未分类" : record.Reason)
                .SelectMany(group =>
                {
                    var result = new List<ReasonFlowItem>();
                    var expenses = BuildReasonFlowItem(group.Key, false, group.Where(record => record.v < 0), exchangeRates);
                    var incomes = BuildReasonFlowItem(group.Key, true, group.Where(record => record.v > 0), exchangeRates);
                    if (expenses is not null)
                        result.Add(expenses);
                    if (incomes is not null)
                        result.Add(incomes);
                    return result;
                })
                .Where(item => item.Total != 0)
                .OrderBy(item => item.IsIncome)
                .ThenByDescending(item => item.Total)
                .ThenBy(item => item.Reason)
                .ToList();
            return new ReasonFlowSeries
            {
                DisplayName = "",
                Currency = CurrencyType.RMB,
                Month = start,
                MonthLabel = start.ToString("yyyy年MM月"),
                Items = items,
                TotalIncome = Currency.RoundMoney(items.Where(item => item.IsIncome).Sum(item => item.Total)),
                TotalExpense = Currency.RoundMoney(items.Where(item => !item.IsIncome).Sum(item => item.Total))
            };
        }

        private static List<MonthlyFlowSegment> BuildSingleCurrencySegments(CurrencyType currency, decimal value)
        {
            var rounded = Currency.RoundMoney(value);
            return rounded == 0
                ? []
                : [new MonthlyFlowSegment { Currency = currency, Value = rounded, Label = currency.ToString() }];
        }

        private static List<MonthlyFlowSegment> BuildRmbMonthlySegments(
            IEnumerable<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates,
            bool invertSign = false)
        {
            return records
                .GroupBy(record => record.t)
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    var original = group.Sum(record => record.v);
                    if (invertSign)
                        original = -original;
                    var converted = TryConvertToRmb(original, group.Key, exchangeRates) ?? 0;
                    return new MonthlyFlowSegment
                    {
                        Currency = group.Key,
                        Value = converted,
                        Label = group.Key.ToString()
                    };
                })
                .Where(segment => segment.Value != 0)
                .ToList();
        }

        private static ReasonFlowItem? BuildReasonFlowItem(
            string reason,
            bool isIncome,
            IEnumerable<Record> records,
            Dictionary<CurrencyType, decimal> exchangeRates)
        {
            var groups = records
                .GroupBy(record => record.t)
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    var original = group.Sum(record => record.v);
                    if (!isIncome)
                        original = -original;
                    var converted = TryConvertToRmb(original, group.Key, exchangeRates);
                    return new
                    {
                        Currency = group.Key,
                        Original = Currency.RoundMoney(original),
                        Converted = converted
                    };
                })
                .Where(item => item.Converted.HasValue && item.Converted.Value != 0)
                .ToList();
            if (groups.Count == 0)
                return null;

            var total = Currency.RoundMoney(groups.Sum(item => item.Converted!.Value));
            var detailSum = Currency.RoundMoney(groups.Sum(item => item.Converted!.Value));
            if (total != detailSum)
                throw new InvalidOperationException($"Reason currency conversion mismatch: {reason}/{(isIncome ? "income" : "expense")}");

            return new ReasonFlowItem
            {
                Reason = reason,
                IsIncome = isIncome,
                Total = total,
                CurrencyDetails = BuildCurrencyDetailText(groups.Select(item => (item.Currency, item.Original, item.Converted!.Value)))
            };
        }

        private static decimal? TryConvertToRmb(decimal value, CurrencyType currency, Dictionary<CurrencyType, decimal> exchangeRates)
        {
            return exchangeRates.TryGetValue(currency, out var rate)
                ? Currency.RoundMoney(value * rate)
                : null;
        }

        private static string BuildCurrencyDetailText(IEnumerable<(CurrencyType Currency, decimal Original, decimal Converted)> details)
        {
            var currencies = details
                .Where(item => item.Currency != CurrencyType.RMB)
                .Select(item => $"{item.Currency} ¥{item.Converted:N0}")
                .ToList();
            return currencies.Count == 0 ? "" : String.Join("、", currencies);
        }

        private static Dictionary<(int AccountId, string Code), string> BuildHoldingNames(List<Holding> holdings, Dictionary<int, string> accounts)
        {
            return holdings
                .GroupBy(holding => (holding._account_Id, holding.code))
                .ToDictionary(
                    group => group.Key,
                    group =>
                    {
                        var holding = group.First();
                        var display = String.IsNullOrWhiteSpace(holding.displayText)
                            ? holding.code
                            : holding.displayText;
                        if (String.IsNullOrWhiteSpace(display))
                            display = holding.holdingType.ToString();
                        return accounts.TryGetValue(holding._account_Id, out var accountName)
                            ? $"{display} / {accountName}"
                            : display;
                    });
        }

        private static string BuildHoldingInvestmentKey(Record record, Dictionary<(int AccountId, string Code), string> holdingNames)
        {
            if (String.IsNullOrWhiteSpace(record.DestAccount))
                return "未关联持仓";

            return holdingNames.TryGetValue((record._account_Id, record.DestAccount), out var display)
                ? display
                : record.DestAccount;
        }

        private static InvestmentStatistics BuildInvestmentStatistics(
            List<Record> records,
            DateTime today,
            Dictionary<CurrencyType, decimal> exchangeRates,
            Func<Record, string> keySelector)
        {
            var end = today.Date.AddDays(1);
            return new InvestmentStatistics
            {
                Periods =
                [
                    BuildInvestmentStatisticsPeriod("过去一个月", records, today.Date.AddMonths(-1), end, exchangeRates, keySelector),
                    BuildInvestmentStatisticsPeriod("年初至今", records, new DateTime(today.Year, 1, 1), end, exchangeRates, keySelector),
                    BuildInvestmentStatisticsPeriod("开始至今", records, DateTime.MinValue, end, exchangeRates, keySelector)
                ]
            };
        }

        private static InvestmentStatisticsPeriod BuildInvestmentStatisticsPeriod(
            string title,
            List<Record> records,
            DateTime start,
            DateTime end,
            Dictionary<CurrencyType, decimal> exchangeRates,
            Func<Record, string> keySelector)
        {
            var items = records
                .Where(record => record.date >= start && record.date < end && record.v > 0)
                .GroupBy(keySelector)
                .Select(group => new InvestmentStatisticsItem
                {
                    Name = group.Key,
                    Total = Currency.RoundMoney(group
                        .Select(record => TryConvertToRmb(record.v, record.t, exchangeRates))
                        .Where(value => value.HasValue)
                        .Sum(value => value!.Value))
                })
                .Where(item => item.Total > 0)
                .OrderByDescending(item => item.Total)
                .ThenBy(item => item.Name)
                .ToList();
            return new InvestmentStatisticsPeriod
            {
                Title = title,
                Items = items
            };
        }

        public DatabaseCleanupResult CleanVolatileData()
        {
            var beforeCounts = ReadCleanupCounts();
            db.Ado.BeginTran();
            try
            {
                db.Deleteable<Record>().ExecuteCommand();
                db.Deleteable<Holding>().ExecuteCommand();
                db.Deleteable<Finance>().ExecuteCommand();
                db.Deleteable<AccountBalance>().ExecuteCommand();
                db.Ado.ExecuteCommand("""
                    delete statementImport
                    from `StatementImports` statementImport
                    left join (
                        select `provider`, min(`time`) as `time`
                        from `StatementImports`
                        where `provider` in ('IBKRReportMail', 'ICBCBillMail')
                        group by `provider`
                    ) fixedImport
                        on statementImport.`provider` = fixedImport.`provider`
                        and statementImport.`time` = fixedImport.`time`
                    where fixedImport.`provider` is null
                    """);
                db.Ado.CommitTran();
            }
            catch
            {
                db.Ado.RollbackTran();
                throw;
            }

            return new DatabaseCleanupResult(
                beforeCounts,
                ReadCleanupCounts(),
                db.Queryable<StatementImport>()
                    .OrderBy(it => it.provider)
                    .OrderBy(it => it.time)
                    .ToList());
        }

        private Dictionary<string, int> ReadCleanupCounts()
        {
            return new Dictionary<string, int>
            {
                ["Accounts"] = db.Queryable<Account>().Count(),
                ["AccountBalances"] = db.Queryable<AccountBalance>().Count(),
                ["StatementImports"] = db.Queryable<StatementImport>().Count(),
                ["Records"] = db.Queryable<Record>().Count(),
                ["Holdings"] = db.Queryable<Holding>().Count(),
                ["Finance"] = db.Queryable<Finance>().Count()
            };
        }

        public void SaveFinance(Finance finance)
        {
            if (finance.currentPriceTime <= 0)
                finance.currentPriceTime = DateTimeOffset.Now.ToUnixTimeSeconds();

            var existing = db.Queryable<Finance>()
                .Where(it => it.code == finance.code && it.holdingType == finance.holdingType)
                .First();
            if (existing is null)
            {
                finance.Id = db.Insertable(finance).ExecuteReturnIdentity();
                return;
            }

            existing._currentPrice_v = finance._currentPrice_v;
            existing._currentPrice_t = finance._currentPrice_t;
            existing.currentPriceTime = finance.currentPriceTime;
            db.Updateable(existing).ExecuteCommand();
        }

        public static DateTime NormalizeStatementImportTime(DateTime time)
        {
            return time.Date;
        }

        private static bool IsDuplicateKeyException(Exception exception)
        {
            for (var current = exception; current is not null; current = current.InnerException)
            {
                if (current.Message.Contains("Duplicate entry", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private void ValidateSchema()
        {
            foreach (var type in SchemaTypes)
            {
                var tableName = GetTableName(type);
                if (!db.DbMaintenance.IsAnyTable(tableName, false))
                    throw new InvalidOperationException($"Database schema mismatch: missing table {tableName}");

                var dbColumns = db.DbMaintenance.GetColumnInfosByTableName(tableName, false)
                    .Select(column => column.DbColumnName)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                var codeColumns = GetColumnNames(type);

                var missingColumns = codeColumns.Except(dbColumns, StringComparer.OrdinalIgnoreCase).ToList();
                var extraColumns = dbColumns.Except(codeColumns, StringComparer.OrdinalIgnoreCase).ToList();
                if (missingColumns.Count > 0 || extraColumns.Count > 0)
                {
                    var missingText = missingColumns.Count > 0 ? String.Join(",", missingColumns) : "none";
                    var extraText = extraColumns.Count > 0 ? String.Join(",", extraColumns) : "none";
                    throw new InvalidOperationException(
                        $"Database schema mismatch: table {tableName}, missing columns [{missingText}], extra columns [{extraText}]");
                }
            }
        }

        private void ValidateForeignKeys()
        {
            foreach (var foreignKey in ForeignKeys)
            {
                var exists = db.Ado.GetInt("""
                    select count(*)
                    from information_schema.key_column_usage
                    where table_schema = database()
                        and constraint_name = @constraintName
                        and table_name = @tableName
                        and column_name = @columnName
                        and referenced_table_name = @referencedTableName
                        and referenced_column_name = @referencedColumnName
                    """,
                    new SugarParameter("@constraintName", foreignKey.ConstraintName),
                    new SugarParameter("@tableName", foreignKey.TableName),
                    new SugarParameter("@columnName", foreignKey.ColumnName),
                    new SugarParameter("@referencedTableName", foreignKey.ReferencedTableName),
                    new SugarParameter("@referencedColumnName", foreignKey.ReferencedColumnName)) > 0;
                if (!exists)
                {
                    throw new InvalidOperationException(
                        $"Database schema mismatch: missing foreign key {foreignKey.ConstraintName} on {foreignKey.TableName}.{foreignKey.ColumnName}");
                }
            }
        }

        private void ValidateAccountPrimaryRelations()
        {
            var accounts = db.Queryable<Account>().ToList();
            var accountsById = accounts.ToDictionary(account => account.Id);
            foreach (var account in accounts.Where(account => account._primaryAccount_Id.HasValue))
            {
                var primaryAccountId = account._primaryAccount_Id!.Value;
                if (primaryAccountId == account.Id)
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name} points to itself");

                if (!accountsById.TryGetValue(primaryAccountId, out var primaryAccount))
                    throw new InvalidOperationException($"Invalid account primary relation: {account.name} points to missing account {primaryAccountId}");

                if (primaryAccount._primaryAccount_Id.HasValue)
                    throw new InvalidOperationException($"Invalid account primary relation: primary account {primaryAccount.name} is also a supplementary account");
            }
        }

        private static string GetTableName(Type type)
        {
            if (type == typeof(Account))
                return "Accounts";
            if (type == typeof(AccountBalance))
                return "AccountBalances";
            if (type == typeof(Record))
                return "Records";
            if (type == typeof(Holding))
                return "Holdings";
            if (type == typeof(Finance))
                return "Finance";
            if (type == typeof(StatementImport))
                return "StatementImports";

            return type.Name;
        }

        private static HashSet<string> GetColumnNames(Type type)
        {
            return type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.GetIndexParameters().Length == 0)
                .Where(property => property.GetCustomAttribute<SugarColumn>()?.IsIgnore != true)
                .Where(property => !HasNavigateAttribute(property))
                .Select(property =>
                {
                    var column = property.GetCustomAttribute<SugarColumn>();
                    return String.IsNullOrWhiteSpace(column?.ColumnName) ? property.Name : column.ColumnName;
                })
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private static bool HasNavigateAttribute(PropertyInfo property)
        {
            return property.GetCustomAttributes()
                .Any(attribute => attribute.GetType().Name is "Navigate" or "NavigateAttribute");
        }

        private static string GetHoldingKey(Holding holding)
        {
            return $"{holding.code}\t{holding.holdingType}";
        }

        private void SaveAccountHoldingsCore(Account account, List<Holding> holdingList)
        {
            account = GetAccountByName(account.name);

            foreach (var holding in holdingList)
            {
                NormalizeHolding(holding);
                holding.Account = account;
                holding._account_Id = account.Id;
            }

            ValidateAccountHoldingsBalance(account, holdingList);

            var existingHoldings = db.Queryable<Holding>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            var currentKeys = holdingList.Select(GetHoldingKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var deletedIds = existingHoldings
                .Where(holding => !currentKeys.Contains(GetHoldingKey(holding)))
                .Select(holding => holding.Id)
                .ToList();
            if (deletedIds.Count > 0)
                db.Deleteable<Holding>().In(deletedIds).ExecuteCommand();

            foreach (var holding in holdingList)
            {
                var existing = existingHoldings.FirstOrDefault(it =>
                    it.code == holding.code && it.holdingType == holding.holdingType);
                if (existing is null)
                {
                    db.Insertable(holding).ExecuteCommand();
                    continue;
                }

                existing.quantity = holding.quantity;
                existing.desc = holding.desc;
                existing.displayText = holding.displayText;
                existing._currentPrice_v = holding._currentPrice_v;
                existing._currentPrice_t = holding._currentPrice_t;
                existing._account_Id = account.Id;
                db.Updateable(existing).ExecuteCommand();
            }
        }

        private static void NormalizeHolding(Holding holding)
        {
            if (Holding.IsSingleValueAsset(holding.holdingType))
                holding.quantity = 1;
        }

        private void ValidateAccountHoldingsBalance(Account account, List<Holding> holdings)
        {
            var accountBalances = db.Queryable<AccountBalance>()
                .Where(it => it._account_Id == account.Id)
                .ToList();
            if (accountBalances.Count == 0)
                throw new InvalidOperationException($"Missing account balance for holdings: {account.name}");

            var balanceSums = accountBalances
                .GroupBy(balance => balance.t)
                .ToDictionary(group => group.Key, group => group.Sum(balance => balance.v));
            var holdingSums = holdings
                .GroupBy(holding => holding.currentPrice.t)
                .ToDictionary(group => group.Key, group => group.Sum(holding => holding.totalPrice.v));
            var currencies = balanceSums.Keys
                .Union(holdingSums.Keys)
                .ToList();

            foreach (var currency in currencies)
            {
                var accountBalance = balanceSums.TryGetValue(currency, out var balance) ? balance : 0;
                var holdingTotal = holdingSums.TryGetValue(currency, out var total) ? total : 0;
                if (accountBalance != holdingTotal)
                {
                    throw new InvalidOperationException(
                        $"Holding balance mismatch for {account.name}/{currency}: accountBalance={accountBalance}, holdings={holdingTotal}");
                }
            }
        }

        private sealed record ForeignKeyDefinition(
            string ConstraintName,
            string TableName,
            string ColumnName,
            string ReferencedTableName,
            string ReferencedColumnName);
    }

    class DatabaseCleanupResult
    {
        public DatabaseCleanupResult(
            Dictionary<string, int> beforeCounts,
            Dictionary<string, int> afterCounts,
            List<StatementImport> fixedStatementImports)
        {
            BeforeCounts = beforeCounts;
            AfterCounts = afterCounts;
            FixedStatementImports = fixedStatementImports;
        }

        public Dictionary<string, int> BeforeCounts { get; }
        public Dictionary<string, int> AfterCounts { get; }
        public List<StatementImport> FixedStatementImports { get; }
    }

    class StatementRecordHoldingImport
    {
        public StatementRecordHoldingImport(
            StatementImportProvider provider,
            DateTime time,
            string statementKey,
            Account holdingAccount,
            List<Record> records,
            List<Holding> holdings,
            List<AccountBalance> accountBalances,
            List<AccountBalance> beginningAccountBalances)
        {
            Provider = provider;
            Time = time;
            StatementKey = statementKey;
            HoldingAccount = holdingAccount;
            Records = records;
            Holdings = holdings;
            AccountBalances = accountBalances;
            BeginningAccountBalances = beginningAccountBalances;
        }

        public StatementImportProvider Provider { get; }
        public DateTime Time { get; }
        public string StatementKey { get; }
        public Account HoldingAccount { get; }
        public List<Record> Records { get; }
        public List<Holding> Holdings { get; }
        public List<AccountBalance> AccountBalances { get; }
        public List<AccountBalance> BeginningAccountBalances { get; }
    }
}
