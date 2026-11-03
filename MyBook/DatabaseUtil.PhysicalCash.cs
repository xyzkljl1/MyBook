namespace MyBook;

partial class DatabaseUtil
{
    internal static bool IsPhysicalCashAccount(Account account) => GetAccountType(account.name) == "CASH";

    internal Account GetPhysicalCashAccount()
    {
        var accounts = GetAllAccounts().Where(IsPhysicalCashAccount).ToList();
        if (accounts.Count != 1 || !accounts[0].relativeBalance || accounts[0].isCredit || accounts[0]._primaryAccount_Id is not null)
            throw new InvalidOperationException("Exactly one relative-balance physical cash account is required.");
        return accounts[0];
    }

    internal void PrepareFirstTradeAtmWithdrawal(Record record)
    {
        var rates = GetExchangeLossRates(GetRateHistory(RateSource.GoogleFinance));
        var rate = GetExchangeLossRate(rates, record.t, record.date.Date, fromRmb: false)
            ?? throw new InvalidOperationException($"ATM withdrawal has no usable Google quote on {record.date:yyyy-MM-dd}.");
        var amount = InferRmbAtmCashAmount(-record.v, rate);
        AttachPhysicalCashTransfer(record, GetPhysicalCashAccount(), new Currency(amount, CurrencyType.RMB));
        record.reasonDetail = record.MatchedRecord!.reasonDetail = FormattableString.Invariant(
            $"ATM cash inferred from Google {record.t}/RMB benchmark {rate}, requested date {record.date:yyyy-MM-dd}, tolerance 1%, RMB 100 increments.");
    }

    internal static decimal InferRmbAtmCashAmount(decimal principal, decimal rateToRmb)
    {
        if (principal <= 0 || rateToRmb <= 0)
            throw new InvalidOperationException("ATM principal and benchmark rate must be positive.");
        // User-approved rule: these ATM withdrawals receive RMB in multiples of 100.
        // Accept only one candidate within 1% of the benchmark, never round to the nearest hundred.
        var minimum = Decimal.Ceiling(principal * rateToRmb * 0.99m / 100m) * 100m;
        var maximum = Decimal.Floor(principal * rateToRmb * 1.01m / 100m) * 100m;
        if (minimum <= 0 || minimum != maximum)
            throw new InvalidOperationException("ATM withdrawal requires exactly one RMB 100 multiple within 1% of the benchmark.");
        return minimum;
    }

    // Modules call this only for identified cash withdrawals/deposits, excluding their separate fees.
    internal static void AttachPhysicalCashTransfer(Record principal, Account cashAccount, Currency cashAmount)
    {
        if (!IsPhysicalCashAccount(cashAccount) || !cashAccount.relativeBalance || principal.Account is null
            || IsPhysicalCashAccount(principal.Account) || !HasOppositeSigns(principal.v, cashAmount.v))
            throw new InvalidOperationException("Physical cash transfer requires opposite bank and cash principal amounts.");
        principal.isInternal = true;
        principal.DestAccount = cashAccount.name;
        principal.Reason = principal.t == cashAmount.t ? "转账" : "换汇";
        var counterpart = new Record
        {
            Account = cashAccount, v = cashAmount.v, t = cashAmount.t,
            date = principal.date, postingDate = principal.postingDate, updateTime = DateTime.Now,
            Reason = principal.Reason, isInternal = true, DestAccount = principal.Account.name,
            Source = ("PhysicalCash|" + principal.Source)[..Math.Min(1024, 13 + principal.Source.Length)]
        };
        if (principal.t != cashAmount.t)
        {
            var debit = principal.v < 0 ? principal : counterpart;
            var credit = principal.v > 0 ? principal : counterpart;
            debit.DescCurrency = new Currency(-credit.v, credit.t);
        }
        principal.MatchedRecord = counterpart;
    }
}
