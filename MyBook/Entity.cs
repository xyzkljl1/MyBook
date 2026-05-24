using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SqlSugar;
using SqlSugar.DbConvert;

namespace MyBook
{
    public class MailParseException : Exception
    {
        public MailParseException(string? message) : base(message)
        {
        }
    }
    public class CurrencyParseException : Exception
    {
        public CurrencyParseException(string? message) : base(message)
        {
        }
    }
    public enum HoldingType
    {
        // 纳斯达克交易所上市资产。
        NASDAQ,
        // NYSE Arca 交易所上市资产。
        ARCA,
        // 美国国债。
        UST,
        // 上海交易所上市资产。
        SHANGHAI,
        // 国内基金。
        CNFUND,
        // 现金类持仓。
        Cash,
        // 应计、待结算、或还未实际入账但已计入账户净资产的项目。
        Accrued
    };

    // 账户持有的股票、债券、现金、应计项目或其它资产，使用 code + holdingType 区分具体资产。
    // currentPrice 表示账单/报表导入时记录的价格；单值资产的 quantity 固定为 1，currentPrice 表示该项总额。
    [SugarIndex("unique_Holdings_account_code_holding_type", nameof(Holding._account_Id), OrderByType.Asc, nameof(Holding.code), OrderByType.Asc, nameof(Holding.holdingType), OrderByType.Asc, true)]
    [SugarTable("Holdings")]
    public class Holding
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string code { get; set; } = "";

        [SugarColumn(DefaultValue = "NASDAQ", ColumnDataType = MySqlEnumColumnTypes.HoldingType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public HoldingType holdingType { get; set; } = HoldingType.NASDAQ;

        [SugarColumn(DefaultValue = "0")]
        public int quantity
        {
            get => IsSingleValueAsset(holdingType) ? 1 : _quantity;
            set => _quantity = IsSingleValueAsset(holdingType) ? 1 : value;
        }

        [SugarColumn(DefaultValue = "''")]
        public string desc { get; set; } = "";

        // 面向界面的可读名称，仅在同步持仓列表时刷新。
        [SugarColumn(DefaultValue = "''")]
        public string displayText { get; set; } = "";

        [SugarColumn(IsIgnore = true)]
        // 账单/报表导入时记录的价格，金额和币种分别存储。
        public Currency currentPrice
        {
            get { return new Currency(_currentPrice_v, _currentPrice_t); }
            set
            {
                _currentPrice_v = value.v;
                _currentPrice_t = value.t;
            }
        }

        [SugarColumn(IsIgnore = true)]
        public Currency totalPrice
        {
            get { return new Currency(quantity * currentPrice.v, currentPrice.t); }
        }

        public Holding()
        {
        }

        public Holding(string _c, HoldingType _t)
        {
            code = _c;
            holdingType = _t;
        }

        public static bool IsSingleValueAsset(HoldingType holdingType)
        {
            return holdingType is HoldingType.Cash or HoldingType.Accrued;
        }

        // 用于存储
        [SugarColumn(DefaultValue = "0", ColumnDataType = "decimal(24,12)")]
        public decimal _currentPrice_v { get; set; } = 0;

        [SugarColumn(DefaultValue = "RMB", ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType _currentPrice_t { get; set; } = CurrencyType.RMB;

        [SugarColumn(DefaultValue = "0")]
        public int _account_Id { get; set; } = 0;

        private int _quantity = 0;
    }

    // 从互联网获取的最新股票价格或汇率，不关联 Account。
    [SugarIndex("unique_Finance_code_holding_type", nameof(Finance.code), OrderByType.Asc, nameof(Finance.holdingType), OrderByType.Asc, true)]
    [SugarTable("Finance")]
    public class Finance
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string code { get; set; } = "";

        [SugarColumn(DefaultValue = "NASDAQ", ColumnDataType = MySqlEnumColumnTypes.HoldingType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public HoldingType holdingType { get; set; } = HoldingType.NASDAQ;

        [SugarColumn(IsIgnore = true)]
        public Currency currentPrice
        {
            get { return new Currency(_currentPrice_v, _currentPrice_t); }
            set
            {
                _currentPrice_v = value.v;
                _currentPrice_t = value.t;
            }
        }

        [SugarColumn(DefaultValue = "0")]
        public long currentPriceTime { get; set; } = 0;

        public Finance()
        {
        }

        public Finance(string _c, HoldingType _t)
        {
            code = _c;
            holdingType = _t;
        }

        public static Finance FromHolding(Holding holding)
        {
            var code = holding.holdingType == HoldingType.Cash
                ? holding.currentPrice.t.ToString()
                : holding.code;
            return new Finance(code, holding.holdingType);
        }

        // 用于存储
        [SugarColumn(DefaultValue = "0", ColumnDataType = "decimal(24,12)")]
        public decimal _currentPrice_v { get; set; } = 0;

        [SugarColumn(DefaultValue = "RMB", ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType _currentPrice_t { get; set; } = CurrencyType.RMB;
    }

    // 数据库中的枚举列尽量使用 MySQL ENUM 类型。
    static class MySqlEnumColumnTypes
    {
        public const string CurrencyType = "enum('RMB','USD','JPY','SGD','HKD')";
        public const string HoldingType = "enum('NASDAQ','ARCA','UST','SHANGHAI','CNFUND','Cash','Accrued')";
        public const string StatementImportProvider = "enum('IBKRReportMail','ICBCBillMail','WiseMail','OCBCMail','OCBCStatementMail','NexusDpMonthlyReport','PayPalMail','Manual')";
        public const string SnapshotSource = "enum('AutoDaily','Manual')";
        public const string SnapshotItemType = "enum('AccountBalance','Holding')";
        public const string AccountUsage = "enum('Life','Investment','Transit','Unknown')";
    }

    public enum AccountUsage
    {
        Life,
        Investment,
        Transit,
        Unknown
    }

    [SugarIndex("unique_AccountInternalIds_account_card_no", nameof(AccountInternalId._account_Id), OrderByType.Asc, nameof(AccountInternalId.cardNo), OrderByType.Asc, true)]
    [SugarTable("AccountInternalIds")]
    public class AccountInternalId
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(128)")]
        public string cardNo { get; set; } = "";

        [SugarColumn(DefaultValue = "''")]
        public string desc { get; set; } = "";

        // 只作为人工备注使用，匹配内部转账时不使用该列。
        [SugarColumn(IsNullable = true, ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType? currencyType { get; set; }

        [SugarColumn(IsIgnore = true)]
        public string sourceText { get; set; } = "";

        [SugarColumn(DefaultValue = "0")]
        public int _account_Id { get; set; } = 0;
    }

    // 账户。一个账户可以同时拥有多个币种余额，具体余额保存在 AccountBalances 中。
    // 主副卡关系按账户存储；副卡账户指向对应主卡账户。
    [SugarIndex("unique_Accounts_name", nameof(Account.name), OrderByType.Asc, true)]
    [SugarTable("Accounts")]
    public class Account
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        // 副卡账户指向对应主卡账户；为空表示该账户本身不是副卡。
        [Navigate(NavigateType.ManyToOne, nameof(_primaryAccount_Id), nameof(MyBook.Account.Id))]
        public Account? PrimaryAccount { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string name { get; set; } = "";

        [SugarColumn(DefaultValue = "''")]
        public string desc { get; set; } = "";

        [SugarColumn(IsNullable = true, ColumnDataType = "varchar(255)")]
        public string? email { get; set; } = null;

        [SugarColumn(DefaultValue = "1")]
        public bool relativeBalance { get; set; } = true; // 无法直接获取真实余额，只能根据变动值推算。

        // 仅用于 UI 统计分组，不影响余额、流水和导入逻辑。
        [SugarColumn(DefaultValue = "Life", ColumnDataType = MySqlEnumColumnTypes.AccountUsage, SqlParameterDbType = typeof(EnumToStringConvert))]
        public AccountUsage usage { get; set; } = AccountUsage.Life;

        [SugarColumn(IsNullable = true)]
        public int? _primaryAccount_Id { get; set; }
    }

    // 单个账户在一种币种下的余额。
    [SugarIndex("unique_AccountBalances_account_currency", nameof(AccountBalance._account_Id), OrderByType.Asc, nameof(AccountBalance.t), OrderByType.Asc, true)]
    [SugarTable("AccountBalances")]
    public class AccountBalance : Currency
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        public AccountBalance()
        {
        }

        public AccountBalance(Account account, Currency balance)
        {
            Account = account;
            CopyFrom(balance);
        }

        [SugarColumn(DefaultValue = "0")]
        public int _account_Id { get; set; } = 0;
    }
    // 无论出还是入，只记录本次变动所影响的账户，而不是记录Src和Dest账户
    // 一方面大多数交易是流向外部，不需要记录对方账户状况，只是有时需要记录对方账户名以区分原因
    // 一方面在自己的账户间的交易，出入金额可能不同(例如手续费、购汇)，记录麻烦
    // 而且主要目的是记录收支状况，完全可以忽略自己账户间的交易
    [SugarTable("Records")]
    public class Record : Currency // 收支记录
    {
        [SugarColumn(ColumnName = "_Currency_v", DefaultValue = "0", ColumnDataType = "decimal(24,12)")]
        public new decimal v
        {
            get => base.v;
            set => base.v = value;
        }

        [SugarColumn(ColumnName = "_Currency_t", DefaultValue = "RMB", ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public new CurrencyType t
        {
            get => base.t;
            set => base.t = value;
        }

        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_statementImport_Id), nameof(MyBook.StatementImport.Id))]
        public StatementImport? StatementImport { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(matchedRecordId), nameof(MyBook.Record.Id))]
        public Record? MatchedRecord { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string DestAccount { get; set; } = ""; // 对方账户描述

        [SugarColumn(DefaultValue = "0")]
        public bool isInternal { get; set; } = false; // 原始账单/解析逻辑直接确认的内部交易。

        [SugarColumn(IsNullable = true)]
        public int? matchedRecordId { get; set; } = null; // 跨账单一对一匹配到的另一侧内部交易记录。

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(1024)")]
        public string matchedRecordReason { get; set; } = ""; // 跨账单匹配依据，仅用英文原因标识。

        [SugarColumn(DefaultValue = "0")]
        public bool isRefundMatched { get; set; } = false; // 是否已匹配到对应退款/消费；默认不计入界面统计图表。

        [SugarColumn(DefaultValue = "0")]
        public int HoldingQuantity { get; set; } = 0; // 交易涉及的持仓数量，非持仓交易为 0。

        public DateTime date { get; set; } // 发生时间
        public DateTime updateTime { get; set; }
        // 表面交易金额，区别于记账金额，极少使用
        // 例如在Steam国区用visa外币卡购买100RMB的游戏，实际会换算成外币支出，而退款时也是把100RMB换算成外币退款，汇率变动可能导致消费和退款的外币金额不一致
        [SugarColumn(IsIgnore = true)]
        public Currency? DescCurrency
        {
            get
            {
                if (_descCurrency is null && _descCurrency_v.HasValue && _descCurrency_t.HasValue)
                    _descCurrency = new Currency(_descCurrency_v.Value, _descCurrency_t.Value);
                return _descCurrency;
            }
            set
            {
                _descCurrency = value;
                if (value is null)
                {
                    _descCurrency_v = null;
                    _descCurrency_t = null;
                    return;
                }

                _descCurrency_v = value.v;
                _descCurrency_t = value.t;
            }
        }

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(1024)")]
        public string Source { get; set; } = ""; // 获得该信息的途径

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(1024)")]
        public string Reason { get; set; } = ""; // 消费/收入原因

        [SugarColumn(IsNullable = true, ColumnDataType = "json")]
        public string? backup { get; set; } = null; // 非手动来源的记录首次手动编辑前的原始值。

        // 用于存储
        [SugarColumn(DefaultValue = "0")]
        public int _account_Id { get; set; } = 0;

        [SugarColumn(IsNullable = true, ColumnDataType = "decimal(24,12)")]
        public decimal? _descCurrency_v { get; set; }

        [SugarColumn(IsNullable = true, ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType? _descCurrency_t { get; set; }

        public int _statementImport_Id { get; set; }

        private Currency? _descCurrency;
    }
    public class Records : List<Record>
    {
    }

    public enum SnapshotSource
    {
        AutoDaily,
        Manual,
    }

    public enum SnapshotItemType
    {
        AccountBalance,
        Holding,
    }

    [SugarIndex("unique_Snapshots_source_key", nameof(Snapshot.source), OrderByType.Asc, nameof(Snapshot.snapshotKey), OrderByType.Asc, true)]
    [SugarTable("Snapshots")]
    public class Snapshot
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [SugarColumn(DefaultValue = "AutoDaily", ColumnDataType = MySqlEnumColumnTypes.SnapshotSource, SqlParameterDbType = typeof(EnumToStringConvert))]
        public SnapshotSource source { get; set; } = SnapshotSource.AutoDaily;

        [SugarColumn(ColumnDataType = "datetime(6)")]
        public DateTime time { get; set; }

        [SugarColumn(DefaultValue = "1")]
        public int schemaVersion { get; set; } = 1;

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(128)")]
        public string snapshotKey { get; set; } = "";

        [SugarColumn(ColumnDataType = "datetime(6)")]
        public DateTime createdAt { get; set; }
    }

    // SnapshotItems stores a stable indexed envelope plus versioned JSON payloads.
    // Old payload versions must stay readable even when current AccountBalance/Holding changes.
    [SugarIndex("index_SnapshotItems_snapshot_type_key", nameof(SnapshotItem._snapshot_Id), OrderByType.Asc, nameof(SnapshotItem.itemType), OrderByType.Asc, nameof(SnapshotItem.stableKey), OrderByType.Asc)]
    [SugarTable("SnapshotItems")]
    public class SnapshotItem
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_snapshot_Id), nameof(MyBook.Snapshot.Id))]
        public Snapshot? Snapshot { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        [SugarColumn(DefaultValue = "AccountBalance", ColumnDataType = MySqlEnumColumnTypes.SnapshotItemType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public SnapshotItemType itemType { get; set; } = SnapshotItemType.AccountBalance;

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(256)")]
        public string stableKey { get; set; } = "";

        [SugarColumn(DefaultValue = "''", ColumnDataType = "varchar(256)")]
        public string accountName { get; set; } = "";

        [SugarColumn(IsNullable = true, ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType? currencyType { get; set; }

        [SugarColumn(IsNullable = true, ColumnDataType = "decimal(24,12)")]
        public decimal? amount { get; set; }

        [SugarColumn(ColumnDataType = "longtext")]
        public string payloadJson { get; set; } = "";

        // 用于存储
        [SugarColumn(DefaultValue = "0")]
        public int _snapshot_Id { get; set; } = 0;

        [SugarColumn(IsNullable = true)]
        public int? _account_Id { get; set; }
    }

    public enum StatementImportProvider
    {
        IBKRReportMail,
        ICBCBillMail,
        WiseMail,
        OCBCMail,
        OCBCStatementMail,
        NexusDpMonthlyReport,
        PayPalMail,
        Manual,
    }

    [SugarIndex("unique_StatementImports_provider_time_key", nameof(StatementImport.provider), OrderByType.Asc, nameof(StatementImport.time), OrderByType.Asc, nameof(StatementImport.statementKey), OrderByType.Asc, true)]
    [SugarTable("StatementImports")]
    public class StatementImport
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [SugarColumn(DefaultValue = "Manual", ColumnDataType = MySqlEnumColumnTypes.StatementImportProvider, SqlParameterDbType = typeof(EnumToStringConvert))]
        public StatementImportProvider provider { get; set; } = StatementImportProvider.Manual;

        [SugarColumn(ColumnDataType = "datetime(6)")]
        public DateTime time { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string statementKey { get; set; } = "";
    }

    //币种
    public enum CurrencyType
    {
        RMB,
        USD,
        JPY,
        SGD,
        HKD,
    };
    // 任意币种*数量的组合
    class Money
    {
        Dictionary<CurrencyType, Currency> currencies = new();
    }
    // 单个币种
    public class Currency : IEquatable<Currency>
    {
        [SugarColumn(ColumnName = "amount", DefaultValue = "0", ColumnDataType = "decimal(24,12)")]
        public decimal v { get; set; } = 0; // 金额，decimal应当可以避免精度问题

        [SugarColumn(ColumnName = "currency_type", DefaultValue = "RMB", ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType t { get; set; } = CurrencyType.RMB;
        private static char[] seperator = new char[] { '/', '(', ')' };
        public void CopyFrom(Currency? c)
        {
            if (c == null) return;
            v= c.v;
            t = c.t;
        }
        public bool Equals(Currency? other)
        {
            return other is not null && v == other.v && t == other.t;
        }
        public override bool Equals(object? obj)
        {
            return Equals(obj as Currency);
        }
        public override int GetHashCode()
        {
            return HashCode.Combine(v, t);
        }
        public static bool operator ==(Currency? left, Currency? right)
        {
            return left is null ? right is null : left.Equals(right);
        }
        public static bool operator !=(Currency? left, Currency? right)
        {
            return !(left == right);
        }
        // 形如 123.2/RMB、123.2/RMB(存入) 或 123.2/RMB(支出)，支出会被解析为负数。
        public static Currency Parse(string _t)
        {
            var list = _t.Split(seperator);
            if(list.Length >= 2)
            {
                var currency = new Currency(list[0], list[1]);
                if (_t.EndsWith("(存入)", StringComparison.Ordinal))
                {
                    if (currency.v < 0)
                        throw new CurrencyParseException($"Parse Fail:{_t}");
                    currency.v = Math.Abs(currency.v);
                }
                else if (_t.EndsWith("(支出)", StringComparison.Ordinal))
                {
                    if (currency.v < 0)
                        throw new CurrencyParseException($"Parse Fail:{_t}");
                    currency.v = -Math.Abs(currency.v);
                }
                else if (_t.Contains('(') || _t.Contains(')'))
                    throw new CurrencyParseException($"Parse Fail:{_t}");
                return currency;
            }
            throw new CurrencyParseException($"Parse Fail:{_t}");
        }

        public Currency()
        {
        }
        public Currency(string _t) : this(0, _t)
        {
        }
        public Currency(decimal _v, string _t)
        {
            v = _v;
            t = ParseCurrencyType(_t);
        }
        public Currency(string _v, string _t)
        {
            v = decimal.Parse(_v, NumberStyles.Currency);
            t = ParseCurrencyType(_t);
        }
        public Currency(decimal _v, CurrencyType _t)
        {
            v = _v;
            t = _t;
        }
        public Currency(string _v, CurrencyType _t)
        {
            v = decimal.Parse(_v, NumberStyles.Currency);
            t = _t;
        }

        public static decimal RoundMoney(decimal value)
        {
            return Decimal.Round(value, 2, MidpointRounding.ToEven);
        }

        private static CurrencyType ParseCurrencyType(string value)
        {
            if (value == "CNY")
                return CurrencyType.RMB;

            if (!Enum.TryParse<CurrencyType>(value, out var currencyType))
                throw new ArgumentException($"不支持的币种 {value}");
            return currencyType;
        }
    }
}
