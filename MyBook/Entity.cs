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
    public enum StockType
    {
        // 纳斯达克交易所上市资产。
        NASDAQ,
        // 美国国债。
        UST,
        // 上海交易所上市资产。
        SHANGHAI,
        // 国内基金。
        CNFUND,
        // 现金类持仓。
        Cash
    };

    // 股票、基金、现金等持仓状态，使用 code + stockType 区分具体资产。
    [SugarIndex("unique_Stocks_account_code_type", nameof(Stock._account_Id), OrderByType.Asc, nameof(Stock.code), OrderByType.Asc, nameof(Stock.stockType), OrderByType.Asc, true)]
    [SugarTable("Stocks")]
    public class Stock
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string code { get; set; } = "";

        [SugarColumn(DefaultValue = "NASDAQ", ColumnDataType = MySqlEnumColumnTypes.StockType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public StockType stockType { get; set; } = StockType.NASDAQ;

        [SugarColumn(DefaultValue = "0")]
        public decimal quantity { get; set; } = 0;

        [SugarColumn(DefaultValue = "''")]
        public string desc { get; set; } = "";

        [SugarColumn(IsIgnore = true)]
        // 当前单价，金额和币种分别存储。
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
        public DateTime currentPriceTime { get; set; }

        public Stock()
        {
        }

        public Stock(string _c, StockType _t)
        {
            code = _c;
            stockType = _t;
        }

        // 用于存储
        [SugarColumn(DefaultValue = "0")]
        public decimal _currentPrice_v { get; set; } = 0;

        [SugarColumn(DefaultValue = "RMB", ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType _currentPrice_t { get; set; } = CurrencyType.RMB;

        [SugarColumn(IsNullable = true)]
        public int? _account_Id { get; set; }
    }

    // 数据库中的枚举列尽量使用 MySQL ENUM 类型。
    static class MySqlEnumColumnTypes
    {
        public const string CurrencyType = "enum('RMB','USD','JPY','SGD','HKD')";
        public const string StockType = "enum('NASDAQ','UST','SHANGHAI','CNFUND','Cash')";
    }

    // 账户
    // 一个账户下的不同币种余额视作多个账户。
    // Account.v 表示该账户中现金、股票、负债等所有种类资产的总和余额。
    [SugarIndex("unique_Accounts_name_currency", nameof(Account.name), OrderByType.Asc, nameof(Account._v_t), OrderByType.Asc, true)]
    [SugarTable("Accounts")]
    public class Account
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [SugarColumn(IsIgnore = true)]
        public Currency v
        {
            get { return new Currency(_v_v, _v_t); }
            set
            {
                _v_v = value.v;
                _v_t = value.t;
            }
        } // 余额

        [SugarColumn(DefaultValue = "''")]
        public string name { get; set; } = "";

        [SugarColumn(DefaultValue = "''")]
        public string desc { get; set; } = "";

        [SugarColumn(DefaultValue = "0")]
        public decimal _v_v { get; set; } = 0;

        [SugarColumn(DefaultValue = "RMB", ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType _v_t { get; set; } = CurrencyType.RMB;
    }
    // 无论出还是入，只记录本次变动所影响的账户，而不是记录Src和Dest账户
    // 一方面大多数交易是流向外部，不需要记录对方账户状况，只是有时需要记录对方账户名以区分原因
    // 一方面在自己的账户间的交易，出入金额可能不同(例如手续费、购汇)，记录麻烦
    // 而且主要目的是记录收支状况，完全可以忽略自己账户间的交易
    [SugarTable("Records")]
    public class Record : Currency // 收支记录
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [Navigate(NavigateType.ManyToOne, nameof(_account_Id), nameof(MyBook.Account.Id))]
        public Account? Account { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string DestAccount { get; set; } = ""; // 对方账户描述

        [SugarColumn(DefaultValue = "0")]
        public bool isIn {  get; set; } = false; // 存入/支出

        [SugarColumn(DefaultValue = "0")]
        public bool isInternal { get; set; } = false; // 是否自己账户间的交易
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

        // 用于存储
        [SugarColumn(IsNullable = true)]
        public int? _account_Id { get; set; }

        [SugarColumn(IsNullable = true)]
        public decimal? _descCurrency_v { get; set; }

        [SugarColumn(IsNullable = true, ColumnDataType = MySqlEnumColumnTypes.CurrencyType, SqlParameterDbType = typeof(EnumToStringConvert))]
        public CurrencyType? _descCurrency_t { get; set; }

        private Currency? _descCurrency;
    }
    public class Records : List<Record>
    {
    }

    [SugarIndex("unique_StatementImports_provider_month", nameof(StatementImport.provider), OrderByType.Asc, nameof(StatementImport.month), OrderByType.Asc, true)]
    [SugarTable("StatementImports")]
    public class StatementImport
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)]
        public int Id { get; set; }

        [SugarColumn(DefaultValue = "''")]
        public string provider { get; set; } = "";

        [SugarColumn(DefaultValue = "''")]
        public string month { get; set; } = "";
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
        [SugarColumn(ColumnName = "amount", DefaultValue = "0")]
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
        // 形如 123.2/RMB 或 123.2/RMB(存入)
        public static Currency Parse(string _t)
        {
            var list = _t.Split(seperator);
            if(list.Length >= 2)
                return new Currency(list[0], list[1]);
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
            if (!Enum.TryParse<CurrencyType>(_t, out var currencyType))
                throw new ArgumentException($"不支持的币种 {_t}");
            t = currencyType;
        }
        public Currency(string _v, string _t)
        {
            v = decimal.Parse(_v, NumberStyles.Currency);
            if (!Enum.TryParse<CurrencyType>(_t, out var currencyType))
                throw new ArgumentException($"不支持的币种 {_t}");
            t = currencyType;
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
    }
}
