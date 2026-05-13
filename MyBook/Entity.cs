using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SqlSugar;

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
    enum StockType
    {
        US,
        SHANGHAI,
        CNFUND
    };
    class Stock
    {
        public StockType t;
        public string code;
        public Stock(string _c, StockType _t)
        {
            code = _c;
            t = _t;
        }
    }
    // 账户
    // 一个账户下的不同余额视作多个账户
    [SugarIndex("unique_Accounts_name", nameof(Account.name), OrderByType.Asc, true)]
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

        [SugarColumn(DefaultValue = "0")]
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

        [SugarColumn(DefaultValue = "''")]
        public string Source { get; set; } = ""; // 获得该信息的途径

        [SugarColumn(DefaultValue = "''")]
        public string Reason { get; set; } = ""; // 消费/收入原因

        // 用于存储
        public int? _account_Id { get; set; }
        public decimal? _descCurrency_v { get; set; }
        public CurrencyType? _descCurrency_t { get; set; }

        private Currency? _descCurrency;
    }
    public class Records : List<Record>
    {
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

        [SugarColumn(ColumnName = "currency_type", DefaultValue = "0")]
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
