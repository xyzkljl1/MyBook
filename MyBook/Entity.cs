using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyBook
{
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
    class Account
    {
        public Currency v { get; set; } = new Currency(0, CurrencyType.RMB);// 余额
        public string name { get; set; } = "";
    }
    class Record : Currency // 收支记录
    {
        Account Account { get; set; }
        DateTime date { get; set; } // 发生时间
        DateTime updateTime { get; set; }
        string Source { get; set; } = ""; // 获得该信息的途径
        string Reason { get; set; } = ""; // 消费/收入原因
    }
    class Records : List<Record>
    {
    }

    //币种
    enum CurrencyType
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
    class Currency
    {
        public decimal v = 0; // 金额，decimal应当可以避免精度问题
        public CurrencyType t = CurrencyType.RMB;
        public Currency()
        {
        }
        public Currency(string _t) : this(0, _t)
        {
        }
        public Currency(decimal _v, string _t)
        {
            v = _v;
            if (!Enum.TryParse<CurrencyType>(_t, out t))
                throw new ArgumentException($"不支持的币种 {_t}");
        }
        public Currency(string _v, string _t)
        {
            v = decimal.Parse(_v, NumberStyles.Currency);
            if (!Enum.TryParse<CurrencyType>(_t, out t))
                throw new ArgumentException($"不支持的币种 {_t}");
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
