using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyBook
{
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
        public Currency(string _t):this(0,_t)
        {
        }
        public Currency(decimal _v,string _t)
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
