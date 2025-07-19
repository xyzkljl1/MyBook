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
        int v = 0;//单位：分,用decimal也不会有问题？
        CurrencyType t = CurrencyType.RMB;
        public Currency()
        {
        }
        public Currency(string _t):this(0,_t)
        {
        }
        public Currency(int _v,string _t)
        {
            v = _v;
            if (!Enum.TryParse<CurrencyType>(_t, out t))
                throw new ArgumentException($"不支持的币种 {_t}");
        }
        public Currency(string _v, string _t)
        {
            decimal value = decimal.Parse(_v, NumberStyles.Currency);
            v = (int)(value * 100);
            if (!Enum.TryParse<CurrencyType>(_t, out t))
                throw new ArgumentException($"不支持的币种 {_t}");
        }
    }
}
