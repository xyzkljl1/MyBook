using System;
using System.Collections.Generic;
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
}
