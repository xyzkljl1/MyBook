using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyBook
{
    class Fetcher
    {
        IConfigurationRoot config;
        MailUtil mail;
        StockUtil stock;
        public Fetcher()
        {
        }
        public void RunSchedule()
        {
            config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            mail = new(config);
            stock = new(config);
            //mail.SearchICBCBill(DateTime.Now.AddMonths(-2));
            stock.Fetch(new Stock("QQQ", StockType.US));
            stock.Fetch(new Stock("021282", StockType.CNFUND));
        }
    }
}
