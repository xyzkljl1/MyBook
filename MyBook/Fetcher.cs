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
        IConfigurationRoot? config;
        MailUtil? mail;
        StockUtil? stock;
        DatabaseUtil? database;
        public Fetcher()
        {
        }
        public void RunSchedule()
        {
            config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            database = new(config);
            mail = new(config, database);
            stock = new(config, database);
            _ = mail.FetchICBCBills();
            //stock.Fetch(new Finance("QQQ", StockType.NASDAQ));
            //stock.Fetch(new Finance("021282", StockType.CNFUND));
        }
    }
}
