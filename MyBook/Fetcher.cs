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
        public Fetcher()
        {
        }
        public void RunSchedule()
        {
            config = new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
            mail = new(config);
            mail.SearchICBCBill(DateTime.Now.AddMonths(-2));
        }
    }
}
