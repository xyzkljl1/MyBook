using Microsoft.Extensions.Configuration;

namespace MyBook
{
    // Shared web scraping entry point; site-specific logic lives in suffix files.
    partial class WebUtil
    {
        private readonly IConfigurationRoot config;
        private readonly DatabaseUtil database;

        public WebUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            this.config = config;
            this.database = database;
        }
    }
}
