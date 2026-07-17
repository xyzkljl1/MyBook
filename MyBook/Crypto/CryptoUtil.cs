using Microsoft.Extensions.Configuration;

namespace MyBook
{
    partial class CryptoUtil
    {
        private readonly DatabaseUtil database;
        private readonly IConfigurationRoot config;

        public CryptoUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            this.database = database;
            this.config = config;
        }

        private string EtherscanApiKey => RequiredConfig(config, "etherscan_api_key");

        public Task FetchDailyReportsAsync(CancellationToken cancellationToken = default)
        {
            return FetchETHDailyReportsAsync(cancellationToken);
        }
    }
}
