using Microsoft.Extensions.Configuration;
using System.Net.Http;

namespace MyBook
{
    partial class CryptoUtil
    {
        private static readonly TimeSpan CryptoRequestTimeout = TimeSpan.FromSeconds(20);
        private static readonly HttpClient sharedHttpClient = CreateSharedHttpClient();
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

        private static HttpClient CreateSharedHttpClient()
        {
            var client = new HttpClient { Timeout = CryptoRequestTimeout };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("MyBook/1.0 CryptoUtil");
            return client;
        }
    }
}
