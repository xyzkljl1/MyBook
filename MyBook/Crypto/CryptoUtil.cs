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
        private readonly CryptoPriceUtil cryptoPrice;

        public CryptoUtil(IConfigurationRoot config, DatabaseUtil database, CryptoPriceUtil? cryptoPrice = null)
        {
            this.database = database;
            this.config = config;
            this.cryptoPrice = cryptoPrice ?? new CryptoPriceUtil();
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
