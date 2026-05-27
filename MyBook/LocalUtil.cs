using Microsoft.Extensions.Configuration;

namespace MyBook
{
    // Shared local-file import entry point; source-specific logic lives in suffix files.
    partial class LocalUtil
    {
        private readonly IConfigurationRoot config;
        private readonly DatabaseUtil database;

        public LocalUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            this.config = config;
            this.database = database;
        }
    }
}
