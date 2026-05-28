using Microsoft.Extensions.Configuration;

namespace MyBook
{
    // Shared file import entry point; source-specific logic lives in suffix files.
    partial class FileUtil
    {
        private readonly IConfigurationRoot config;
        private readonly DatabaseUtil database;

        public FileUtil(IConfigurationRoot config, DatabaseUtil database)
        {
            this.config = config;
            this.database = database;
        }
    }
}
