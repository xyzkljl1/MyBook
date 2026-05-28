using Microsoft.Extensions.Configuration;
using System.Data;
using System.Globalization;
using System.IO;

namespace MyBook
{
    internal static partial class LocalDebugEntrypoint
    {
        internal sealed class Result
        {
            public bool Handled { get; set; }
            public int ExitCode { get; set; }
        }

        public static bool TryRun(string[] args, out int exitCode)
        {
            var result = new Result();
            TryRunTracked(args, result);
            if (!result.Handled)
                TryRunLocal(args, result);
            exitCode = result.ExitCode;
            return result.Handled;
        }

        private static void TryRunTracked(string[] args, Result result)
        {
            if (HasArgument(args, "--clean-database"))
            {
                Run(result, "CleanDatabase", () =>
                {
                    var config = LoadConfig();
                    var cleanupResult = new DatabaseUtil(config).CleanVolatileData(ReadCleanToSnapshotId(args));
                    WriteCleanupResult(cleanupResult);
                });
                return;
            }

            if (HasArgument(args, "--clean-wise-data"))
            {
                Run(result, "CleanWiseImportedData", () =>
                {
                    var config = LoadConfig();
                    new DatabaseUtil(config).CleanWiseImportedData();
                    Console.WriteLine("Clean Wise imported data done.");
                });
                return;
            }

            if (HasArgumentOrValue(args, "--debug-sql") || HasArgumentOrValue(args, "--debug-sql-file"))
            {
                Run(result, "Debug SQL", () =>
                {
                    var sql = ReadDebugSqlArgument(args);
                    var config = LoadConfig();
                    var database = new DatabaseUtil(config);
                    if (IsQueryDebugSql(sql))
                    {
                        WriteDebugSqlResult(database.QueryDebugSql(sql));
                    }
                    else
                    {
                        var affectedRows = database.ExecuteDebugSql(sql);
                        Console.WriteLine($"Debug SQL done: affectedRows={affectedRows}");
                    }
                });
                return;
            }

        }

        private static void Run(Result result, string name, Action action)
        {
            result.Handled = true;
            try
            {
                action();
            }
            catch (Exception exception)
            {
                result.ExitCode = 1;
                Console.WriteLine($"{name} failed: {exception.Message}");
            }
        }

        private static IConfigurationRoot LoadConfig()
        {
            return new ConfigurationBuilder().AddJsonFile("config.json", false).Build();
        }

        private static bool HasArgument(string[] args, string name)
        {
            return args.Any(arg => arg.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        private static bool HasArgumentOrValue(string[] args, string name)
        {
            var prefix = name + "=";
            return args.Any(arg => arg.Equals(name, StringComparison.OrdinalIgnoreCase)
                || arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        private static int? ReadCleanToSnapshotId(string[] args)
        {
            var value = GetArgumentValue(args, "--clean-to-snapshot");
            if (String.IsNullOrWhiteSpace(value))
                return null;
            if (!Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var snapshotId) || snapshotId <= 0)
                throw new ArgumentException($"Invalid --clean-to-snapshot value: {value}");
            return snapshotId;
        }

        private static string? GetArgumentValue(string[] args, string name)
        {
            for (var i = 0; i < args.Length; i++)
            {
                if (args[i].Equals(name, StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                    return args[i + 1];

                var prefix = name + "=";
                if (args[i].StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return args[i][prefix.Length..];
            }

            return null;
        }

        private static void WriteCleanupResult(DatabaseCleanupResult result)
        {
            Console.WriteLine("Before: " + FormatCounts(result.BeforeCounts));
            Console.WriteLine("After: " + FormatCounts(result.AfterCounts));
            Console.WriteLine("Fixed StatementImports:");
            foreach (var statementImport in result.FixedStatementImports)
                Console.WriteLine($"{statementImport.Id} {statementImport.provider} {statementImport.time:yyyy-MM-dd HH:mm:ss.ffffff}");
        }

        private static string FormatCounts(Dictionary<string, int> counts)
        {
            return String.Join(", ", counts.Select(item => $"{item.Key}={item.Value}"));
        }

        private static string ReadDebugSqlArgument(string[] args)
        {
            var file = GetArgumentValue(args, "--debug-sql-file");
            if (!String.IsNullOrWhiteSpace(file))
                return File.ReadAllText(file);

            var sql = GetArgumentValue(args, "--debug-sql");
            if (!String.IsNullOrWhiteSpace(sql))
                return sql;

            throw new ArgumentException("Missing --debug-sql or --debug-sql-file value.");
        }

        private static bool IsQueryDebugSql(string sql)
        {
            var trimmed = TrimSqlLeadingTrivia(sql);
            return trimmed.StartsWith("select", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("show", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("describe", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("desc", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("with", StringComparison.OrdinalIgnoreCase);
        }

        private static string TrimSqlLeadingTrivia(string sql)
        {
            var current = sql.TrimStart();
            while (true)
            {
                if (current.StartsWith("--", StringComparison.Ordinal))
                {
                    var newline = current.IndexOf('\n');
                    current = newline < 0 ? "" : current[(newline + 1)..].TrimStart();
                    continue;
                }

                if (current.StartsWith("/*", StringComparison.Ordinal))
                {
                    var end = current.IndexOf("*/", StringComparison.Ordinal);
                    current = end < 0 ? "" : current[(end + 2)..].TrimStart();
                    continue;
                }

                return current;
            }
        }

        private static void WriteDebugSqlResult(DataTable table)
        {
            Console.WriteLine($"Rows: {table.Rows.Count}");
            if (table.Columns.Count == 0)
                return;

            Console.WriteLine(String.Join("\t", table.Columns.Cast<DataColumn>().Select(column => column.ColumnName)));
            foreach (DataRow row in table.Rows)
            {
                Console.WriteLine(String.Join(
                    "\t",
                    table.Columns.Cast<DataColumn>().Select(column => FormatDebugSqlValue(row[column]))));
            }
        }

        private static string FormatDebugSqlValue(object? value)
        {
            if (value is null || value == DBNull.Value)
                return "NULL";
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss.ffffff", CultureInfo.InvariantCulture);
            if (value is byte[] bytes)
                return Convert.ToHexString(bytes);

            return Convert.ToString(value, CultureInfo.InvariantCulture)?
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "    ")
                ?? "";
        }

        static partial void TryRunLocal(string[] args, Result result);
    }
}
