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
            TryRunLocal(args, result);
            exitCode = result.ExitCode;
            return result.Handled;
        }

        static partial void TryRunLocal(string[] args, Result result);
    }
}
