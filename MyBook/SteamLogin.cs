using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Configuration;
using SteamKit2.Authentication;

namespace MyBook;

internal static class SteamLogin
{
    internal static int Run(string[] args)
    {
        var ownConsole = false;
        try
        {
            if (!args.Contains("--saved") && GetConsoleWindow() == IntPtr.Zero && !AttachConsole(UInt32.MaxValue))
                ownConsole = AllocConsole();
            Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
            if (args.Length is < 1 or > 2 || args.Any(arg => arg is not ("--steam-login" or "--saved"))
                || args.Distinct(StringComparer.Ordinal).Count() != args.Length)
                throw new ArgumentException("Usage: MyBook.exe --steam-login [--saved]");

            var config = new ConfigurationBuilder().SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("config.json", false).Build();
            var web = new WebUtil(config, new DatabaseUtil(config));
            string? username = null, password = null;
            IAuthenticator? guard = null;
            if (!args.Contains("--saved"))
            {
                username = ReadInput("Steam login name: ", secret: false);
                password = ReadInput("Password: ");
                guard = new GuardInput();
            }
            Console.WriteLine("Connecting to Steam...");
            var result = Task.Run(() => web.LoginSteamAsync(username, password, guard)).GetAwaiter().GetResult();
            Console.WriteLine("Steam login succeeded; session stored in database.");
            if (result.HasWallet)
                Console.WriteLine($"Wallet: {result.Balance:0.00} {result.Currency}; pending: {result.PendingBalance:0.00} {result.Currency}.");
            else Console.WriteLine("No Steam wallet.");
            return 0;
        }
        catch (Exception exception)
        {
            // Print only messages produced by our own Steam handlers, never credentials or server bodies.
            var message = exception.Message;
            var safe = exception is InvalidOperationException &&
                (message.StartsWith("Steam ClientLogOn:", StringComparison.Ordinal)
                || message.StartsWith("Steam Authentication:", StringComparison.Ordinal)
                || message.StartsWith("Steam session database", StringComparison.Ordinal)
                || message == "Steam session lock lost; requests stopped."
                || message == "Steam Guard code rejected."
                || message == "Steam login/account query timed out or disconnected."
                || message == "No saved Steam session; complete initial authorization first.");
            Console.WriteLine("Steam login failed: " + (safe ? message : exception.GetType().Name));
            return 1;
        }
        finally
        {
            if (ownConsole)
            {
                Console.WriteLine("Press any key to close.");
                try { Console.ReadKey(intercept: true); } catch (InvalidOperationException) { }
                FreeConsole();
            }
        }
    }

    private static string ReadInput(string prompt, bool secret = true)
    {
        if (Console.IsInputRedirected) throw new InvalidOperationException("Interactive terminal required.");
        Console.Write(prompt);
        var text = new StringBuilder();
        while (true)
        {
            var key = Console.ReadKey(intercept: true);
            if (key.Key == ConsoleKey.Enter) break;
            if (key.Key == ConsoleKey.Backspace)
            {
                if (text.Length == 0) continue;
                text.Length--;
                if (!secret) Console.Write("\b \b");
            }
            else if (!Char.IsControl(key.KeyChar))
            {
                text.Append(key.KeyChar);
                if (!secret) Console.Write(key.KeyChar);
            }
        }
        Console.WriteLine();
        return text.Length > 0 ? text.ToString() : throw new ArgumentException("Empty Steam input.");
    }

    private sealed class GuardInput : IAuthenticator
    {
        public Task<string> GetDeviceCodeAsync(bool previousCodeWasIncorrect) => ReadCode(previousCodeWasIncorrect);
        public Task<string> GetEmailCodeAsync(string email, bool previousCodeWasIncorrect) => ReadCode(previousCodeWasIncorrect);
        private static Task<string> ReadCode(bool rejected)
        {
            if (rejected) throw new InvalidOperationException("Steam Guard code rejected.");
            return Task.FromResult(ReadInput("Steam Guard code: "));
        }
        public Task<bool> AcceptDeviceConfirmationAsync()
        {
            Console.WriteLine("Approve this login in the Steam mobile app.");
            return Task.FromResult(true);
        }
    }

    [DllImport("kernel32.dll")] private static extern IntPtr GetConsoleWindow();
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool AttachConsole(uint processId);
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool AllocConsole();
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool FreeConsole();
}
