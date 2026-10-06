using Microsoft.Extensions.Configuration;
using QRCoder;

namespace MyBook;

partial class WebUtil
{
    internal static class BilibiliLogin
    {
        internal static int Run(string[] args)
        {
            using var console = new CommandLineConsole();
            using var cancellation = new CancellationTokenSource();
            ConsoleCancelEventHandler cancel = (_, e) => { e.Cancel = true; cancellation.Cancel(); };
            try
            {
                console.Open(utf8: true);
                if (args.Length != 1 || !args[0].Equals("--bilibili-login", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Usage: MyBook.exe --bilibili-login");
                    return 1;
                }
                if (Console.IsOutputRedirected)
                {
                    Console.WriteLine("Bilibili QR login requires an interactive console; do not redirect its output.");
                    return 1;
                }
                Console.CancelKeyPress += cancel;
                var config = new ConfigurationBuilder().SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("config.json", false).Build();
                var web = new WebUtil(config, new DatabaseUtil(config));
                Console.WriteLine("Scan with the Bilibili app and approve login. Press Ctrl+C to cancel.");
                string? lastStatus = null;
                Task.Run(() => web.LoginBilibiliAsync(url =>
                {
                    var foreground = Console.ForegroundColor;
                    var background = Console.BackgroundColor;
                    try
                    {
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.BackgroundColor = ConsoleColor.White;
                        Console.WriteLine(CreateQrCode(url));
                    }
                    finally
                    {
                        Console.ForegroundColor = foreground;
                        Console.BackgroundColor = background;
                    }
                    return Task.CompletedTask;
                }, status =>
                {
                    if (status == lastStatus) return;
                    lastStatus = status;
                    Console.WriteLine(status);
                }, cancellation.Token)).GetAwaiter().GetResult();
                Console.WriteLine("Bilibili login succeeded; session stored in database.");
                return 0;
            }
            catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
            {
                Console.WriteLine("Bilibili login cancelled.");
                return 130;
            }
            catch (Exception exception)
            {
                var message = exception.Message;
                var safe = exception is BilibiliException || exception is InvalidOperationException &&
                    (message.StartsWith("Bilibili session database", StringComparison.Ordinal)
                     || message == "Bilibili session lock lost; requests stopped.");
                Console.WriteLine("Bilibili login failed: " + (safe ? message : exception.GetType().Name));
                return 1;
            }
            finally
            {
                Console.CancelKeyPress -= cancel;
            }
        }

        internal static string CreateQrCode(string url)
        {
            using var data = QRCodeGenerator.GenerateQrCode(url, QRCodeGenerator.ECCLevel.M);
            using var qrCode = new AsciiQRCode(data);
            // Render black modules on the console's temporary white background.
            return qrCode.GetGraphicSmall(drawQuietZones: true, invert: true);
        }

    }
}
