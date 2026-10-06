using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace MyBook;

internal sealed class CommandLineConsole : IDisposable
{
    private bool ownsConsole;

    internal void Open(bool interactive = true, bool utf8 = false)
    {
        if (interactive && GetConsoleWindow() == IntPtr.Zero && !AttachConsole(UInt32.MaxValue))
            ownsConsole = AllocConsole();
        if (utf8) Console.OutputEncoding = Encoding.UTF8;
        Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
    }

    public void Dispose()
    {
        if (!ownsConsole) return;
        Console.WriteLine("Press any key to close.");
        try { Console.ReadKey(intercept: true); } catch (InvalidOperationException) { }
        FreeConsole();
        ownsConsole = false;
    }

    [DllImport("kernel32.dll")] private static extern IntPtr GetConsoleWindow();
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool AttachConsole(uint processId);
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool AllocConsole();
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool FreeConsole();
}
