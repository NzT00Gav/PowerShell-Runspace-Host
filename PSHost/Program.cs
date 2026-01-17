using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;

namespace PSHost
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                Start();
                Banner();
                cmdline.RunCMDLine();
            }
            finally
            {
                cmdline.Cleanup();
            }
        }

        private static void Start()
        {
            Console.Title = "Custom PowerShell Host";
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;

            cmdline.runspace = RunspaceFactory.CreateRunspace();
            cmdline.runspace.Open();

            cmdline.ps = PowerShell.Create();
            cmdline.ps.Runspace = cmdline.runspace;

            Console.WriteLine("[+] PowerShell runspace initialized");
            Console.WriteLine();
        }

        internal static void Banner()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
            Console.WriteLine("║           CUSTOM POWERSHELL HOST v1.0                    ║");
            Console.WriteLine("║      System.Management.Automation Implementation         ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════╝");
            Console.ResetColor();
            Console.WriteLine();
        }
    }
}
