using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Text;

namespace PSHost
{
    internal static class cmdline
    {
        [DllImport("kernel32.dll")]
        private static extern void Sleep(uint dwMilliseconds);

        internal static Runspace runspace;
        internal static PowerShell ps;

        private static readonly List<string> CommandHistory = new List<string>();
        private static readonly string[] BuiltInCommands = { "clear", "cls", "history", "help", "exit", "quit" };

        private static int _inputLine;
        private static int _inputStartCol;
        private static int _historyIndex = -1;
        private static volatile bool _shouldExit;

        internal static void RunCMDLine()
        {
            while (!_shouldExit)
            {
                try
                {
                    DisplayPrompt();

                    string input = ReadLineWithHistory();
                    if (string.IsNullOrWhiteSpace(input))
                        continue;

                    if (HandleInternalCommand(input))
                        continue;

                    AddToHistory(input);
                    ExecutePSHCommand(input);
                }
                catch (Exception ex)
                {
                    WriteError($"Fatal Error: {ex.Message}");
                }
            }
        }

        private static void AddToHistory(string input)
        {
            CommandHistory.Add(input);
            _historyIndex = CommandHistory.Count;
        }

        private static int DisplayPrompt()
        {
            string cwd = runspace.SessionStateProxy.Path.CurrentFileSystemLocation.Path;

            if (cwd.Length > 40)
                cwd = "..." + cwd.Substring(cwd.Length - 37);

            Console.Write($"[PSHost] {cwd}> ");
            return Console.CursorLeft;
        }

        private static void RedrawInputLine(string text, int cursorPos)
        {
            cursorPos = Math.Max(0, Math.Min(cursorPos, text.Length));

            Console.SetCursorPosition(0, _inputLine);

            int cols = Console.BufferWidth;
            Console.Write(new string(' ', Math.Max(0, cols - 1)));

            Console.SetCursorPosition(0, _inputLine);

            _inputStartCol = DisplayPrompt();
            Console.Write(text);

            Console.SetCursorPosition(_inputStartCol + cursorPos, _inputLine);
        }

        private static string ReadLineWithHistory()
        {
            var input = new StringBuilder();
            int cursorPos = 0;

            _historyIndex = CommandHistory.Count;

            _inputLine = Console.CursorTop;
            _inputStartCol = Console.CursorLeft;

            while (true)
            {
                var key = Console.ReadKey(intercept: true);

                switch (key.Key)
                {
                    case ConsoleKey.Enter:
                        Console.WriteLine();
                        _historyIndex = CommandHistory.Count;
                        return input.ToString();

                    case ConsoleKey.Escape:
                        input.Clear();
                        cursorPos = 0;
                        _historyIndex = CommandHistory.Count;
                        RedrawInputLine("", 0);
                        break;

                    case ConsoleKey.Backspace:
                        if (cursorPos > 0)
                        {
                            input.Remove(cursorPos - 1, 1);
                            cursorPos--;
                            RedrawInputLine(input.ToString(), cursorPos);
                        }
                        break;

                    case ConsoleKey.Delete:
                        if (cursorPos < input.Length)
                        {
                            input.Remove(cursorPos, 1);
                            RedrawInputLine(input.ToString(), cursorPos);
                        }
                        break;

                    case ConsoleKey.LeftArrow:
                        if (cursorPos > 0)
                        {
                            cursorPos--;
                            Console.SetCursorPosition(_inputStartCol + cursorPos, _inputLine);
                        }
                        break;

                    case ConsoleKey.RightArrow:
                        if (cursorPos < input.Length)
                        {
                            cursorPos++;
                            Console.SetCursorPosition(_inputStartCol + cursorPos, _inputLine);
                        }
                        break;

                    case ConsoleKey.UpArrow:
                        HistoryUp(input, ref cursorPos);
                        break;

                    case ConsoleKey.DownArrow:
                        HistoryDown(input, ref cursorPos);
                        break;

                    case ConsoleKey.Tab:
                        AutoCompleteBuiltInCommands(input, ref cursorPos);
                        break;

                    default:
                        if (!char.IsControl(key.KeyChar))
                        {
                            input.Insert(cursorPos, key.KeyChar);
                            cursorPos++;
                            RedrawInputLine(input.ToString(), cursorPos);
                        }
                        break;
                }
            }
        }

        private static void HistoryUp(StringBuilder input, ref int cursorPos)
        {
            if (CommandHistory.Count == 0 || _historyIndex <= 0)
                return;

            _historyIndex--;
            input.Clear();
            input.Append(CommandHistory[_historyIndex]);
            cursorPos = input.Length;
            RedrawInputLine(input.ToString(), cursorPos);
        }

        private static void HistoryDown(StringBuilder input, ref int cursorPos)
        {
            if (CommandHistory.Count == 0)
                return;

            if (_historyIndex < CommandHistory.Count - 1)
            {
                _historyIndex++;
                input.Clear();
                input.Append(CommandHistory[_historyIndex]);
                cursorPos = input.Length;
                RedrawInputLine(input.ToString(), cursorPos);
                return;
            }

            if (_historyIndex == CommandHistory.Count - 1)
            {
                _historyIndex = CommandHistory.Count;
                input.Clear();
                cursorPos = 0;
                RedrawInputLine("", 0);
            }
        }

        private static void AutoCompleteBuiltInCommands(StringBuilder input, ref int cursorPos)
        {
            string partial = input.ToString();

            foreach (var cmd in BuiltInCommands)
            {
                if (cmd.StartsWith(partial, StringComparison.OrdinalIgnoreCase))
                {
                    input.Clear();
                    input.Append(cmd);
                    cursorPos = cmd.Length;
                    RedrawInputLine(cmd, cursorPos);
                    break;
                }
            }
        }

        private static bool HandleInternalCommand(string input)
        {
            string cmd = input.Trim().ToLowerInvariant();

            switch (cmd)
            {
                case "clear":
                case "cls":
                    Console.Clear();
                    Program.Banner();
                    return true;

                case "history":
                    ShowCommandHistory();
                    return true;

                case "exit":
                case "quit":
                    _shouldExit = true;
                    return true;

                default:
                    return false;
            }
        }

        private static void ShowCommandHistory()
        {
            if (CommandHistory.Count == 0)
            {
                Console.WriteLine("No command history.");
                return;
            }

            Console.WriteLine();
            Console.WriteLine("Command History:");
            Console.WriteLine(new string('─', 60));

            for (int i = 0; i < CommandHistory.Count; i++)
            {
                WriteColored($"{i + 1,4}: ", ConsoleColor.DarkGray);
                Console.WriteLine(CommandHistory[i]);
            }

            Console.WriteLine(new string('─', 60));
        }

        private static void ExecutePSHCommand(string command)
        {
            try
            {
                ps.AddScript(command);

                var results = ps.Invoke();

                WriteFormatted(results);
                WriteStreams();
            }
            catch (RuntimeException rex)
            {
                WriteError($"PowerShell Runtime Error: {rex.Message}");
            }
            catch (Exception ex)
            {
                WriteError($"Error: {ex.Message}");
            }
            finally
            {
                ps.Commands.Clear();
            }
        }

        private static void WriteStreams()
        {
            if (ps.Streams.Error.Count > 0)
            {
                WriteColoredErrors(ps.Streams.Error);
                ps.Streams.Error.Clear();
            }

            if (ps.Streams.Warning.Count > 0)
            {
                WriteColoredWarnings(ps.Streams.Warning);
                ps.Streams.Warning.Clear();
            }
        }

        private static int SafeWidth()
        {
            try
            {
                return Math.Max(40, Console.BufferWidth - 1);
            }
            catch
            {
                return 200;
            }
        }

        private static void WriteFormatted(ICollection<PSObject> results)
        {
            if (results == null || results.Count == 0)
                return;

            int width = SafeWidth();

            using (var fmt = PowerShell.Create())
            {
                fmt.Runspace = runspace;

                fmt.AddCommand("Out-String")
                    .AddParameter("Width", width)
                    .AddParameter("Stream")
                    .AddParameter("InputObject", results);

                var lines = fmt.Invoke();

                foreach (var line in lines)
                {
                    var s = line?.BaseObject?.ToString();
                    if (!string.IsNullOrEmpty(s))
                        Console.WriteLine(s);
                }

                if (fmt.Streams.Error.Count > 0)
                {
                    WriteColoredErrors(fmt.Streams.Error);
                    fmt.Streams.Error.Clear();
                }
            }
        }

        private static void WriteColored(string message, ConsoleColor color)
        {
            var original = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.Write(message);
            Console.ForegroundColor = original;
        }

        private static void WriteColoredErrors(ICollection<ErrorRecord> errors)
        {
            WriteColoredCollection(errors, ConsoleColor.Red, prefix: "ERROR: ");
        }

        private static void WriteColoredWarnings(ICollection<WarningRecord> warnings)
        {
            WriteColoredCollection(warnings, ConsoleColor.Yellow, prefix: "WARNING: ");
        }

        private static void WriteColoredCollection<T>(ICollection<T> items, ConsoleColor color, string prefix)
        {
            var original = Console.ForegroundColor;
            Console.ForegroundColor = color;

            foreach (var item in items)
                Console.WriteLine($"{prefix}{item}");

            Console.ForegroundColor = original;
        }

        private static void WriteError(string message)
        {
            WriteColored(message + Environment.NewLine, ConsoleColor.Red);
        }

        internal static void Cleanup()
        {
            try
            {
                ps?.Dispose();
                runspace?.Close();

                Console.WriteLine();
                Console.WriteLine("[+] Resources cleaned up");
                Console.WriteLine("[+] Goodbye!");
                Sleep(2000);
            }
            catch
            {
            }
        }
    }
}
