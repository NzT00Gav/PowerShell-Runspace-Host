# PowerShell-Runspace-Host (C#)

Minimal interactive PowerShell Host built in C# using **System.Management.Automation** + **Runspaces**.
It includes a custom console line editor with history navigation and basic built-in commands.

## Features
- PowerShell Runspace initialization (SMA)
- Custom prompt with current working directory (truncated when long)
- Line editing: left/right arrows, delete/backspace
- History navigation: up/down arrows
- Built-in commands: clear/cls/history/exit/quit (+ tab autocomplete)
- Output formatting using `Out-String` with safe console width
- Colored error/warning stream rendering

# Reference
https://github.com/hiatus/PowerSpace
