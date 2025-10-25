This script elevates itself (UAC) if needed and upgrades installed applications via winget.
It writes a transcript log and pre/post package snapshots to the current script directory by default
(or a custom directory with -OutDir). It prints an elevated-session banner, a concise system summary,
and a completion line with status, timestamps, and duration.

Key features:
- Automatic admin elevation with UAC prompt
- Transcript logging (timestamped) and package snapshots (pre/post) as JSON
- System summary: OS, uptime, RAM, CPU, disk, locale
- Robust winget discovery and error handling
- Output directory control with -OutDir (defaults to script directory)

.PARAMETER OutDir
Optional. Directory where logs and package snapshots will be saved. Defaults to the current script directory.
If not writable, the script falls back to %TEMP%.

.PARAMETER Help
Optional. Show detailed help and examples, then exit.

.EXAMPLE
pwsh -File "C:\Users\Ofer\Desktop\Windows-Upgrade-All-Apps.ps1"
Runs with default settings, saving outputs next to the script.

.EXAMPLE
pwsh -File "C:\Users\Ofer\Desktop\Windows-Upgrade-All-Apps.ps1" -OutDir "D:\UpgradeLogs"
Saves log and snapshots into D:\UpgradeLogs.

.NOTES
Requires Windows 10/11 with "App Installer" (winget). PowerShell 5.1+ or 7+.
#>
