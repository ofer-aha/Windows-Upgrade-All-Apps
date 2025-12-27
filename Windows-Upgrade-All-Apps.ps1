#Requires -RunAsAdministrator

<############################################################################
# Windows-Upgrade-All-Apps.ps1
# ---------------------------------------------------------------------------
# This PowerShell script upgrades all installed applications on Windows.
# It can be executed in a normal PowerShell session and will attempt to
# elevate itself if required.
#############################################################################>

[CmdletBinding()]
param(
    [switch]$NoWinget,
    [switch]$NoChocolatey,
    [switch]$NoScoop,
    [switch]$NoNpm,
    [switch]$NoPip,
    [switch]$NoDotNetTools,
    [switch]$NoWindowsUpdate,
    [switch]$NoStoreApps,
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Start-SelfElevated {
    [CmdletBinding()]
    param(
        [string[]]$Arguments
    )

    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
        $exe = if ($pwsh) { $pwsh.Source } else { 'powershell.exe' }

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $exe
        $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $Arguments"
        $psi.Verb = 'runas'
        [System.Diagnostics.Process]::Start($psi) | Out-Null
        exit
    }
}

# ... rest of the script content restored from commit da7f4dfad648ae94cad87d0af7ad1e1d7b3e9d12 ...
