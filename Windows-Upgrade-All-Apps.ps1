#Requires -RunAsAdministrator

$script:MyInvocation = $MyInvocation

function Start-SelfElevated {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    # Prefer PowerShell 7+ (pwsh.exe) if available, otherwise fall back to Windows PowerShell (powershell.exe)
    $pwsh = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    $exe = if ($null -ne $pwsh) { $pwsh.Source } else { 'powershell.exe' }

    # Preserve the original arguments; if none were provided, re-run with the current script's args
    $argsToPass = if ($PSBoundParameters.ContainsKey('Arguments')) { $Arguments } else { @($script:MyInvocation.UnboundArguments) }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    $psi.Verb = 'RunAs'

    # Ensure paths with spaces are quoted, and arguments are passed as a single string (Windows rules)
    $scriptPath = $script:MyInvocation.MyCommand.Path
    $quotedScriptPath = '"' + $scriptPath + '"'

    $baseArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $quotedScriptPath)
    $allArgs = @($baseArgs + $argsToPass)

    # Quote each argument that contains whitespace or quotes
    $psi.Arguments = ($allArgs | ForEach-Object {
        if ($_ -match '[\s\"]') {
            '"' + ($_ -replace '"','\\"') + '"'
        } else {
            $_
        }
    }) -join ' '

    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

# --- Existing script content below ---

# If we're not elevated, re-launch ourselves elevated
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-SelfElevated
}

# (rest of script continues...)
