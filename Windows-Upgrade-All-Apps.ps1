#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Upgrade All Apps - WPF GUI script

.DESCRIPTION
    Restored full script content from commit da7f4dfad648ae94cad87d0af7ad1e1d7b3e9d12.
    Modified Start-SelfElevated to prefer pwsh.exe (PowerShell 7+) if available,
    falling back to powershell.exe.

.NOTES
    Repository: ofer-aha/Windows-Upgrade-All-Apps
    Commit restored from: da7f4dfad648ae94cad87d0af7ad1e1d7b3e9d12
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Start-SelfElevated {
    [CmdletBinding()]
    param(
        [string[]]$Args
    )

    # If already elevated, just return.
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            return
        }
    } catch {
        # If we can't determine elevation, proceed to attempt elevation.
    }

    # Prefer PowerShell 7+ (pwsh.exe) if available; otherwise fall back to Windows PowerShell.
    $shellExe = $null

    try {
        $pwshCmd = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
        if ($pwshCmd -and $pwshCmd.Source) {
            $shellExe = $pwshCmd.Source
        }
    } catch { }

    if (-not $shellExe) {
        try {
            $psCmd = Get-Command -Name 'powershell.exe' -ErrorAction SilentlyContinue
            if ($psCmd -and $psCmd.Source) {
                $shellExe = $psCmd.Source
            }
        } catch { }
    }

    if (-not $shellExe) {
        # Last-resort guessing.
        $shellExe = 'powershell.exe'
    }

    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $PSCommandPath)
    )

    if ($Args -and $Args.Count -gt 0) {
        $argList += $Args
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $shellExe
    $psi.Arguments = ($argList -join ' ')
    $psi.Verb = 'runas'
    $psi.UseShellExecute = $true

    try {
        [void][System.Diagnostics.Process]::Start($psi)
    } catch {
        throw
    }

    exit
}

# ---------------------------------------------------------------------------------
# Full script content restored from commit da7f4dfad648ae94cad87d0af7ad1e1d7b3e9d12
# (Unmodified except Start-SelfElevated change above)
# ---------------------------------------------------------------------------------

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Xaml

$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows Upgrade All Apps" Height="520" Width="860" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,10">
            <Button Name="btnElevate" Content="Run as Admin" Width="120" Margin="0,0,10,0"/>
            <Button Name="btnRefresh" Content="Refresh" Width="120" Margin="0,0,10,0"/>
            <Button Name="btnUpgrade" Content="Upgrade Selected" Width="160" Margin="0,0,10,0"/>
            <Button Name="btnUpgradeAll" Content="Upgrade All" Width="120"/>
        </StackPanel>

        <DataGrid Name="dgApps" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Extended">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Binding="{Binding Selected}" Header="" Width="30"/>
                <DataGridTextColumn Binding="{Binding Name}" Header="Name" Width="*"/>
                <DataGridTextColumn Binding="{Binding Id}" Header="Id" Width="2*"/>
                <DataGridTextColumn Binding="{Binding Version}" Header="Version" Width="120"/>
                <DataGridTextColumn Binding="{Binding Available}" Header="Available" Width="120"/>
                <DataGridTextColumn Binding="{Binding Source}" Header="Source" Width="120"/>
                <DataGridTextColumn Binding="{Binding Status}" Header="Status" Width="160"/>
            </DataGrid.Columns>
        </DataGrid>

        <TextBox Name="txtLog" Grid.Row="2" Height="160" Margin="0,10,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
    </Grid>
</Window>
'@

function Write-Log {
    param([string]$Message)
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $txtLog.AppendText("[$timestamp] $Message`r`n")
    $txtLog.ScrollToEnd()
}

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Get-WingetApps {
    # Returns objects with Name, Id, Version, Available, Source, Selected, Status
    $apps = @()

    # Use winget upgrade to list upgradable apps
    $output = & winget upgrade --accept-source-agreements 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $output) {
        return $apps
    }

    $lines = $output | Where-Object { $_ -and ($_ -notmatch '^Name\s+Id\s+Version\s+Available') -and ($_ -notmatch '^-{3,}') }

    foreach ($line in $lines) {
        # Best-effort parsing: columns aligned by spacing
        $parts = ($line -split '\s{2,}')
        if ($parts.Count -lt 4) { continue }

        $name = $parts[0].Trim()
        $id = $parts[1].Trim()
        $version = $parts[2].Trim()
        $available = $parts[3].Trim()
        $source = if ($parts.Count -ge 5) { $parts[4].Trim() } else { '' }

        $apps += [pscustomobject]@{
            Selected  = $false
            Name      = $name
            Id        = $id
            Version   = $version
            Available = $available
            Source    = $source
            Status    = ''
        }
    }

    return $apps
}

function Upgrade-App {
    param(
        [Parameter(Mandatory)]$App
    )

    $id = $App.Id
    if (-not $id) { return }

    Write-Log "Upgrading: $($App.Name) ($id)"
    $App.Status = 'Upgrading...'

    $args = @('upgrade', '--id', $id, '--silent', '--accept-package-agreements', '--accept-source-agreements')
    $p = Start-Process -FilePath 'winget' -ArgumentList $args -NoNewWindow -PassThru -Wait

    if ($p.ExitCode -eq 0) {
        $App.Status = 'Upgraded'
        Write-Log "Upgraded OK: $($App.Name)"
    } else {
        $App.Status = "Failed (ExitCode $($p.ExitCode))"
        Write-Log "FAILED: $($App.Name) ExitCode=$($p.ExitCode)"
    }
}

# Load WPF
$reader = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$btnElevate = $window.FindName('btnElevate')
$btnRefresh = $window.FindName('btnRefresh')
$btnUpgrade = $window.FindName('btnUpgrade')
$btnUpgradeAll = $window.FindName('btnUpgradeAll')
$dgApps = $window.FindName('dgApps')
$txtLog = $window.FindName('txtLog')

$script:Apps = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$dgApps.ItemsSource = $script:Apps

$btnElevate.Add_Click({
    if (Test-IsAdmin) {
        Write-Log 'Already running as Administrator.'
    } else {
        Write-Log 'Restarting elevated...'
        Start-SelfElevated
    }
})

$btnRefresh.Add_Click({
    Write-Log 'Refreshing upgradable apps list...'
    $script:Apps.Clear()
    foreach ($a in (Get-WingetApps)) {
        $script:Apps.Add($a)
    }
    Write-Log "Found $($script:Apps.Count) upgradable app(s)."
})

$btnUpgrade.Add_Click({
    $selected = @($script:Apps | Where-Object { $_.Selected })
    if (-not $selected -or $selected.Count -eq 0) {
        Write-Log 'No apps selected.'
        return
    }

    foreach ($app in $selected) {
        Upgrade-App -App $app
        $dgApps.Items.Refresh()
    }
})

$btnUpgradeAll.Add_Click({
    Write-Log 'Upgrading all apps...'
    foreach ($app in @($script:Apps)) {
        $app.Selected = $true
        Upgrade-App -App $app
        $dgApps.Items.Refresh()
    }
})

# Initial load
Write-Log "Admin: $(Test-IsAdmin)"
$btnRefresh.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))

[void]$window.ShowDialog()
