#requires -Version 5.1
<#!
Windows-Upgrade-All-Apps.ps1

- Adds a small WinForms UI to list/upgrade applications via winget.
- This update adds:
  * Ui-SetBusy accepts an optional status string and toggles controls.
  * Live output + progress during `winget list` using Start-ProcessWithLiveOutput.
  * Upgrade handler writes a start line and guarantees busy state on/off.
  * Optional UI heartbeat while long-running operations execute.
!>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Helpers
# ----------------------------

function Ui-AppendOutput {
    param(
        [Parameter(Mandatory=$true)][string]$Text
    )
    if ($null -ne $txtOutput -and -not $txtOutput.IsDisposed) {
        $txtOutput.AppendText($Text)
        if (-not $Text.EndsWith("`r`n")) { $txtOutput.AppendText("`r`n") }
        $txtOutput.SelectionStart = $txtOutput.TextLength
        $txtOutput.ScrollToCaret()
    }
}

function Ui-SetStatus {
    param([string]$Text = "")
    if ($null -ne $lblStatus -and -not $lblStatus.IsDisposed) {
        $lblStatus.Text = $Text
    }
}

function Ui-SetBusy {
    param(
        [Parameter()][bool]$Busy,
        [Parameter()][string]$Status
    )

    if ($PSBoundParameters.ContainsKey('Status')) {
        Ui-SetStatus $Status
    }

    if ($null -ne $progressBar -and -not $progressBar.IsDisposed) {
        $progressBar.Style = if ($Busy) { 'Marquee' } else { 'Blocks' }
        $progressBar.MarqueeAnimationSpeed = if ($Busy) { 30 } else { 0 }
        $progressBar.Visible = $Busy
    }

    # Main action buttons
    if ($null -ne $btnList)   { $btnList.Enabled   = -not $Busy }
    if ($null -ne $btnUpgrade){ $btnUpgrade.Enabled= -not $Busy }

    # Output-related buttons disabled while busy
    if ($null -ne $btnCopyOutput) { $btnCopyOutput.Enabled = -not $Busy }
    if ($null -ne $btnClear)      { $btnClear.Enabled      = -not $Busy }

    # Allow exit only when not busy (prevents half-finished runs)
    if ($null -ne $btnExit)       { $btnExit.Enabled       = -not $Busy }
}

function Start-UiHeartbeat {
    # Optional: simple heartbeat so user sees UI is alive during long operations.
    if ($null -eq $script:HeartbeatTimer -or $script:HeartbeatTimer.IsDisposed) {
        $script:HeartbeatTimer = New-Object System.Windows.Forms.Timer
        $script:HeartbeatTimer.Interval = 3000
        $script:HeartbeatTimer.Add_Tick({
            if ($null -ne $lblStatus -and -not $lblStatus.IsDisposed) {
                # Don't overwrite a custom status, just add a subtle dot pulse.
                if ($lblStatus.Text -match '\.\.\.$') {
                    $lblStatus.Text = ($lblStatus.Text -replace '\.\.\.$','')
                } else {
                    $lblStatus.Text = ($lblStatus.Text + '.')
                }
            }
        })
    }
    $script:HeartbeatTimer.Start()
}

function Stop-UiHeartbeat {
    if ($null -ne $script:HeartbeatTimer -and -not $script:HeartbeatTimer.IsDisposed) {
        $script:HeartbeatTimer.Stop()
    }
}

function Start-ProcessWithLiveOutput {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter()][string]$Arguments = "",
        [Parameter()][string]$WorkingDirectory,
        [Parameter()][int]$PollMilliseconds = 100
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    if ($WorkingDirectory) { $psi.WorkingDirectory = $WorkingDirectory }
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    [void]$p.Start()

    while (-not $p.HasExited) {
        while (-not $p.StandardOutput.EndOfStream) {
            $line = $p.StandardOutput.ReadLine()
            if ($line -ne $null) {
                $form.BeginInvoke([Action]{ Ui-AppendOutput $line }) | Out-Null
            }
        }
        while (-not $p.StandardError.EndOfStream) {
            $line = $p.StandardError.ReadLine()
            if ($line -ne $null) {
                $form.BeginInvoke([Action]{ Ui-AppendOutput $line }) | Out-Null
            }
        }
        Start-Sleep -Milliseconds $PollMilliseconds
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Drain remaining output
    while (-not $p.StandardOutput.EndOfStream) {
        $line = $p.StandardOutput.ReadLine()
        if ($line -ne $null) {
            $form.BeginInvoke([Action]{ Ui-AppendOutput $line }) | Out-Null
        }
    }
    while (-not $p.StandardError.EndOfStream) {
        $line = $p.StandardError.ReadLine()
        if ($line -ne $null) {
            $form.BeginInvoke([Action]{ Ui-AppendOutput $line }) | Out-Null
        }
    }

    $p.WaitForExit()
    return $p.ExitCode
}

# ----------------------------
# UI
# ----------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Windows Upgrade All Apps'
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = 'CenterScreen'

$btnList = New-Object System.Windows.Forms.Button
$btnList.Text = 'List'
$btnList.Location = New-Object System.Drawing.Point(10, 10)
$btnList.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnList)

$btnUpgrade = New-Object System.Windows.Forms.Button
$btnUpgrade.Text = 'Upgrade'
$btnUpgrade.Location = New-Object System.Drawing.Point(110, 10)
$btnUpgrade.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnUpgrade)

$btnCopyOutput = New-Object System.Windows.Forms.Button
$btnCopyOutput.Text = 'Copy Output'
$btnCopyOutput.Location = New-Object System.Drawing.Point(210, 10)
$btnCopyOutput.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnCopyOutput)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(330, 10)
$btnClear.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClear)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = 'Exit'
$btnExit.Location = New-Object System.Drawing.Point(430, 10)
$btnExit.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnExit)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 50)
$progressBar.Size = New-Object System.Drawing.Size(860, 16)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 70)
$lblStatus.Size = New-Object System.Drawing.Size(860, 16)
$lblStatus.Text = ''
$form.Controls.Add($lblStatus)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(10, 95)
$txtOutput.Size = New-Object System.Drawing.Size(860, 455)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = 'Vertical'
$txtOutput.ReadOnly = $true
$txtOutput.Font = New-Object System.Drawing.Font('Consolas', 10)
$form.Controls.Add($txtOutput)

# ----------------------------
# Handlers
# ----------------------------

$btnCopyOutput.Add_Click({
    if ($txtOutput.TextLength -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($txtOutput.Text)
        Ui-SetStatus 'Output copied to clipboard.'
    }
})

$btnClear.Add_Click({
    $txtOutput.Clear()
    Ui-SetStatus ''
})

$btnExit.Add_Click({
    $form.Close()
})

$btnList.Add_Click({
    try {
        Ui-AppendOutput "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting: winget list"
        Ui-SetBusy -Busy $true -Status 'Listing apps...'
        Start-UiHeartbeat

        $exit = Start-ProcessWithLiveOutput -FilePath 'winget' -Arguments 'list'

        Ui-AppendOutput "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished: winget list (exit $exit)"
        Ui-SetStatus "List completed (exit $exit)."
    } catch {
        Ui-AppendOutput "ERROR: $($_.Exception.Message)"
        Ui-SetStatus 'List failed.'
    } finally {
        Stop-UiHeartbeat
        Ui-SetBusy -Busy $false
    }
})

$btnUpgrade.Add_Click({
    try {
        Ui-AppendOutput "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting: winget upgrade --all"
        Ui-SetBusy -Busy $true -Status 'Upgrading apps...'
        Start-UiHeartbeat

        $exit = Start-ProcessWithLiveOutput -FilePath 'winget' -Arguments 'upgrade --all'

        Ui-AppendOutput "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished: winget upgrade --all (exit $exit)"
        Ui-SetStatus "Upgrade completed (exit $exit)."
    } catch {
        Ui-AppendOutput "ERROR: $($_.Exception.Message)"
        Ui-SetStatus 'Upgrade failed.'
    } finally {
        Stop-UiHeartbeat
        Ui-SetBusy -Busy $false
    }
})

# Initial state
Ui-SetBusy -Busy $false -Status 'Ready.'

[void]$form.ShowDialog()
