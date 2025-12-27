#Requires -Version 5.1
<##############################################
# Windows-Upgrade-All-Apps.ps1
#
# WPF GUI runner for upgrading apps using winget.
# - Live output (stdout/stderr) streamed into the GUI.
# - List installed packages (winget list) and view/export.
# - Optional HTML fallback viewer embedded as PackagesViewer.html content.
# - Runs elevated when needed.
#
# Author: ofer-aha
#
##############################################>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------
# Embedded HTML fallback viewer
# ----------------------------
# NOTE: This is kept embedded as requested. The script can optionally write it out
# to a temp file and open it in the default browser.
$script:PackagesViewerHtml = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Packages Viewer</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 16px; color: #111; }
    h1 { font-size: 18px; margin: 0 0 10px 0; }
    .meta { margin: 8px 0 16px 0; color: #444; }
    textarea { width: 100%; height: 240px; font-family: Consolas, monospace; font-size: 12px; }
    table { border-collapse: collapse; width: 100%; margin-top: 12px; }
    th, td { border: 1px solid #ddd; padding: 6px 8px; font-size: 12px; }
    th { background: #f5f5f5; position: sticky; top: 0; z-index: 1; }
    .tools { display: flex; gap: 8px; flex-wrap: wrap; margin: 10px 0; }
    button { padding: 6px 10px; cursor: pointer; }
    input[type="search"] { padding: 6px 10px; width: min(520px, 100%); }
    .small { font-size: 12px; color: #555; }
    .hint { color: #666; font-size: 12px; margin-top: 6px; }
  </style>
</head>
<body>
  <h1>Packages Viewer (winget list output)</h1>
  <div class="meta">
    Paste the raw <code>winget list</code> output below, then click <b>Parse</b>.
  </div>

  <div class="tools">
    <input id="q" type="search" placeholder="Filter (name/id/version/source)" />
    <button onclick="parse()">Parse</button>
    <button onclick="downloadCsv()">Download CSV</button>
    <button onclick="clearAll()">Clear</button>
  </div>

  <textarea id="raw" spellcheck="false" placeholder="Paste winget list output here..."></textarea>
  <div class="hint">Tip: You can copy from the WPF app and paste here.</div>

  <div id="out" class="small"></div>
  <div style="max-height: 65vh; overflow: auto;">
    <table id="tbl"></table>
  </div>

<script>
let rows = [];

function clearAll(){
  document.getElementById('raw').value='';
  document.getElementById('tbl').innerHTML='';
  document.getElementById('out').textContent='';
  rows=[];
}

function parseWingetList(text){
  const lines = text.replace(/\r/g,'').split('\n').filter(l=>l.trim().length>0);
  if(lines.length < 2) return [];

  // Find header line (usually contains: Name  Id  Version  Available  Source)
  let headerIdx = 0;
  for(let i=0;i<Math.min(5, lines.length); i++){
    if(/\bName\b/.test(lines[i]) && /\bId\b/.test(lines[i]) && /\bVersion\b/.test(lines[i])) { headerIdx = i; break; }
  }

  // The line after header is a separator row (----)
  const hdr = lines[headerIdx];
  // Determine column starts by regex of multiple spaces
  // We'll split by 2+ spaces.
  const headerCols = hdr.trim().split(/\s{2,}/).map(s=>s.trim());

  const dataLines = lines.slice(headerIdx+2); // skip header + separator
  const parsed = [];
  for(const line of dataLines){
    const cols = line.trim().split(/\s{2,}/);
    // Map columns by position (best effort)
    const obj = {};
    for(let i=0;i<headerCols.length;i++) obj[headerCols[i]] = (cols[i] ?? '').trim();
    parsed.push(obj);
  }
  return parsed;
}

function parse(){
  const raw = document.getElementById('raw').value;
  rows = parseWingetList(raw);
  render();
}

function render(){
  const q = document.getElementById('q').value.toLowerCase();
  const filtered = rows.filter(r => {
    const s = Object.values(r).join(' ').toLowerCase();
    return s.includes(q);
  });

  document.getElementById('out').textContent = `${filtered.length} row(s)`;

  const tbl = document.getElementById('tbl');
  tbl.innerHTML='';
  if(filtered.length===0) return;

  const cols = Object.keys(filtered[0]);
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  cols.forEach(c=>{ const th=document.createElement('th'); th.textContent=c; trh.appendChild(th); });
  thead.appendChild(trh);
  tbl.appendChild(thead);

  const tbody = document.createElement('tbody');
  filtered.forEach(r=>{
    const tr = document.createElement('tr');
    cols.forEach(c=>{ const td=document.createElement('td'); td.textContent = r[c] ?? ''; tr.appendChild(td); });
    tbody.appendChild(tr);
  });
  tbl.appendChild(tbody);
}

document.getElementById('q').addEventListener('input', render);

function downloadCsv(){
  if(!rows.length){ alert('Nothing to export.'); return; }
  const cols = Object.keys(rows[0]);
  const esc = (v) => '"' + String(v ?? '').replace(/"/g,'""') + '"';
  const csv = [cols.map(esc).join(',')].concat(rows.map(r => cols.map(c=>esc(r[c])).join(','))).join('\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href=url;
  a.download='winget-list.csv';
  a.click();
  URL.revokeObjectURL(url);
}
</script>
</body>
</html>
'@

# ----------------------------
# Helpers
# ----------------------------
function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-SelfElevated {
  param(
    [string[]]$PassThruArgs = @()
  )
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = 'powershell.exe'
  $psi.Arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File', ('"{0}"' -f $PSCommandPath)) + $PassThruArgs | ForEach-Object { $_ } |  
    ForEach-Object { $_ } |  
    ForEach-Object { $_ } |  
    ForEach-Object { $_ }
  $psi.Arguments = ($psi.Arguments -join ' ')
  $psi.Verb = 'runas'
  $psi.UseShellExecute = $true
  [Diagnostics.Process]::Start($psi) | Out-Null
  exit
}

function Get-WingetPath {
  # Prefer winget.exe on PATH
  $cmd = Get-Command winget -ErrorAction SilentlyContinue
  if($cmd -and $cmd.Source) { return $cmd.Source }

  # Try default WindowsApps install location (best-effort)
  $wa = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
  if(Test-Path $wa) { return $wa }

  return $null
}

function Invoke-WingetListRaw {
  param(
    [string]$WingetExe
  )
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $WingetExe
  $psi.Arguments = 'list'
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  $null = $p.Start()
  $stdout = $p.StandardOutput.ReadToEnd()
  $stderr = $p.StandardError.ReadToEnd()
  $p.WaitForExit()

  return [pscustomobject]@{
    ExitCode = $p.ExitCode
    StdOut   = $stdout
    StdErr   = $stderr
  }
}

function Write-PackagesViewerHtmlToTempAndOpen {
  $path = Join-Path $env:TEMP ('PackagesViewer_{0}.html' -f ([guid]::NewGuid().ToString('N')))
  [IO.File]::WriteAllText($path, $script:PackagesViewerHtml, [Text.Encoding]::UTF8)
  Start-Process $path
}

# ----------------------------
# WPF UI
# ----------------------------
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

[xml]$xaml = @'
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Windows Upgrade All Apps (winget)" Height="720" Width="1100"
  WindowStartupLocation="CenterScreen">

  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" LastChildFill="False">
      <TextBlock Text="Windows Upgrade All Apps" FontSize="18" FontWeight="SemiBold" DockPanel.Dock="Left" Margin="0,0,12,6" />
      <TextBlock x:Name="LblStatus" Text="Ready" VerticalAlignment="Bottom" Foreground="#444" />
    </DockPanel>

    <WrapPanel Grid.Row="1" Margin="0,0,0,10">
      <Button x:Name="BtnUpgrade" Content="Upgrade All" Width="140" Height="32" Margin="0,0,8,0" />
      <Button x:Name="BtnList" Content="List Packages" Width="140" Height="32" Margin="0,0,8,0" />
      <Button x:Name="BtnOpenHtml" Content="Open HTML Viewer" Width="160" Height="32" Margin="0,0,8,0" />
      <Button x:Name="BtnCopyOutput" Content="Copy Output" Width="140" Height="32" Margin="0,0,8,0" />
      <Button x:Name="BtnClear" Content="Clear" Width="100" Height="32" Margin="0,0,8,0" />
      <CheckBox x:Name="ChkIncludeUnknown" Content="Include unknown" Margin="10,0,0,0" VerticalAlignment="Center" />
      <CheckBox x:Name="ChkForce" Content="Force" Margin="10,0,0,0" VerticalAlignment="Center" />
      <CheckBox x:Name="ChkSilent" Content="Silent" Margin="10,0,0,0" VerticalAlignment="Center" IsChecked="True" />
      <CheckBox x:Name="ChkAccept" Content="Accept agreements" Margin="10,0,0,0" VerticalAlignment="Center" IsChecked="True" />
    </WrapPanel>

    <Grid Grid.Row="2">
      <Grid.RowDefinitions>
        <RowDefinition Height="*" />
        <RowDefinition Height="Auto" />
      </Grid.RowDefinitions>

      <TextBox x:Name="TxtOutput" Grid.Row="0"
               FontFamily="Consolas" FontSize="12"
               Background="#0f0f0f" Foreground="#e6e6e6"
               VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
               TextWrapping="NoWrap" IsReadOnly="True" />

      <ProgressBar x:Name="Prg" Grid.Row="1" Height="16" Margin="0,10,0,0" IsIndeterminate="True" Visibility="Collapsed" />
    </Grid>

    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="BtnExit" Content="Exit" Width="100" Height="32" />
    </StackPanel>

  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$BtnUpgrade     = $window.FindName('BtnUpgrade')
$BtnList        = $window.FindName('BtnList')
$BtnOpenHtml    = $window.FindName('BtnOpenHtml')
$BtnCopyOutput  = $window.FindName('BtnCopyOutput')
$BtnClear       = $window.FindName('BtnClear')
$BtnExit        = $window.FindName('BtnExit')
$TxtOutput      = $window.FindName('TxtOutput')
$LblStatus      = $window.FindName('LblStatus')
$Prg            = $window.FindName('Prg')
$ChkIncludeUnknown = $window.FindName('ChkIncludeUnknown')
$ChkForce          = $window.FindName('ChkForce')
$ChkSilent         = $window.FindName('ChkSilent')
$ChkAccept         = $window.FindName('ChkAccept')

$script:WingetExe = Get-WingetPath

function Ui-AppendLine([string]$line) {
  if($null -eq $line) { return }
  $window.Dispatcher.Invoke([action]{
    $TxtOutput.AppendText($line + [Environment]::NewLine)
    $TxtOutput.ScrollToEnd()
  }) | Out-Null
}

function Ui-SetStatus([string]$text) {
  $window.Dispatcher.Invoke([action]{ $LblStatus.Text = $text }) | Out-Null
}

function Ui-SetBusy([bool]$busy) {
  $window.Dispatcher.Invoke([action]{
    $Prg.Visibility = if($busy) { 'Visible' } else { 'Collapsed' }
    $BtnUpgrade.IsEnabled = -not $busy
    $BtnList.IsEnabled = -not $busy
    $BtnOpenHtml.IsEnabled = -not $busy
  }) | Out-Null
}

function Start-ProcessWithLiveOutput {
  param(
    [Parameter(Mandatory)] [string]$FilePath,
    [Parameter(Mandatory)] [string]$Arguments
  )

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $FilePath
  $psi.Arguments = $Arguments
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
  $psi.StandardErrorEncoding  = [Text.Encoding]::UTF8

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  $p.EnableRaisingEvents = $true

  $p.add_OutputDataReceived({ if($_.Data){ Ui-AppendLine $_.Data } })
  $p.add_ErrorDataReceived({ if($_.Data){ Ui-AppendLine ('[err] ' + $_.Data) } })

  Ui-AppendLine ("> {0} {1}" -f $FilePath, $Arguments)

  $null = $p.Start()
  $p.BeginOutputReadLine()
  $p.BeginErrorReadLine()
  $p.WaitForExit()

  return $p.ExitCode
}

if(-not $script:WingetExe) {
  Ui-AppendLine 'winget.exe was not found on PATH.'
  Ui-AppendLine 'Install App Installer from Microsoft Store or ensure winget is available.'
  Ui-SetStatus 'winget not found'
}
else {
  Ui-AppendLine ("Using winget: {0}" -f $script:WingetExe)
}

$BtnClear.Add_Click({
  $TxtOutput.Clear()
  Ui-SetStatus 'Cleared'
})

$BtnExit.Add_Click({
  $window.Close()
})

$BtnCopyOutput.Add_Click({
  [Windows.Clipboard]::SetText($TxtOutput.Text)
  Ui-SetStatus 'Output copied to clipboard'
})

$BtnOpenHtml.Add_Click({
  Write-PackagesViewerHtmlToTempAndOpen
  Ui-SetStatus 'Opened HTML viewer'
})

$BtnList.Add_Click({
  if(-not $script:WingetExe) { return }
  Ui-SetBusy $true
  Ui-SetStatus 'Listing packages...'

  Start-Job -ScriptBlock {
    param($WingetExe)
    & $WingetExe list 2>&1 | Out-String
  } -ArgumentList $script:WingetExe | ForEach-Object {
    $job = $_
    while($job.State -eq 'Running') { Start-Sleep -Milliseconds 150 }
    $text = Receive-Job $job -Keep
    Remove-Job $job -Force | Out-Null

    Ui-AppendLine '---- winget list ----'
    ($text -split "`r?`n") | ForEach-Object { Ui-AppendLine $_ }
    Ui-AppendLine '---------------------'
    Ui-SetStatus 'List complete'
    Ui-SetBusy $false
  }
})

$BtnUpgrade.Add_Click({
  if(-not $script:WingetExe) { return }

  # winget upgrade often requires elevation for machine-wide installs
  if(-not (Test-IsAdmin)) {
    $res = [System.Windows.MessageBox]::Show(
      'Upgrading all apps may require Administrator privileges. Relaunch as Administrator?',
      'Elevation required',
      [System.Windows.MessageBoxButton]::YesNo,
      [System.Windows.MessageBoxImage]::Question
    )
    if($res -eq [System.Windows.MessageBoxResult]::Yes) {
      Start-SelfElevated
    }
  }

  $args = @('upgrade','--all')
  if($ChkIncludeUnknown.IsChecked) { $args += '--include-unknown' }
  if($ChkForce.IsChecked)          { $args += '--force' }
  if($ChkSilent.IsChecked)         { $args += '--silent' }
  if($ChkAccept.IsChecked) {
    $args += '--accept-package-agreements'
    $args += '--accept-source-agreements'
  }

  Ui-SetBusy $true
  Ui-SetStatus 'Upgrading...'

  # run on a background runspace to keep UI responsive
  $ps = [PowerShell]::Create()
  $ps.Runspace = [RunspaceFactory]::CreateRunspace()
  $ps.Runspace.ApartmentState = 'STA'
  $ps.Runspace.ThreadOptions = 'ReuseThread'
  $ps.Runspace.Open()

  $null = $ps.AddScript({
    param($WingetExe, $ArgsLine)
    $exit = Start-ProcessWithLiveOutput -FilePath $WingetExe -Arguments $ArgsLine
    return $exit
  }).AddArgument($script:WingetExe).AddArgument(($args -join ' '))

  $async = $ps.BeginInvoke()

  Register-ObjectEvent -InputObject $async -EventName Completed -Action {
    try {
      $exitCode = $ps.EndInvoke($async)
      Ui-AppendLine ("Process exited with code: {0}" -f $exitCode)
      Ui-SetStatus ("Upgrade finished (exit {0})" -f $exitCode)
    } catch {
      Ui-AppendLine ('[err] ' + $_.Exception.Message)
      Ui-SetStatus 'Upgrade failed'
    } finally {
      Ui-SetBusy $false
      $ps.Runspace.Close()
      $ps.Dispose()
      Unregister-Event -SourceIdentifier $eventSubscriber.SourceIdentifier -ErrorAction SilentlyContinue
    }
  } | Out-Null
})

$window.ShowDialog() | Out-Null
