<#
.SYNOPSIS
Upgrade all apps with winget, safely and verbosely.

.DESCRIPTION
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

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutDir,

    [Parameter(Mandatory=$false)]
    [Alias('h')]
    [switch]$Help,

    [Parameter(Mandatory=$false)]
    [switch]$AggressiveCleanup,

    [Parameter(Mandatory=$false)]
    [switch]$RemoveWindowsOld
)

<#
    Elevation + Disclaimer
    - If not running as admin, relaunch self elevated via UAC and exit current process.
    - Show disclaimer from the elevated instance.
#>

$ErrorActionPreference = 'Stop'

# If just asking for help, show it and exit before any elevation.
function Show-UsageHelp {
        try {
                $sp = $MyInvocation.MyCommand.Path
                if ($sp) { Get-Help -Full $sp; return }
        } catch {}
@"
Usage: pwsh -File "Windows-Upgrade-All-Apps.ps1" [-OutDir <path>] [-Help]

Options:
    -OutDir <path>  Destination folder for logs and snapshots (default: script folder; fallback: %TEMP%).
    -Help, -h       Show this help and exit.

Examples:
    pwsh -File ".\Windows-Upgrade-All-Apps.ps1"
    pwsh -File ".\Windows-Upgrade-All-Apps.ps1" -OutDir "D:\UpgradeLogs"
"@
}

if ($Help) { Show-UsageHelp; exit 0 }

# Detect admin
function Test-IsAdmin {
    try {
        $current = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($current)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

if (-not (Test-IsAdmin)) {
    Write-Host "Requesting administrator privileges..."
    $psExe = (Get-Process -Id $PID).Path
    # Build argument list to re-run this script with same parameters
    $escapedArgs = ($MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
        $k = $_.Key; $v = $_.Value
        if ($null -eq $v) { "-$k" }
        elseif ($v -is [switch]) { if ($v) { "-$k" } }
        else { "-$k `"$v`"" }
    }) -join ' '
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not $scriptPath) {
        # If running interactively, prompt the user
        Write-Error "Cannot auto-elevate: script path unknown. Save the script and run again."
        exit 1
    }
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" $escapedArgs"
    try {
        Start-Process -FilePath $psExe -ArgumentList $argList -Verb RunAs
        exit 0
    } catch {
        Write-Error "Elevation canceled or failed: $($_.Exception.Message)"
        exit 1
    }
}

# Please read and acknowledge before using the software.

# Elevated session banner
try {
    $who = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
} catch { $who = $env:USERNAME }
$isAdmin = Test-IsAdmin
$arch = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
$psver = $PSVersionTable.PSVersion.ToString()
$scriptStart = Get-Date
Write-Host "[elevated] User: $who | Admin: $isAdmin | PS: $psver | Arch: $arch | PID: $PID | Start: $($scriptStart.ToString('yyyy-MM-dd HH:mm:ss'))"

# Begin transcript logging to script folder (fallback to %TEMP%)
$logStamp = $scriptStart.ToString('yyyyMMdd_HHmmss')
# Determine preferred output directory
if ($OutDir) {
    $outDir = $OutDir
} else {
    $outDir = $PSScriptRoot
    if (-not $outDir) { $outDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    if (-not $outDir) { $outDir = $PWD.Path }
}

# Validate write access; fallback to TEMP if needed
try {
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $testFile = Join-Path $outDir ".write_test_${logStamp}.tmp"
    Set-Content -LiteralPath $testFile -Value "test" -Encoding ASCII -ErrorAction Stop
    Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
} catch {
    Write-Warning "Output folder not writable: $outDir. Falling back to TEMP."
    $outDir = $env:TEMP
}

Write-Host "[outdir] $outDir"

$logPath = Join-Path $outDir "Windows-Upgrade-All-Apps_${logStamp}.log"
try { Start-Transcript -Path $logPath -Append -ErrorAction Stop | Out-Null } catch { Write-Warning "Could not start transcript: $($_.Exception.Message)" }

# Helper: build file:// URI for local file paths
function Get-FileUri {
    param([Parameter(Mandatory=$true)][string]$Path)
    try { return ([Uri]$Path).AbsoluteUri } catch { return $Path }
}

# Copy or create PackagesViewer.html into output folder for easy access
$viewerDst = $null
try {
    $scriptRoot = $PSScriptRoot
    if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
    if (-not $scriptRoot) { $scriptRoot = $PWD.Path }
    $viewerSrc = Join-Path $scriptRoot 'Quick Start\Tools\PackagesViewer.html'
    $viewerDst = Join-Path $outDir 'PackagesViewer.html'
    
    if (Test-Path -LiteralPath $viewerSrc) {
        Copy-Item -LiteralPath $viewerSrc -Destination $viewerDst -Force -ErrorAction Stop
        Write-Host "[viewer] Copied PackagesViewer.html to $viewerDst"
    } else {
        # Create embedded viewer if source not found
        Write-Host "[viewer] PackagesViewer.html not found, creating embedded version..."
        $viewerHtml = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Windows Upgrade Packages Viewer</title>
  <style>
    :root { --bg:#0f172a; --muted:#334155; --fg:#e2e8f0; --accent:#22d3ee; --ok:#22c55e; --warn:#f59e0b; --bad:#ef4444; }
    html,body { margin:0; padding:0; background:var(--bg); color:var(--fg); font:14px/1.45 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
    header { padding:16px 20px; border-bottom:1px solid var(--muted); display:flex; gap:12px; align-items:center; flex-wrap:wrap; }
    h1 { font-size:18px; margin:0 14px 0 0; }
    .card { margin:16px 20px; padding:16px; border:1px solid var(--muted); border-radius:10px; background:#111827; }
    .row { display:flex; gap:16px; flex-wrap:wrap; }
    .col { flex:1 1 360px; }
    .files label { display:block; margin:8px 0 6px; color:#cbd5e1; }
    input[type="file"] { background:#0b1220; border:1px solid var(--muted); color:#93c5fd; padding:8px; border-radius:8px; width:100%; }
    .controls { display:flex; gap:12px; align-items:center; flex-wrap:wrap; margin-top:8px; }
    .btn { background:var(--muted); color:var(--fg); border:0; padding:8px 12px; border-radius:8px; cursor:pointer; }
    .btn.primary { background:var(--accent); color:#001015; font-weight:600; }
    .btn:disabled { opacity:0.5; cursor:not-allowed; }
    .search { flex:1 1 240px; }
    .search input { width:100%; padding:8px 10px; border:1px solid var(--muted); border-radius:8px; background:#0b1220; color:var(--fg); }
    .summary { display:flex; gap:16px; flex-wrap:wrap; }
    .pill { padding:6px 10px; border-radius:999px; background:#0b1220; border:1px solid var(--muted); color:#cbd5e1; }
    table { width:100%; border-collapse:collapse; }
    th, td { padding:8px 10px; border-bottom:1px solid #1f2937; }
    th { text-align:left; color:#cbd5e1; cursor:pointer; position:relative; }
    th .dir { position:absolute; right:6px; top:50%; transform:translateY(-50%); opacity:0.6; }
    tr:hover { background:#0b1220; }
    .tag { padding:3px 8px; border-radius:999px; font-size:12px; border:1px solid var(--muted); }
    .tag.add { background:#052e1b; color:#86efac; border-color:#166534; }
    .tag.remove { background:#3a0a0a; color:#fecaca; border-color:#991b1b; }
    .tag.update { background:#052433; color:#a5f3fc; border-color:#155e75; }
    footer { padding:16px 20px; color:#94a3b8; border-top:1px solid var(--muted); }
    code { background:#0b1220; padding:2px 6px; border-radius:6px; border:1px solid var(--muted); }
    .hidden { display:none; }
  </style>
</head>
<body>
  <header>
    <h1>Windows Upgrade Packages Viewer</h1>
    <div class="pill">Load the JSON snapshots exported by your script (Before/After).</div>
  </header>
  <div class="card files">
    <div class="row">
      <div class="col">
        <label for="beforeFile">Before JSON (Windows-Upgrade-Packages-Before_*.json)</label>
        <input type="file" id="beforeFile" accept="application/json,.json" />
      </div>
      <div class="col">
        <label for="afterFile">After JSON (Windows-Upgrade-Packages_*.json)</label>
        <input type="file" id="afterFile" accept="application/json,.json" />
      </div>
    </div>
    <div class="controls">
      <button class="btn primary" id="parseBtn" disabled>Parse selected files</button>
      <div class="search"><input id="search" type="search" placeholder="Filter by name, id, version, source…" /></div>
      <span id="meta" class="pill"></span>
    </div>
  </div>
  <div class="card" id="diffCard" class="hidden">
    <h3 style="margin-top:0">Differences</h3>
    <div class="summary" id="diffSummary"></div>
    <div id="diffTables"></div>
  </div>
  <div class="card" id="afterCard" class="hidden">
    <h3 style="margin-top:0">After snapshot</h3>
    <div id="afterTable"></div>
  </div>
  <div class="card" id="beforeCard" class="hidden">
    <h3 style="margin-top:0">Before snapshot</h3>
    <div id="beforeTable"></div>
  </div>
  <footer>
    Tip: Open this file locally, then use the file pickers above to load your JSON snapshots.
  </footer>
  <script>
  (function(){
    const $ = (sel,el=document)=>el.querySelector(sel);
    const $$ = (sel,el=document)=>Array.from(el.querySelectorAll(sel));
    const beforeInput = $('#beforeFile');
    const afterInput = $('#afterFile');
    const parseBtn = $('#parseBtn');
    const searchBox = $('#search');
    const meta = $('#meta');
    let beforeData = null;
    let afterData = null;
    let currentFilter = '';
    function enableParseIfReady(){
      parseBtn.disabled = !(beforeInput.files.length || afterInput.files.length);
    }
    beforeInput.addEventListener('change', enableParseIfReady);
    afterInput.addEventListener('change', enableParseIfReady);
    parseBtn.addEventListener('click', async () => {
      beforeData = beforeInput.files.length ? await readJson(beforeInput.files[0]) : null;
      afterData = afterInput.files.length ? await readJson(afterInput.files[0]) : null;
      const beforePkgs = normalizePackages(beforeData);
      const afterPkgs = normalizePackages(afterData);
      $('#beforeCard').classList.toggle('hidden', !beforePkgs.length);
      $('#afterCard').classList.toggle('hidden', !afterPkgs.length);
      $('#diffCard').classList.toggle('hidden', !(beforePkgs.length && afterPkgs.length));
      if (beforePkgs.length) renderTable('#beforeTable', beforePkgs, {title:'Before'});
      if (afterPkgs.length) renderTable('#afterTable', afterPkgs, {title:'After'});
      if (beforePkgs.length && afterPkgs.length) {
        renderDiff(beforePkgs, afterPkgs);
        meta.textContent = `${beforePkgs.length} → ${afterPkgs.length} packages`;
      } else if (beforePkgs.length) {
        meta.textContent = `${beforePkgs.length} packages (before only)`;
      } else if (afterPkgs.length) {
        meta.textContent = `${afterPkgs.length} packages (after only)`;
      } else {
        meta.textContent = '';
      }
    });
    searchBox.addEventListener('input', () => {
      currentFilter = searchBox.value.trim().toLowerCase();
      $$('.data-table').forEach(applyFilter);
    });
    function readJson(file){
      return new Promise((resolve,reject)=>{
        const fr = new FileReader();
        fr.onerror = () => reject(fr.error);
        fr.onload = () => {
          try {
            const text = String(fr.result || '').trim();
            const json = JSON.parse(text);
            resolve(json);
          } catch(e){ reject(new Error('Invalid JSON: '+ e.message)); }
        };
        fr.readAsText(file);
      });
    }
    function normalizePackages(obj){
      if (!obj) return [];
      let raw = [];
      if (Array.isArray(obj)) raw = obj;
      else if (Array.isArray(obj.Packages)) raw = obj.Packages;
      else if (Array.isArray(obj.apps)) raw = obj.apps;
      const mapField = (pkg, a, b, c) => pkg?.[a] ?? pkg?.[b] ?? pkg?.[c] ?? '';
      return raw.map(pkg=>({
        id: mapField(pkg,'Id','PackageIdentifier','id'),
        name: mapField(pkg,'Name','PackageName','name'),
        version: String(mapField(pkg,'Version','version','installedVersion')||''),
        source: mapField(pkg,'Source','source','Repository') || inferSource(obj) || '',
      })).filter(p=>p.id);
    }
    function inferSource(root){
      try {
        const srcs = root?.Sources || root?.sources;
        if (Array.isArray(srcs) && srcs.length){
          const wing = srcs.find(s=>/winget/i.test(s.Name||s.name||''));
          if (wing) return 'winget';
          if (/store/i.test(JSON.stringify(srcs))) return 'msstore';
        }
      } catch {}
      return '';
    }
    function renderTable(targetSel, data, opts={}){
      const target = $(targetSel);
      target.innerHTML = '';
      const table = document.createElement('table');
      table.className = 'data-table';
      const thead = document.createElement('thead');
      const headers = [
        {key:'name', label:'Name'},
        {key:'id', label:'Id'},
        {key:'version', label:'Version'},
        {key:'source', label:'Source'},
      ];
      let sortKey = 'name';
      let sortDir = 1;
      const trh = document.createElement('tr');
      headers.forEach(h=>{
        const th = document.createElement('th');
        th.textContent = h.label;
        const dir = document.createElement('span');
        dir.className = 'dir';
        th.appendChild(dir);
        th.addEventListener('click', ()=>{
          if (sortKey===h.key) sortDir*=-1; else { sortKey=h.key; sortDir=1; }
          renderBody();
          updateDirs();
        });
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      table.appendChild(tbody);
      function updateDirs(){
        $$('.dir', table).forEach(el=>el.textContent='');
        const idx = headers.findIndex(h=>h.key===sortKey);
        const t = $$('.dir', table)[idx];
        if (t) t.textContent = sortDir>0 ? '▲' : '▼';
      }
      function renderBody(){
        const filtered = data.filter(row => filterRow(row));
        filtered.sort((a,b)=>{
          const av = (a[sortKey]||'').toString().toLowerCase();
          const bv = (b[sortKey]||'').toString().toLowerCase();
          if (sortKey==='version') return compareVersion(av,bv)*sortDir;
          if (av<bv) return -1*sortDir; if (av>bv) return 1*sortDir; return 0;
        });
        tbody.innerHTML='';
        for (const r of filtered){
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td>${esc(r.name)}</td>
            <td><code>${esc(r.id)}</code></td>
            <td>${esc(r.version)}</td>
            <td>${esc(r.source)}</td>
          `;
          tbody.appendChild(tr);
        }
      }
      target.appendChild(table);
      renderBody();
      updateDirs();
      applyFilter(table);
    }
    function renderDiff(before, after){
      const byId = (arr)=>{
        const m = new Map();
        for (const p of arr) m.set(p.id, p);
        return m;
      };
      const B = byId(before), A = byId(after);
      const added=[], removed=[], updated=[];
      for (const [id, pkg] of A){
        if (!B.has(id)) added.push(pkg);
      }
      for (const [id, pkg] of B){
        if (!A.has(id)) removed.push(pkg);
        else {
          const np = A.get(id);
          if (normalizeVersion(pkg.version) !== normalizeVersion(np.version)){
            updated.push({ id, name: np.name||pkg.name, from: pkg.version, to: np.version, source: np.source||pkg.source });
          }
        }
      }
      const summary = $('#diffSummary');
      summary.innerHTML = '';
      const mkpill = (label, cls)=>{ const s=document.createElement('span'); s.className='pill'; s.style.borderColor='var(--muted)'; s.innerHTML = `<span class="tag ${cls}">${cls}</span> ${label}`; return s; };
      if (updated.length) summary.appendChild(mkpill(`${updated.length} updated`, 'update'));
      if (added.length) summary.appendChild(mkpill(`${added.length} added`, 'add'));
      if (removed.length) summary.appendChild(mkpill(`${removed.length} removed`, 'remove'));
      if (!summary.childNodes.length) summary.textContent = 'No differences detected.';
      const host = $('#diffTables');
      host.innerHTML = '';
      if (updated.length){ host.appendChild(buildTable('Updated packages', updated, ['name','id','from','to','source'], {
        decorators:{ from:(v)=>`<span class="tag remove">${esc(v)}</span>`, to:(v)=>`<span class="tag add">${esc(v)}</span>` }
      })); }
      if (added.length){ host.appendChild(buildTable('Added packages', added, ['name','id','version','source'])); }
      if (removed.length){ host.appendChild(buildTable('Removed packages', removed, ['name','id','version','source'])); }
    }
    function buildTable(title, rows, cols, opts={}){
      const card = document.createElement('div');
      card.style.marginTop = '12px';
      const h = document.createElement('h4');
      h.textContent = title;
      h.style.margin = '6px 0 10px';
      card.appendChild(h);
      const container = document.createElement('div');
      card.appendChild(container);
      const normalized = rows.map(r=>({
        name: r.name ?? r.Name ?? '',
        id: r.id ?? r.Id ?? r.PackageIdentifier ?? '',
        version: r.version ?? r.Version ?? '',
        source: r.source ?? r.Source ?? '',
        from: r.from ?? '',
        to: r.to ?? '',
      }));
      renderTable(container, normalized, cols, opts);
      return card;
    }
    function renderTable(container, rows, cols, opts={}){
      if (typeof container === 'string') container = $(container);
      container.innerHTML = '';
      const table = document.createElement('table'); table.className='data-table';
      const thead = document.createElement('thead'); const trh = document.createElement('tr');
      let sortKey = cols[0]; let sortDir = 1; const decorators = opts.decorators||{};
      for (const c of cols){ const th=document.createElement('th'); th.textContent = titleCase(c); const dir=document.createElement('span'); dir.className='dir'; th.appendChild(dir); th.addEventListener('click',()=>{ if(sortKey===c) sortDir*=-1; else {sortKey=c;sortDir=1;} renderBody(); updateDir(); }); trh.appendChild(th); }
      thead.appendChild(trh); table.appendChild(thead);
      const tbody=document.createElement('tbody'); table.appendChild(tbody); container.appendChild(table);
      function updateDir(){ $$('.dir',table).forEach(e=>e.textContent=''); const idx=cols.indexOf(sortKey); const node=$$('.dir',table)[idx]; if(node) node.textContent = sortDir>0?'▲':'▼'; }
      function renderBody(){ const filtered = rows.filter(r=>filterRow(r)); filtered.sort((a,b)=>{ const av=(a[sortKey]||'').toString().toLowerCase(); const bv=(b[sortKey]||'').toString().toLowerCase(); if (sortKey==='version'||sortKey==='from'||sortKey==='to') return compareVersion(av,bv)*sortDir; if(av<bv) return -1*sortDir; if(av>bv) return 1*sortDir; return 0;}); tbody.innerHTML=''; for(const r of filtered){ const tr=document.createElement('tr'); const cells = cols.map(c=>{ const v = r[c] ?? ''; const deco = decorators[c]; const out = deco ? deco(v) : esc(v); return `<td>${out}</td>`; }).join(''); tr.innerHTML = cells; tbody.appendChild(tr);} }
      renderBody(); updateDir(); applyFilter(table);
    }
    function titleCase(s){ return String(s||'').replace(/(^|_|-|\s)([a-z])/g,(m,sep,ch)=> (sep? ' ':'')+ch.toUpperCase()); }
    function esc(s){ return String(s ?? '').replace(/[&<>\"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\\':'\\'})[c]); }
    function compareVersion(a,b){
      const pa = normalizeVersion(a).split(/[^0-9a-zA-Z]+/).filter(Boolean);
      const pb = normalizeVersion(b).split(/[^0-9a-zA-Z]+/).filter(Boolean);
      const len = Math.max(pa.length, pb.length);
      for (let i=0;i<len;i++){
        const av = pa[i] ?? '0'; const bv = pb[i] ?? '0';
        const an = Number(av), bn = Number(bv);
        const aNum = !Number.isNaN(an) && /^\d+$/.test(av);
        const bNum = !Number.isNaN(bn) && /^\d+$/.test(bv);
        if (aNum && bNum){ if (an !== bn) return an < bn ? -1 : 1; }
        else { if (av !== bv) return av < bv ? -1 : 1; }
      }
      return 0;
    }
    function normalizeVersion(v){ return String(v||'').trim().replace(/^v/i,''); }
    function filterRow(row){
      if (!currentFilter) return true;
      const s = (row.name+' '+row.id+' '+(row.version||'')+' '+(row.source||'')).toLowerCase();
      return s.includes(currentFilter);
    }
    function applyFilter(table){
      const rows = $$('tbody tr', table);
      let shown = 0;
      for (const tr of rows){
        const t = tr.textContent.toLowerCase();
        const keep = !currentFilter || t.includes(currentFilter);
        tr.style.display = keep ? '' : 'none';
        if (keep) shown++;
      }
    }
  })();
  </script>
</body>
</html>
'@
        Set-Content -LiteralPath $viewerDst -Value $viewerHtml -Encoding UTF8 -ErrorAction Stop
        Write-Host "[viewer] Created PackagesViewer.html at $viewerDst"
    }
} catch {
    Write-Warning "Could not copy or create PackagesViewer.html: $($_.Exception.Message)"
}

# Planned package export path (post-upgrade snapshot)
$pkgExportPath = Join-Path $outDir "Windows-Upgrade-Packages_${logStamp}.json"
# Pre-upgrade snapshot path
$pkgPreExportPath = Join-Path $outDir "Windows-Upgrade-Packages-Before_${logStamp}.json"

# System summary (logged after transcript starts)
try {
    $comp = $env:COMPUTERNAME
    $domain = $env:USERDOMAIN
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    $caption = $os.Caption
    $osVer = $os.Version
    $build = $os.BuildNumber
    $osArch = $os.OSArchitecture
    $lastBoot = $os.LastBootUpTime
    $uptime = (Get-Date) - $lastBoot
    $totalMemGB = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    Write-Host "[host] $comp.$domain | OS: $caption $osVer (Build $build, $osArch) | Uptime: $($uptime.ToString('dd\.hh\:mm\:ss')) | RAM: ${totalMemGB}GB"
    # Extra: CPU, Disk, Locale
    $cpu = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1 Name, NumberOfCores, NumberOfLogicalProcessors
    $cpuName = $cpu.Name
    $cores = $cpu.NumberOfCores
    $threads = $cpu.NumberOfLogicalProcessors
    Write-Host "[cpu] $cpuName | Cores: $cores | Threads: $threads"

    $sysDrive = $env:SystemDrive
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$sysDrive'"
    if ($disk) {
        $freeGB = [Math]::Round($disk.FreeSpace / 1GB, 2)
        $sizeGB = [Math]::Round($disk.Size / 1GB, 2)
        Write-Host "[disk] $sysDrive Free: ${freeGB}GB / ${sizeGB}GB"
    }

    $culture = [System.Globalization.CultureInfo]::CurrentCulture.Name
    $tz = (Get-TimeZone).Id
    Write-Host "[locale] $culture | TimeZone: $tz"
} catch {
    Write-Warning "System summary failed: $($_.Exception.Message)"
}

Write-Host "--------------------------------------------"
Write-Host "SOFTWARE DISCLAIMER: NO IMPLIED WARRANTY"
Write-Host "--------------------------------------------"

Write-Host "By using this software, you acknowledge and agree to the following terms:"
Write-Host ""

Write-Host "1. No Warranty: This software is provided 'as-is' and without any express or implied warranties, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose."
Write-Host ""

Write-Host "2. Use at Your Own Risk: The use of this software is at your own risk. The author(s) and contributors shall not be liable for any direct, indirect, incidental, special, exemplary, or consequential damages (including, but not limited to, procurement of substitute goods or services; loss of use, data, or profits; or business interruption) however caused and on any theory of liability, whether in contract, strict liability, or tort (including negligence or otherwise) arising in any way out of the use of this software, even if advised of the possibility of such damage."
Write-Host ""

Write-Host "3. No Support: The author(s) of this software may not provide support, maintenance, updates, or enhancements for this software."
Write-Host ""

Write-Host "4. Compliance: It is your responsibility to ensure that your use of this software complies with all applicable laws and regulations."
Write-Host ""

Write-Host "If you do not agree with these terms, please do not use the software."

Write-Host "--------------------------------------------"

# Resolve winget path robustly (fixes prior lookup bug)
function Get-WingetPath {
    # Prefer PATH resolution
    $cmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Path) {
        return $cmd.Path
    }

    # Fallback to Local WindowsApps shim
    $localShim = Join-Path $Env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    if (Test-Path $localShim) {
        return $localShim
    }

    throw "winget.exe not found. Ensure 'App Installer' is installed from Microsoft Store."
}

function Set-WingetRecommendedSources {
    [CmdletBinding()]
    param([string]$WingetPath)
    try {
        Write-Host "[winget] Checking sources…"
        $sourcesOut = & "$WingetPath" source list 2>$null
        $hasWinget = $false
        $hasMsstore = $false
        if ($sourcesOut) {
            $hasWinget = ($sourcesOut -match "\bwinget\b")
            $hasMsstore = ($sourcesOut -match "\bmsstore\b")
        }
        if (-not $hasWinget -or -not $hasMsstore) {
            Write-Warning "[winget] Missing default sources. Resetting sources to defaults…"
            & "$WingetPath" source reset --force
            $sourcesOut = & "$WingetPath" source list 2>$null
            $hasWinget = ($sourcesOut -match "\bwinget\b")
            $hasMsstore = ($sourcesOut -match "\bmsstore\b")
        }
        if (-not $hasMsstore) {
            # Fallback explicit add for Microsoft Store if still missing
            Write-Host "[winget] Adding msstore source…"
            try {
                & "$WingetPath" source add --name msstore --arg "https://storeedgefd.dsx.mp.microsoft.com/v9.0" --type Microsoft.Rest | Out-Null
            } catch { Write-Warning "[winget] Could not add msstore: $($_.Exception.Message)" }
        }
        Write-Host "[winget] Sources ready."
    } catch {
        Write-Warning "[winget] Ensure sources failed: $($_.Exception.Message)"
    }
}

function Get-FreeSpaceBytes {
    param(
        [Parameter(Mandatory=$false)] [string]$DriveLetter = $env:SystemDrive
    )
    try {
    if ($DriveLetter.Length -eq 1) { $DriveLetter = ($DriveLetter + ':') }
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$DriveLetter'"
        if ($null -ne $disk) { return [int64]$disk.FreeSpace }
    } catch {}
    return [int64]0
}

function Convert-BytesToReadable {
    param([long]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes B" }
    elseif ($Bytes -lt 1MB) { return "{0:N2} KB" -f ($Bytes/1KB) }
    elseif ($Bytes -lt 1GB) { return "{0:N2} MB" -f ($Bytes/1MB) }
    else { return "{0:N2} GB" -f ($Bytes/1GB) }
}

function Remove-PathContentsSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$Path
    )
    try {
        if (Test-Path -LiteralPath $Path) {
            Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    } catch {}
}

function Remove-ItemSafe {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Path)
    try { if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue } } catch {}
}

function Restart-ServicesSafe {
    param([string[]]$Names)
    foreach ($n in $Names) {
        try { Start-Service -Name $n -ErrorAction SilentlyContinue } catch {}
    }
}

function Stop-ServicesSafe {
    param([string[]]$Names)
    foreach ($n in $Names) {
        try { Stop-Service -Name $n -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Invoke-ComponentStoreCleanup {
    param([switch]$Aggressive)
    try {
        $dism = "$env:WINDIR\System32\dism.exe"
        if (-not (Test-Path $dism)) { return }
        if ($Aggressive) {
            Write-Host "[cleanup] DISM StartComponentCleanup /ResetBase (aggressive)…"
            & $dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase | Write-Output
        } else {
            Write-Host "[cleanup] DISM StartComponentCleanup (safe)…"
            & $dism /Online /Cleanup-Image /StartComponentCleanup | Write-Output
        }
        Write-Host "[cleanup] DISM AnalyzeComponentStore (summary)…"
        & $dism /Online /Cleanup-Image /AnalyzeComponentStore | Write-Output
    } catch { Write-Warning "DISM cleanup error: $($_.Exception.Message)" }
}

function Invoke-PostUpgradeCleanup {
    [CmdletBinding()]
    param(
        [switch]$Aggressive,
        [switch]$PurgeWindowsOld
    )
    $sys = $env:SystemDrive
    $before = Get-FreeSpaceBytes -DriveLetter $sys
    Write-Host "[cleanup] Starting post-upgrade cleanup on $sys…"

    # 1) Temp folders (user + system)
    $paths = @()
    $paths += $env:TEMP, $env:TMP
    $paths += Join-Path $sys 'Windows\Temp'
    # All user profiles temp
    try {
        $userRoots = Get-ChildItem -LiteralPath (Join-Path $sys 'Users') -Directory -Force -ErrorAction SilentlyContinue
        foreach ($u in $userRoots) {
            $paths += Join-Path $u.FullName 'AppData\Local\Temp'
        }
    } catch {}
    $paths = $paths | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
    foreach ($p in $paths) { Write-Host "[cleanup] Temp -> $p"; Remove-PathContentsSafe -Path $p }

    # 2) Windows Error Reporting
    Remove-PathContentsSafe -Path (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportQueue')
    Remove-PathContentsSafe -Path (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportArchive')

    # 3) Memory dumps
    Remove-ItemSafe -Path (Join-Path $sys 'Windows\MEMORY.DMP')
    Remove-PathContentsSafe -Path (Join-Path $sys 'Windows\Minidump')

    # 4) Windows Update cache + Delivery Optimization
    $svc = @('wuauserv','bits','cryptSvc','msiserver','dosvc')
    Stop-ServicesSafe -Names $svc
    Remove-PathContentsSafe -Path (Join-Path $sys 'Windows\SoftwareDistribution\Download')
    Remove-PathContentsSafe -Path (Join-Path $sys 'Windows\SoftwareDistribution\DeliveryOptimization')
    # Additional DO cache location
    Remove-PathContentsSafe -Path (Join-Path $env:ProgramData 'Microsoft\Windows\DeliveryOptimization\Cache')
    Restart-ServicesSafe -Names $svc

    # 5) Component Store cleanup (DISM)
    Invoke-ComponentStoreCleanup -Aggressive:$Aggressive

    # 6) Recycle Bin (all drives)
    try { Get-PSDrive -PSProvider FileSystem | ForEach-Object { Clear-RecycleBin -DriveLetter $_.Name -Force -ErrorAction SilentlyContinue } } catch {}

    # 7) Optional: Previous Windows installation (Windows.old)
    if ($PurgeWindowsOld) {
        $winOld = Join-Path $sys 'Windows.old'
        if (Test-Path -LiteralPath $winOld) {
            Write-Warning "[cleanup] Removing Windows.old (cannot roll back OS after this)."
            try {
                # Attempt standard removal; may fail due to ACLs
                Remove-Item -LiteralPath $winOld -Recurse -Force -ErrorAction Stop
            } catch {
                # Fallback: take ownership and ACL, then remove
                try {
                    & "$env:WINDIR\System32\takeown.exe" /F "$winOld" /R /D Y | Out-Null
                    & "$env:WINDIR\System32\icacls.exe" "$winOld" /grant administrators:F /T /C | Out-Null
                    Remove-Item -LiteralPath $winOld -Recurse -Force -ErrorAction SilentlyContinue
                } catch {}
            }
        }
    }

    $after = Get-FreeSpaceBytes -DriveLetter $sys
    $reclaimed = [math]::Max([int64]0, ($after - $before))
    $readable = Convert-BytesToReadable -Bytes $reclaimed
    Write-Host "[cleanup] Completed. Space reclaimed: $readable"
    return [pscustomobject]@{ Before=$before; After=$after; Reclaimed=$reclaimed; ReclaimedReadable=$readable }
}

${status} = 'OK'
try {
    $Winget = Get-WingetPath
    Write-Host "Using winget at: $Winget"

    # WinGet version/info
    & "$Winget" --info

    # Ensure recommended/default sources (winget + msstore)
    Set-WingetRecommendedSources -WingetPath $Winget

    # Update sources
    & "$Winget" source update

    # Export pre-upgrade package snapshot
    try {
        & "$Winget" export --output "$pkgPreExportPath" --include-versions --include-unknown
    } catch {
        try { & "$Winget" export -o "$pkgPreExportPath" --include-versions --include-unknown } catch { Write-Warning "Pre-upgrade export failed: $($_.Exception.Message)" }
    }
    if (Test-Path $pkgPreExportPath) { Write-Host "[packages-before] Exported to $pkgPreExportPath" }

    # Upgrade all apps silently, accept agreements
    & "$Winget" upgrade --all --silent --accept-source-agreements --accept-package-agreements

    # Export package snapshot (after upgrade)
    try {
        & "$Winget" export --output "$pkgExportPath" --include-versions --include-unknown
    } catch {
        # Fallback for older winget using -o
        try {
            & "$Winget" export -o "$pkgExportPath" --include-versions --include-unknown
        } catch {
            Write-Warning "Package export failed: $($_.Exception.Message)"
        }
    }
    if (Test-Path $pkgExportPath) {
        Write-Host "[packages] Exported to $pkgExportPath"
    }
}
catch {
    Write-Error "Failed to run winget operations: $($_.Exception.Message)"
    ${status} = 'FAILED'
}
finally {
    # Post-upgrade cleanup and disk reclaim summary (logged inside transcript)
    try {
        $cleanupResult = Invoke-PostUpgradeCleanup -Aggressive:$AggressiveCleanup -PurgeWindowsOld:$RemoveWindowsOld
    } catch { Write-Warning "Cleanup encountered an error: $($_.Exception.Message)" }

    $end = Get-Date
    $duration = $end - $scriptStart
    $durStr = $duration.ToString('hh\:mm\:ss')
    if ($cleanupResult) {
        $rec = $cleanupResult.ReclaimedReadable
        Write-Host "[completed] Status: ${status} | End: $($end.ToString('yyyy-MM-dd HH:mm:ss')) | Duration: $durStr | Reclaimed: $rec"
    } else {
        Write-Host "[completed] Status: ${status} | End: $($end.ToString('yyyy-MM-dd HH:mm:ss')) | Duration: $durStr"
    }
    if ($logPath) {
        try { Stop-Transcript | Out-Null } catch { }
        Write-Host "[log] $logPath"
    }
    if ($viewerDst -and (Test-Path -LiteralPath $viewerDst)) {
        try {
            $viewerUri = Get-FileUri -Path $viewerDst
            Write-Host "[viewer] $viewerDst"
            Write-Host "[open] $viewerUri"
        } catch {}
    }
    if ($pkgPreExportPath -and (Test-Path $pkgPreExportPath)) { Write-Host "[packages-before] $pkgPreExportPath" }
    if ($pkgExportPath -and (Test-Path $pkgExportPath)) { Write-Host "[packages-after]  $pkgExportPath" }
}

if (${status} -ne 'OK') { exit 1 }
