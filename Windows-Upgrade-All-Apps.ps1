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

# ... (rest of original WPF script omitted for brevity)
