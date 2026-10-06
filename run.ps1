<#
  Nature Index -> PNU Scholar collection agent  (one-command runner)

  Pipeline:
    1) agent-browser opens https://www.nature.com/nature-index/faq (a real Chrome,
       which clears the idp.nature.com auth-cookie gate) and dumps the page text,
       from which Q12's journal list is parsed.
    2) pipeline.py collects each journal's ISSN/eISSN (Crossref + OpenAlex when
       available), searches scholar.pusan.ac.kr by ISSN/name, verifies the match,
       reads whether "Nature Index" is among the journal's Registered DBs, and
       records the journal_id + detail URL.
    3) Results are written to .xlsx in a dated output folder.

  Usage:
    powershell -ExecutionPolicy Bypass -File run.ps1
    powershell -ExecutionPolicy Bypass -File run.ps1 -Date 20260726        # custom date/folder
    powershell -ExecutionPolicy Bypass -File run.ps1 -Root "D:\...\natureindex"
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),   # default: parent of this agent folder
    [string]$Date = (Get-Date -Format "yyyyMMdd"),
    [switch]$FreshCache                                    # ignore/rebuild the HTTP cache
)
$ErrorActionPreference = "Stop"
$AgentDir = $PSScriptRoot
$OutDir   = Join-Path $Root $Date
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
Write-Host "== Nature Index agent ==  output -> $OutDir" -ForegroundColor Cyan

# --- locate agent-browser (installed under Node 20 via nvm; exe path is version-pinned) ---
function Resolve-AgentBrowser {
    $c = Get-Command agent-browser -ErrorAction SilentlyContinue
    if ($c) { return "agent-browser" }
    $cands = @(
        "$env:LOCALAPPDATA\nvm\v20.19.6\node_modules\agent-browser\bin\agent-browser-win32-x64.exe",
        "C:\nvm4w\nodejs\node_modules\agent-browser\bin\agent-browser-win32-x64.exe"
    )
    foreach ($p in $cands) { if (Test-Path $p) { return $p } }
    throw "agent-browser not found. Activate Node 20 (nvm use 20.19.6) or run: npm install -g agent-browser; agent-browser install"
}
$AB = Resolve-AgentBrowser

# --- 1) fetch Nature Index FAQ (Q12 journal list) ---
$Faq = Join-Path $OutDir "faq_raw.txt"
Write-Host "[1/2] Fetching Nature Index FAQ via agent-browser ..." -ForegroundColor Yellow
& $AB batch "open https://www.nature.com/nature-index/faq" "wait 2500" "get text body" 2>&1 |
    Out-File -FilePath $Faq -Encoding utf8
if (-not (Select-String -Path $Faq -Pattern "articles\)" -Quiet)) {
    throw "FAQ fetch failed (no journal list found in page text). See $Faq"
}
$njournals = (Select-String -Path $Faq -Pattern "\(\d+ articles?\)").Count
Write-Host ("      captured FAQ page, ~{0} journal lines" -f $njournals)

# --- 2) collect ISSN/eISSN + match in PNU Scholar + write Excel ---
Write-Host "[2/2] Collecting ISSN/eISSN and matching in PNU Scholar ..." -ForegroundColor Yellow
$env:PYTHONIOENCODING = "utf-8"
$cache = Join-Path $OutDir "_http_cache.json"
if ($FreshCache -and (Test-Path $cache)) { Remove-Item $cache -Force }
python (Join-Path $AgentDir "pipeline.py") --faq $Faq --out $OutDir --date $Date --cache $cache
if ($LASTEXITCODE -ne 0) { throw "pipeline.py failed (exit $LASTEXITCODE)" }

Write-Host "== Done ==  files in $OutDir :" -ForegroundColor Cyan
Get-ChildItem $OutDir -Filter *.xlsx | Select-Object Name, Length | Format-Table -AutoSize
