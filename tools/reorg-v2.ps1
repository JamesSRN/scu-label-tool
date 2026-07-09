<#
  reorg-v2.ps1 - Light-touch V2 folder cleanup for SCU Label Printing.

  RUN LOCALLY from the repo root (this is where git works and the files are real).
  It PREVIEWS by default (changes nothing). Add -Execute to actually apply.

  Build-critical files are NEVER moved: MedParser.bas, Build-Release.vbs,
  scu_emblem.png, and MedicationDispensing.xlsm stay at the repo root.

  What it does:
    1. Delete the redundant nested repo clone under _backups\ (frees ~10 MB; git-ignored).
    2. Delete stale Excel lock files (~$*).
    3. Stop tracking throwaway temp images/blobs in tools\ (git rm --cached; kept on disk).
    4. Move reference docs into docs\ and one-off notes into docs\archive\.
    5. Move logo-source art into assets\logo-source\.

  Preview:  powershell -ExecutionPolicy Bypass -File tools\reorg-v2.ps1
  Apply:    powershell -ExecutionPolicy Bypass -File tools\reorg-v2.ps1 -Execute
  Then review 'git status' and commit in GitHub Desktop.
#>
param([switch]$Execute)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Test-Path (Join-Path $root 'MedParser.bas'))) {
    Write-Host "This does not look like the repo root (MedParser.bas not found). Aborting." -ForegroundColor Red
    exit 1
}

$mode = if ($Execute) { "APPLY" } else { "PREVIEW (no changes)" }
Write-Host ""
Write-Host "SCU Label Printing - V2 folder reorg   [$mode]" -ForegroundColor Cyan
Write-Host "Root: $root"
Write-Host ""

# 1) Delete the redundant nested repo clone.
$nested = Join-Path $root '_backups\GIT_VERSION_SCU Label Printing'
if (Test-Path -LiteralPath $nested) {
    Write-Host "  delete nested clone: _backups\GIT_VERSION_SCU Label Printing\  (frees ~10 MB)"
    if ($Execute) { Remove-Item -Recurse -Force -LiteralPath $nested }
} else {
    Write-Host "  (nested clone already gone)"
}

# 2) Delete stale Excel lock files.
Get-ChildItem -Path $root -Filter '~$*' -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
    $lock = $_.FullName; $ln = $_.Name
    Write-Host "  delete stale lock: $ln"
    if ($Execute) { Remove-Item -Force -LiteralPath $lock }
}

# 3) Stop tracking throwaway temp images/blobs in tools\ (they stay on disk, just untracked).
$tracked = @()
$tracked += (git ls-files 'tools/_*')
$tracked += (git ls-files 'tools/*.generated.txt')
foreach ($f in ($tracked | Where-Object { $_ -ne '' } | Select-Object -Unique)) {
    Write-Host "  git untrack: $f"
    if ($Execute) { git rm --cached --ignore-unmatch -- "$f" | Out-Null }
}

# Helper: move a tracked file into a destination folder using git mv (keeps history).
function Move-Tracked([string]$src, [string]$dstDir) {
    if (-not (Test-Path -LiteralPath (Join-Path $root $src))) { return }
    $gitSrc = $src -replace '\\','/'
    $gitDst = ($dstDir -replace '\\','/') + '/'
    Write-Host "  move $src  ->  $dstDir\"
    if ($Execute) {
        New-Item -ItemType Directory -Force -Path (Join-Path $root $dstDir) | Out-Null
        git mv -k -- "$gitSrc" "$gitDst" | Out-Null
    }
}

# 4) Reference docs -> docs\ ; one-off notes -> docs\archive\.
Move-Tracked 'HANDOFF.md'            'docs'
Move-Tracked 'SETUP_INSTRUCTIONS.md' 'docs'
Move-Tracked 'LABEL_REDESIGN.md'     'docs'
Move-Tracked 'COMMIT_COMMANDS.md'        'docs\archive'
Move-Tracked 'GITHUB_ISSUE_RESPONSES.md' 'docs\archive'

# 5) Logo-source art -> assets\logo-source\.
Move-Tracked 'Black SCU Logo + Transparent Background.png'                'assets\logo-source'
Move-Tracked 'cropped_Black SCU Logo + Transparent Background - Copy.png' 'assets\logo-source'

Write-Host ""
if ($Execute) {
    Write-Host "Done. Review 'git status', then commit in GitHub Desktop." -ForegroundColor Green
    Write-Host "Left at root (build-critical): scu_emblem.png, MedParser.bas, Build-Release.vbs, MedicationDispensing.xlsm"
} else {
    Write-Host "Preview only. Re-run to apply:" -ForegroundColor Yellow
    Write-Host "  powershell -ExecutionPolicy Bypass -File tools\reorg-v2.ps1 -Execute" -ForegroundColor Yellow
}
Write-Host ""

# ---------------------------------------------------------------------------
# OPTIONAL (manual): track a clean, PHI-FREE template so a fresh clone can build.
# Do this ONLY after opening the file and confirming it has NO patient data:
#   git mv "Broken_PrettyPrint_MedicationDispensing.xlsm" "templates/Template_MedicationDispensing.xlsm"
#   # then uncomment the '!templates/Template_MedicationDispensing.xlsm' line in .gitignore
#   git add -f "templates/Template_MedicationDispensing.xlsm"
# Build-Release.vbs already prefers templates\Template_MedicationDispensing.xlsm when present.
# ---------------------------------------------------------------------------
