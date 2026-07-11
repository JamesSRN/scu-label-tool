<#
  make-release-zip.ps1 - Build the downloadable ZIP for a GitHub Release.

  RUN LOCALLY from the repo root. Produces:
      dist\SCU-Label-Printing-v<Version>.zip
  containing: MedicationDispensing.xlsm + scu_emblem.png + SCU_QuickStart_Card.pdf
  + INSTALL.txt (end-user instructions).

  IMPORTANT - PHI: the packaged workbook must contain NO patient data (empty Log,
  no name/DOB). Use a freshly built, never-used copy or a clean template. The script
  makes you confirm before it packages.

  Usage:
    powershell -ExecutionPolicy Bypass -File tools\make-release-zip.ps1 -Version 2.0
    powershell -ExecutionPolicy Bypass -File tools\make-release-zip.ps1 -Version 2.0 -Workbook "C:\path\clean.xlsm"
#>
param(
    [string]$Version = "2.0",
    [string]$Workbook = ""
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# Choose the workbook to package (prefer a clean template, else the working copy).
if ($Workbook -eq "") {
    $Workbook = @(
        (Join-Path $root 'templates\Template_MedicationDispensing.xlsm'),
        (Join-Path $root 'MedicationDispensing.xlsm')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $Workbook -or -not (Test-Path $Workbook)) {
    Write-Host "No workbook found to package. Pass -Workbook <path to a clean .xlsm>." -ForegroundColor Red
    exit 1
}

$emblem = Join-Path $root 'scu_emblem.png'
$card   = Join-Path $root 'docs\SCU_QuickStart_Card.pdf'
foreach ($f in @($emblem, $card)) {
    if (-not (Test-Path $f)) { Write-Host "Missing required file: $f" -ForegroundColor Red; exit 1 }
}

Write-Host ""
Write-Host "Packaging SCU Label Printing v$Version" -ForegroundColor Cyan
Write-Host "  workbook: $Workbook"
Write-Host ""
Write-Host "  *** PHI CHECK: confirm this workbook has NO patient data and an empty Log. ***" -ForegroundColor Yellow
$ans = Read-Host "  Type YES to continue"
if ($ans -ne 'YES') { Write-Host "Aborted." -ForegroundColor Red; exit 1 }

$stage = Join-Path $env:TEMP ("scu_rel_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $stage | Out-Null
Copy-Item -LiteralPath $Workbook -Destination (Join-Path $stage 'MedicationDispensing.xlsm')
Copy-Item -LiteralPath $emblem   -Destination (Join-Path $stage 'scu_emblem.png')
Copy-Item -LiteralPath $card     -Destination (Join-Path $stage 'SCU_QuickStart_Card.pdf')

$install = @"
SCU Label Printing  -  v$Version
================================

WHAT THIS IS
  An offline Excel tool for printing medication labels on a Brother QL-1100c
  (DK-1202 62 x 100 mm) for the Saturday Clinic for the Uninsured.

INSTALL (one time)
  1. Keep MedicationDispensing.xlsm and scu_emblem.png together in ONE folder.
  2. Open MedicationDispensing.xlsm in Microsoft Excel (Windows).
  3. If prompted, ENABLE MACROS - the tool is macro-driven.
  4. Recommended: add this folder as a Trusted Location
     (File > Options > Trust Center > Trust Center Settings > Trusted Locations).
  5. Install the Brother QL-1100c driver and load the DK-1202 (62 x 100 mm) roll.

USE
  The workbook opens to a "Start Here" guide - follow the 4 steps.
  A printable version is included: SCU_QuickStart_Card.pdf.

PRIVACY
  Patient info and the on-screen Log are cleared when you close the file.
  A dated CSV copy of each day's dispensing is saved locally in a "dispense-log"
  folder next to the workbook (for your records) - it never leaves this PC.
"@
Set-Content -Path (Join-Path $stage 'INSTALL.txt') -Value $install -Encoding ASCII

$distDir = Join-Path $root 'dist'
New-Item -ItemType Directory -Force -Path $distDir | Out-Null
$zip = Join-Path $distDir ("SCU-Label-Printing-v$Version.zip")
if (Test-Path $zip) { Remove-Item -Force $zip }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item -Recurse -Force $stage

Write-Host ""
Write-Host "Created: $zip" -ForegroundColor Green
Write-Host "Next: create a GitHub Release (tag v$Version) and attach this ZIP. See docs\RELEASE.md." -ForegroundColor Green
Write-Host ""
