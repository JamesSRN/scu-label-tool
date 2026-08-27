<#
  check-encoding.ps1 - Validate the SCU Label Printing VBA / VBScript source files.

  Excel's VBA importer is picky: MedParser.bas and Build-Release.vbs MUST be pure
  ASCII with Windows CRLF line endings, or the import mojibakes or fails. This script
  catches those problems (and, as an advisory, some structural ones) BEFORE a bad file
  is built.

  Checks per file:
    ENCODING (build-breaking -> exit 1):
      - no UTF-8 / UTF-16 byte-order mark (BOM)
      - pure ASCII (no byte >= 128)
      - Windows CRLF line endings (no lone LF, no doubled CR)
    STRUCTURE (advisory -> exit 3):
      - balanced block keywords: Sub, Function, Property, With, For, Do, Select, Type
      (advisory because single-line forms make this a heuristic; the real structural
       check is the VBA compile that SetupWorkbook forces during the build.)

  Exit codes:
     0 = all good
     1 = ENCODING problem (do not build)
     2 = a target file was missing
     3 = STRUCTURE (balance) warning only; encoding is fine

  Usage (from the repo root):
    powershell -ExecutionPolicy Bypass -File tools\check-encoding.ps1
    powershell -ExecutionPolicy Bypass -File tools\check-encoding.ps1 MedParser.bas
#>
param([string[]]$Files)

$root = Split-Path -Parent $PSScriptRoot   # tools\ -> repo root
if (-not $Files -or $Files.Count -eq 0) {
    $Files = @(
        (Join-Path $root 'MedParser.bas')
    )
}

$script:encProblems = 0
$script:balProblems = 0
$script:missing     = 0

function Count-Lines([string[]]$ls, [string]$pattern) {
    ($ls | Where-Object { $_ -match $pattern }).Count
}

function Test-SourceFile([string]$path) {
    if (-not (Test-Path $path)) {
        Write-Host ("  MISSING: {0}" -f $path) -ForegroundColor Red
        $script:missing++
        return
    }
    $name  = Split-Path -Leaf $path
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $enc = 0
    $bal = 0

    # --- byte-order marks ---
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Write-Host ("  [BOM]     {0}: UTF-8 BOM at start of file (must be none)" -f $name) -ForegroundColor Red
        $enc++
    }
    if ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF))) {
        Write-Host ("  [BOM]     {0}: UTF-16 BOM at start of file" -f $name) -ForegroundColor Red
        $enc++
    }

    # --- non-ASCII + line-ending scan (walk bytes, track line numbers) ---
    $line = 1
    $nonAscii = 0; $firstNonAscii = 0
    $loneLf   = 0; $firstLoneLf   = 0
    $doubleCr = 0; $firstDoubleCr = 0
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $b = $bytes[$i]
        if ($b -ge 128) {
            $nonAscii++
            if ($firstNonAscii -eq 0) { $firstNonAscii = $line }
        }
        if ($b -eq 13 -and ($i + 1) -lt $bytes.Length -and $bytes[$i + 1] -eq 13) {
            $doubleCr++
            if ($firstDoubleCr -eq 0) { $firstDoubleCr = $line }
        }
        if ($b -eq 10) {
            $prev = if ($i -gt 0) { $bytes[$i - 1] } else { 0 }
            if ($prev -ne 13) {
                $loneLf++
                if ($firstLoneLf -eq 0) { $firstLoneLf = $line }
            }
            $line++
        }
    }
    if ($nonAscii -gt 0) {
        Write-Host ("  [ASCII]   {0}: {1} non-ASCII byte(s), first near line {2}" -f $name, $nonAscii, $firstNonAscii) -ForegroundColor Red
        $enc++
    }
    if ($loneLf -gt 0) {
        Write-Host ("  [CRLF]    {0}: {1} line(s) not ending in CRLF, first near line {2}" -f $name, $loneLf, $firstLoneLf) -ForegroundColor Red
        $enc++
    }
    if ($doubleCr -gt 0) {
        Write-Host ("  [CR]      {0}: {1} doubled-CR sequence(s), first near line {2}" -f $name, $doubleCr, $firstDoubleCr) -ForegroundColor Red
        $enc++
    }

    # --- structural balance (advisory) ---
    $text  = [System.Text.Encoding]::ASCII.GetString($bytes)
    $lines = $text -split "`n"
    $pairs = @(
        @('Sub',      '^\s*(Public\s+|Private\s+|Friend\s+|Static\s+)*Sub\b',      '^\s*End\s+Sub\b'),
        @('Function', '^\s*(Public\s+|Private\s+|Friend\s+|Static\s+)*Function\b', '^\s*End\s+Function\b'),
        @('Property', '^\s*(Public\s+|Private\s+|Friend\s+|Static\s+)*Property\b', '^\s*End\s+Property\b'),
        @('With',     '^\s*With\b',          '^\s*End\s+With\b'),
        @('For',      '^\s*For\b',           '^\s*Next\b'),
        @('Do',       '^\s*Do\b',            '^\s*Loop\b'),
        @('Select',   '^\s*Select\s+Case\b', '^\s*End\s+Select\b'),
        @('Type',     '^\s*(Public\s+|Private\s+)*Type\b', '^\s*End\s+Type\b')
    )
    foreach ($p in $pairs) {
        $open  = Count-Lines $lines $p[1]
        $close = Count-Lines $lines $p[2]
        if ($open -ne $close) {
            Write-Host ("  [BALANCE] {0}: {1} open={2} close={3} (advisory)" -f $name, $p[0], $open, $close) -ForegroundColor Yellow
            $bal++
        }
    }

    if ($enc -eq 0 -and $bal -eq 0) {
        Write-Host ("  OK: {0}" -f $name) -ForegroundColor Green
    } elseif ($enc -eq 0) {
        Write-Host ("  ENCODING OK (structure advisory only): {0}" -f $name) -ForegroundColor Yellow
    }
    $script:encProblems += $enc
    $script:balProblems += $bal
}

Write-Host ""
Write-Host "SCU Label Printing - source encoding + structure check"
Write-Host "------------------------------------------------------"
foreach ($f in $Files) { Test-SourceFile $f }
Write-Host ""

if ($script:missing -gt 0) {
    Write-Host ("FAIL: {0} target file(s) missing." -f $script:missing) -ForegroundColor Red
    exit 2
}
if ($script:encProblems -gt 0) {
    Write-Host ("FAIL: {0} ENCODING problem(s) found - do NOT build until fixed." -f $script:encProblems) -ForegroundColor Red
    exit 1
}
if ($script:balProblems -gt 0) {
    Write-Host ("WARN: {0} structure balance warning(s). Encoding is fine; the VBA compile during build is the authoritative structural check." -f $script:balProblems) -ForegroundColor Yellow
    exit 3
}
Write-Host "PASS: all files are pure ASCII, CRLF, and structurally balanced." -ForegroundColor Green
exit 0
