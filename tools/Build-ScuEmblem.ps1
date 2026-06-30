# Use the manually cropped clinic emblem (Copy file) as scu_emblem.png.
$ErrorActionPreference = 'Stop'

$repo = 'C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing'
$source = Join-Path $repo 'cropped_Black SCU Logo + Transparent Background - Copy.png'
$out = Join-Path $repo 'scu_emblem.png'

if (-not (Test-Path -LiteralPath $source)) {
    throw "Missing manual crop: $source"
}

Copy-Item -LiteralPath $source -Destination $out -Force

$raw = [IO.File]::ReadAllBytes($out)
$w = ($raw[16] -shl 24) -bor ($raw[17] -shl 16) -bor ($raw[18] -shl 8) -bor $raw[19]
$h = ($raw[20] -shl 24) -bor ($raw[21] -shl 16) -bor ($raw[22] -shl 8) -bor $raw[23]
Write-Host "Copied manual crop to scu_emblem.png (${w}x${h}, aspect $([Math]::Round($w / [double]$h, 3)))"
