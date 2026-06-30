# Same legacy emblem framing (256x180 snake crop), rendered at high resolution from source.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$repo = 'C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing'
$source = Join-Path $repo 'Black SCU Logo + Transparent Background.png'
$reference = Join-Path $repo 'tools\_embedded_emblem.png'
$out = Join-Path $repo 'scu_emblem.png'

# Legacy emblem aspect from the workbook design.
$legacyW = 256
$legacyH = 180
$outW = 1024
$outH = 720

function Get-Ink([System.Drawing.Color]$c) {
    if ($c.A -lt 16) { return 0.0 }
    $lum = (0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B) / 255.0
    return [Math]::Min(1.0, $lum * ($c.A / 255.0))
}

function Convert-ToBlackEmblem([System.Drawing.Bitmap]$bmp) {
    $w = $bmp.Width; $h = $bmp.Height
    $out = New-Object System.Drawing.Bitmap $w, $h, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    for ($y = 0; $y -lt $h; $y++) {
        for ($x = 0; $x -lt $w; $x++) {
            $ink = Get-Ink ($bmp.GetPixel($x, $y))
            if ($ink -lt 0.08) {
                $out.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, 0, 0, 0))
            } else {
                $alpha = [Math]::Min(255, [int]($ink * 255))
                $out.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
            }
        }
    }
    return $out
}

if (-not (Test-Path $reference)) { throw "Missing reference emblem: $reference" }

$refImg = [System.Drawing.Image]::FromFile($reference)
$refBmp = New-Object System.Drawing.Bitmap $refImg
$legacyW = $refBmp.Width
$legacyH = $refBmp.Height
$outW = $legacyW * 4
$outH = $legacyH * 4
$refAspect = $legacyW / [double]$legacyH

$srcImg = [System.Drawing.Image]::FromFile($source)
$srcBmp = New-Object System.Drawing.Bitmap $srcImg
Write-Host "Legacy design: ${legacyW}x${legacyH} -> ${outW}x${outH}"
Write-Host "Source logo: $($srcBmp.Width)x$($srcBmp.Height)"

# Find the crop box in the source that best matches the legacy emblem (coarse pass).
$previewW = 480
$previewH = [int][Math]::Round($srcBmp.Height * ($previewW / [double]$srcBmp.Width))
$preview = New-Object System.Drawing.Bitmap $previewW, $previewH, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$pg = [System.Drawing.Graphics]::FromImage($preview)
$pg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$pg.DrawImage($srcBmp, 0, 0, $previewW, $previewH)
$pg.Dispose()

$thumbH = [int][Math]::Round(64 / $refAspect)
$thumb = New-Object System.Drawing.Bitmap 64, $thumbH, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$tg = [System.Drawing.Graphics]::FromImage($thumb)
$tg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$tg.DrawImage($refBmp, 0, 0, $thumb.Width, $thumb.Height)
$tg.Dispose()

$bestScore = -1.0
$best = @{ X = 0; Y = 0; W = 0; H = 0 }
$maxY = [int]($previewH * 0.62)
for ($h = [int]($previewH * 0.22); $h -le [int]($previewH * 0.38); $h += 8) {
    $w = [int][Math]::Round($h * $refAspect)
    for ($y = [int]($previewH * 0.05); $y -le ($maxY - $h); $y += 8) {
        for ($x = [int]($previewW * 0.30); $x -le [int]($previewW * 0.70 - $w); $x += 8) {
            $patch = New-Object System.Drawing.Bitmap $thumb.Width, $thumb.Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $gg = [System.Drawing.Graphics]::FromImage($patch)
            $gg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $gg.DrawImage($preview, (New-Object System.Drawing.Rectangle 0, 0, $thumb.Width, $thumb.Height), (New-Object System.Drawing.Rectangle $x, $y, $w, $h), [System.Drawing.GraphicsUnit]::Pixel)
            $gg.Dispose()
            $score = 0.0
            for ($py = 0; $py -lt $thumb.Height; $py++) {
                for ($px = 0; $px -lt $thumb.Width; $px++) {
                    $rv = Get-Ink ($thumb.GetPixel($px, $py))
                    $sv = Get-Ink ($patch.GetPixel($px, $py))
                    $score += 1.0 - [Math]::Abs($rv - $sv)
                    if ($rv -gt 0.35 -and $sv -gt 0.35) { $score += 0.35 }
                }
            }
            $patch.Dispose()
            $score = $score / ($thumb.Width * $thumb.Height)
            if ($score -gt $bestScore) {
                $bestScore = $score
                $best.X = $x; $best.Y = $y; $best.W = $w; $best.H = $h
            }
        }
    }
}
Write-Host ("Matched crop: x={0} y={1} w={2} h={3} score={4}" -f $best.X, $best.Y, $best.W, $best.H, [Math]::Round($bestScore,4))

$scale = $srcBmp.Width / [double]$previewW
$srcX = [int][Math]::Floor($best.X * $scale)
$srcY = [int][Math]::Floor($best.Y * $scale)
$srcCropW = [int][Math]::Ceiling($best.W * $scale)
$srcCropH = [int][Math]::Ceiling($best.H * $scale)
$srcX = [Math]::Max(0, $srcX)
$srcY = [Math]::Max(0, $srcY)
if ($srcX + $srcCropW -gt $srcBmp.Width) { $srcCropW = $srcBmp.Width - $srcX }
if ($srcY + $srcCropH -gt $srcBmp.Height) { $srcCropH = $srcBmp.Height - $srcY }

$scaled = New-Object System.Drawing.Bitmap $outW, $outH, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$sg = [System.Drawing.Graphics]::FromImage($scaled)
$sg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$sg.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$sg.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$sg.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
$srcRect = New-Object System.Drawing.Rectangle $srcX, $srcY, $srcCropW, $srcCropH
$destRect = New-Object System.Drawing.Rectangle 0, 0, $outW, $outH
$sg.DrawImage($srcBmp, $destRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
$sg.Dispose()

$final = Convert-ToBlackEmblem $scaled
$scaled.Dispose()
$final.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
Write-Host "Wrote $out (${outW}x${outH}, $([IO.File]::ReadAllBytes($out).Length) bytes)"

$b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($out))
$chunk = 96
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('Private Function LogoB64() As String')
$lines.Add('    Dim b As String')
for ($i = 0; $i -lt $b64.Length; $i += $chunk) {
    $len = [Math]::Min($chunk, $b64.Length - $i)
    $part = $b64.Substring($i, $len)
    if ($i -eq 0) { $lines.Add("    b = `"$part`"") }
    else { $lines.Add("    b = b & `"$part`"") }
}
$lines.Add('    LogoB64 = b')
$lines.Add('End Function')
$b64Path = Join-Path $repo 'tools\LogoB64.generated.txt'
$lines -join "`r`n" | Set-Content -Path $b64Path -Encoding UTF8
Write-Host "Wrote $b64Path ($($b64.Length) chars)"

$preview.Dispose(); $thumb.Dispose(); $refBmp.Dispose(); $refImg.Dispose(); $srcBmp.Dispose(); $srcImg.Dispose(); $final.Dispose()
