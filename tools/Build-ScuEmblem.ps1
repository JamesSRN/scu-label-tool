# Copy manual crop -> scu_emblem.png and force all visible pixels to pure black
# (thermal labels print gray anti-aliasing poorly; snake/mark need to be solid black).
$ErrorActionPreference = 'Stop'

$repo = 'C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing'
$source = Join-Path $repo 'cropped_Black SCU Logo + Transparent Background - Copy.png'
$out = Join-Path $repo 'scu_emblem.png'

if (-not (Test-Path -LiteralPath $source)) {
    throw "Missing manual crop: $source"
}

Add-Type -AssemblyName System.Drawing

$src = [System.Drawing.Bitmap]::FromFile($source)
$bmp = New-Object System.Drawing.Bitmap $src.Width, $src.Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$bmp.SetResolution($src.HorizontalResolution, $src.VerticalResolution)

$rect = New-Object System.Drawing.Rectangle 0, 0, $src.Width, $src.Height
$srcData = $src.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, $src.PixelFormat)
$dstData = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::WriteOnly, $bmp.PixelFormat)

try {
    $stride = [Math]::Abs($srcData.Stride)
    $h = $src.Height
    $bytes = $stride * $h
    $srcBuf = New-Object byte[] $bytes
    $dstBuf = New-Object byte[] $bytes
    [Runtime.InteropServices.Marshal]::Copy($srcData.Scan0, $srcBuf, 0, $bytes)
    for ($i = 0; $i -lt $bytes; $i += 4) {
        $a = $srcBuf[$i + 3]
        if ($a -gt 32) {
            $dstBuf[$i] = 0
            $dstBuf[$i + 1] = 0
            $dstBuf[$i + 2] = 0
            $dstBuf[$i + 3] = 255
        } else {
            $dstBuf[$i] = 0
            $dstBuf[$i + 1] = 0
            $dstBuf[$i + 2] = 0
            $dstBuf[$i + 3] = 0
        }
    }
    [Runtime.InteropServices.Marshal]::Copy($dstBuf, 0, $dstData.Scan0, $bytes)
} finally {
    $src.UnlockBits($srcData)
    $bmp.UnlockBits($dstData)
}

$src.Dispose()
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

$check = [System.Drawing.Image]::FromFile($out)
Write-Host "Wrote scu_emblem.png ($($check.Width)x$($check.Height), aspect $([Math]::Round($check.Width / [double]$check.Height, 3)), pure black)"
$check.Dispose()
