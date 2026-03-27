param(
    [string]$ProjectRoot = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
}

$icoPath = Join-Path $ProjectRoot "assets\icon_yun.ico"
$sizes = @(16, 32, 48)

function Get-PreferredFontName {
    $candidates = @("SimHei", "Microsoft YaHei UI", "Microsoft YaHei", "SimSun")
    foreach ($name in $candidates) {
        try {
            $font = New-Object System.Drawing.Font($name, 12, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
            $resolved = $font.Name
            $font.Dispose()
            if ($resolved -eq $name) {
                return $name
            }
        }
        catch {
        }
    }

    return "Microsoft YaHei UI"
}

function New-RoundedRectPath([System.Drawing.RectangleF]$rect, [single]$radius) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
    $path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
    $path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function Draw-YunGlyph {
    param(
        [System.Drawing.Graphics]$Graphics,
        [int]$Size,
        [System.Drawing.Color]$StrokeColor,
        [single]$OffsetX = 0,
        [single]$OffsetY = 0,
        [single]$StrokeScale = 1.0
    )

    $strokeWidth = [single]([Math]::Max(1.6, $Size * 0.11 * $StrokeScale))
    $pen = New-Object System.Drawing.Pen($StrokeColor, $strokeWidth)
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round

    $xLT = [single]($Size * 0.29 + $OffsetX)
    $xLB = [single]($Size * 0.17 + $OffsetX)
    $xR = [single]($Size * 0.74 + $OffsetX)
    $xHook = [single]($Size * 0.55 + $OffsetX)
    $xMid1L = [single]($Size * 0.34 + $OffsetX)
    $xMid1R = [single]($Size * 0.61 + $OffsetX)
    $xMid2L = [single]($Size * 0.31 + $OffsetX)
    $xMid2R = [single]($Size * 0.56 + $OffsetX)

    $yTop = [single]($Size * 0.24 + $OffsetY)
    $yLeft = [single]($Size * 0.41 + $OffsetY)
    $yMid1 = [single]($Size * 0.45 + $OffsetY)
    $yMid2 = [single]($Size * 0.60 + $OffsetY)
    $yBottom = [single]($Size * 0.78 + $OffsetY)
    $yHook = [single]($Size * 0.87 + $OffsetY)

    $Graphics.DrawLine($pen, $xLT, $yTop, $xLB, $yLeft)
    $Graphics.DrawLine($pen, $xLT, $yTop, $xR, $yTop)
    $Graphics.DrawLine($pen, $xR, $yTop, $xR, $yBottom)
    $Graphics.DrawLine($pen, $xR, $yBottom, $xHook, $yHook)
    $Graphics.DrawLine($pen, $xMid1L, $yMid1, $xMid1R, $yMid1)
    $Graphics.DrawLine($pen, $xMid2L, $yMid2, $xMid2R, $yMid2)

    $pen.Dispose()
}

function Convert-BitmapToIcoImageBytes([System.Drawing.Bitmap]$bitmap) {
    $pixelFormat = [System.Drawing.Imaging.PixelFormat]::Format32bppArgb
    $working = New-Object System.Drawing.Bitmap($bitmap.Width, $bitmap.Height, $pixelFormat)
    $graphics = [System.Drawing.Graphics]::FromImage($working)
    $graphics.DrawImage($bitmap, 0, 0, $bitmap.Width, $bitmap.Height)
    $graphics.Dispose()

    $rect = [System.Drawing.Rectangle]::new(0, 0, $working.Width, $working.Height)
    $bitmapData = $working.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, $pixelFormat)
    try {
        $stride = [Math]::Abs($bitmapData.Stride)
        $pixelBytes = New-Object byte[] ($stride * $working.Height)
        [System.Runtime.InteropServices.Marshal]::Copy($bitmapData.Scan0, $pixelBytes, 0, $pixelBytes.Length)
    }
    finally {
        $working.UnlockBits($bitmapData)
    }

    $width = $working.Width
    $height = $working.Height
    $xorRowBytes = $width * 4
    $andRowBytes = [int]([Math]::Ceiling($width / 32.0) * 4)
    $andBytes = New-Object byte[] ($andRowBytes * $height)

    $stream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.BinaryWriter($stream)
    $writer.Write([UInt32]40)
    $writer.Write([Int32]$width)
    $writer.Write([Int32]($height * 2))
    $writer.Write([UInt16]1)
    $writer.Write([UInt16]32)
    $writer.Write([UInt32]0)
    $writer.Write([UInt32]($xorRowBytes * $height))
    $writer.Write([Int32]0)
    $writer.Write([Int32]0)
    $writer.Write([UInt32]0)
    $writer.Write([UInt32]0)

    for ($row = $height - 1; $row -ge 0; $row--) {
        $writer.Write($pixelBytes, $row * $stride, $xorRowBytes)
    }
    $writer.Write($andBytes)
    $writer.Flush()

    $bytes = $stream.ToArray()
    $writer.Dispose()
    $stream.Dispose()
    $working.Dispose()
    return $bytes
}

$icoSizes = New-Object 'System.Collections.Generic.List[int]'
$icoBytesList = New-Object 'System.Collections.Generic.List[byte[]]'
foreach ($size in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    if ($size -le 24) {
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::SingleBitPerPixelGridFit
    }
    else {
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    }

    $bgRect = [System.Drawing.RectangleF]::new([single]0.8, [single]0.8, [single]($size - 1.6), [single]($size - 1.6))
    $radius = [Math]::Max(3, $size * 0.18)
    $path = New-RoundedRectPath $bgRect $radius

    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        ([System.Drawing.PointF]::new(0, 0)),
        ([System.Drawing.PointF]::new($size, $size)),
        ([System.Drawing.Color]::FromArgb(255, 168, 36, 34)),
        ([System.Drawing.Color]::FromArgb(255, 112, 20, 20))
    )
    $g.FillPath($bgBrush, $path)

    $pen = New-Object System.Drawing.Pen(([System.Drawing.Color]::FromArgb(220, 246, 214, 152)), [Math]::Max(1, $size * 0.05))
    $g.DrawPath($pen, $path)

    if ($size -gt 24) {
        Draw-YunGlyph -Graphics $g -Size $size -StrokeColor ([System.Drawing.Color]::FromArgb(70, 36, 24, 18)) -OffsetX ([single]($size * 0.012)) -OffsetY ([single]($size * 0.016)) -StrokeScale 1.06
    }
    Draw-YunGlyph -Graphics $g -Size $size -StrokeColor ([System.Drawing.Color]::FromArgb(255, 255, 248, 236))

    $previewPath = Join-Path $ProjectRoot ("assets\icon_preview_direct_" + $size + ".png")
    $bmp.Save($previewPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $imageBytes = Convert-BitmapToIcoImageBytes $bmp
    $icoSizes.Add($size)
    $icoBytesList.Add($imageBytes)

    $pen.Dispose()
    $bgBrush.Dispose()
    $path.Dispose()
    $g.Dispose()
    $bmp.Dispose()
}

$fs = [System.IO.File]::Open($icoPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
$bw = New-Object System.IO.BinaryWriter($fs)
$bw.Write([UInt16]0)
$bw.Write([UInt16]1)
$bw.Write([UInt16]$icoSizes.Count)

$offset = 6 + (16 * $icoSizes.Count)
for ($i = 0; $i -lt $icoSizes.Count; $i++) {
    $size = $icoSizes[$i]
    $bytes = $icoBytesList[$i]
    $w = if ($size -ge 256) { [byte]0 } else { [byte]$size }
    $h = if ($size -ge 256) { [byte]0 } else { [byte]$size }
    $bw.Write($w)
    $bw.Write($h)
    $bw.Write([byte]0)
    $bw.Write([byte]0)
    $bw.Write([UInt16]1)
    $bw.Write([UInt16]32)
    $bw.Write([UInt32]$bytes.Length)
    $bw.Write([UInt32]$offset)
    $offset += $bytes.Length
}

for ($i = 0; $i -lt $icoBytesList.Count; $i++) {
    $bytes = $icoBytesList[$i]
    $bw.Write($bytes, 0, $bytes.Length)
}

$bw.Flush()
$bw.Dispose()
$fs.Dispose()

Write-Host "Generated icon: $icoPath"
