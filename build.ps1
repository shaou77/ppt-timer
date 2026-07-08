$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $root 'src\PptTimer.cs'
$output = Join-Path $root 'dist'
$assets = Join-Path $root 'assets'
$icon = Join-Path $assets 'PPT-Timer.ico'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'

New-Item -ItemType Directory -Force -Path $output | Out-Null
New-Item -ItemType Directory -Force -Path $assets | Out-Null

Add-Type -AssemblyName System.Drawing

function New-AppIcon {
    param([string]$Path)

    $bitmap = New-Object System.Drawing.Bitmap 64, 64
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::Transparent)

    $shadow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(70, 0, 0, 0))
    $graphics.FillEllipse($shadow, 8, 10, 50, 50)

    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        [System.Drawing.Rectangle]::new(0, 0, 64, 64),
        [System.Drawing.Color]::FromArgb(30, 53, 170),
        [System.Drawing.Color]::FromArgb(70, 135, 255),
        45
    )
    $graphics.FillEllipse($bg, 4, 4, 56, 56)

    $rim = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(245, 255, 255, 255), 2)
    $graphics.DrawEllipse($rim, 6, 6, 52, 52)

    $shine = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        [System.Drawing.Rectangle]::new(10, 7, 44, 28),
        [System.Drawing.Color]::FromArgb(150, 255, 255, 255),
        [System.Drawing.Color]::FromArgb(0, 255, 255, 255),
        90
    )
    $graphics.FillEllipse($shine, 10, 7, 44, 28)

    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $graphics.FillEllipse($white, 16, 13, 32, 32)

    $faceRim = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(230, 222, 228, 255), 2)
    $graphics.DrawEllipse($faceRim, 17, 14, 30, 30)

    $tickPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(40, 68, 190), 2)
    $graphics.DrawLine($tickPen, 32, 17, 32, 20)
    $graphics.DrawLine($tickPen, 32, 38, 32, 41)
    $graphics.DrawLine($tickPen, 20, 29, 23, 29)
    $graphics.DrawLine($tickPen, 41, 29, 44, 29)

    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 151, 45), 4)
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($pen, 32, 29, 32, 20)
    $graphics.DrawLine($pen, 32, 29, 40, 34)

    $center = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 151, 45))
    $graphics.FillEllipse($center, 29, 26, 6, 6)

    $badge = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 151, 45))
    $graphics.FillEllipse($badge, 37, 38, 17, 17)
    $font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $graphics.DrawString('P', $font, $white, 42, 39)

    $iconHandle = $bitmap.GetHicon()
    $appIcon = [System.Drawing.Icon]::FromHandle($iconHandle)
    $stream = [System.IO.File]::Create($Path)
    try {
        $appIcon.Save($stream)
    }
    finally {
        $stream.Dispose()
        $appIcon.Dispose()
        $font.Dispose()
        $badge.Dispose()
        $center.Dispose()
        $pen.Dispose()
        $tickPen.Dispose()
        $faceRim.Dispose()
        $white.Dispose()
        $shine.Dispose()
        $rim.Dispose()
        $shadow.Dispose()
        $bg.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

New-AppIcon $icon

& $compiler /nologo /target:winexe /optimize+ /platform:anycpu `
    /out:"$output\PPT-Timer.exe" `
    /win32icon:"$icon" `
    /reference:System.dll `
    /reference:System.Drawing.dll `
    /reference:System.Windows.Forms.dll `
    $source

if ($LASTEXITCODE -ne 0) {
    throw 'Build failed'
}

Copy-Item (Join-Path $root 'README.md') (Join-Path $output 'README.md') -Force
Write-Host "Built: $output"
