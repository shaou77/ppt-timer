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

    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        [System.Drawing.Rectangle]::new(0, 0, 64, 64),
        [System.Drawing.Color]::FromArgb(39, 68, 190),
        [System.Drawing.Color]::FromArgb(73, 114, 255),
        45
    )
    $graphics.FillEllipse($bg, 4, 4, 56, 56)

    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $graphics.FillEllipse($white, 17, 13, 30, 30)

    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(39, 68, 190), 4)
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($pen, 32, 28, 32, 19)
    $graphics.DrawLine($pen, 32, 28, 39, 32)

    $font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $graphics.DrawString('P', $font, $white, 23, 39)

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
        $pen.Dispose()
        $white.Dispose()
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
