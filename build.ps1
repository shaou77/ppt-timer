$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $root 'src\PptTimer.cs'
$output = Join-Path $root 'dist'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'

New-Item -ItemType Directory -Force -Path $output | Out-Null

& $compiler /nologo /target:winexe /optimize+ /platform:anycpu `
    /out:"$output\PPT-Timer.exe" `
    /reference:System.dll `
    /reference:System.Drawing.dll `
    /reference:System.Windows.Forms.dll `
    $source

if ($LASTEXITCODE -ne 0) {
    throw 'Build failed'
}

Copy-Item (Join-Path $root 'README.md') (Join-Path $output 'README.md') -Force
Write-Host "Built: $output"
