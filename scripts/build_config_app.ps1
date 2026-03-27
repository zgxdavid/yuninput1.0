$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$projectRoot = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $projectRoot "tools"
$outExe = Join-Path $outDir "yuninput_config.exe"
$srcPath = Join-Path $projectRoot "tools\YuninputConfig.cs"
$iconPath = Join-Path $projectRoot "assets\icon_yun.ico"

New-Item -ItemType Directory -Path $outDir -Force | Out-Null

if (-not (Test-Path $srcPath)) {
    throw "Missing source file: $srcPath"
}

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    throw "csc.exe not found"
}

if (Test-Path $iconPath) {
    & $csc /nologo /target:winexe /platform:x64 /win32icon:$iconPath /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /out:$outExe $srcPath
} else {
    & $csc /nologo /target:winexe /platform:x64 /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /out:$outExe $srcPath
}
if ($LASTEXITCODE -ne 0) {
    throw "Failed to build yuninput_config.exe"
}

Get-Item $outExe | Select-Object FullName,Length,LastWriteTime | Format-List
