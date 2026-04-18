param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

# Force UTF-8 console I/O for better compatibility across PowerShell versions.
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$projectRoot = Split-Path -Parent $PSScriptRoot
$cmakeExe = Join-Path $env:ProgramFiles 'CMake\bin\cmake.exe'

if (-not (Test-Path $cmakeExe)) {
    throw "CMake not found: $cmakeExe"
}

Push-Location $projectRoot
try {
    if ($Clean -and (Test-Path 'build')) {
        Remove-Item -Recurse -Force 'build'
    }

    Write-Host 'Configuring CMake project...'
    & $cmakeExe --preset vs2022-x64
    if ($LASTEXITCODE -ne 0) {
        throw "CMake configure failed, exit code: $LASTEXITCODE"
    }

    Write-Host 'Building Release...'
    & $cmakeExe --build --preset release
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed, exit code: $LASTEXITCODE"
    }

    Write-Host 'Build complete: build\\Release\\yuninput.dll'
}
finally {
    Pop-Location
}
