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
$vswhereExe = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'

if (-not (Test-Path $cmakeExe)) {
    throw "CMake not found: $cmakeExe"
}

function Get-VsGeneratorInstanceInfo {
    if (-not (Test-Path $vswhereExe)) {
        return $null
    }

    try {
        # Prefer an instance that actually has VC toolchain available.
        $queries = @(
            @('-all', '-latest', '-products', '*', '-requires', 'Microsoft.VisualStudio.Component.VC.Tools.x86.x64', '-version', '[17.0,18.0)', '-format', 'json'),
            @('-all', '-products', '*', '-requires', 'Microsoft.VisualStudio.Component.VC.Tools.x86.x64', '-version', '[17.0,18.0)', '-format', 'json'),
            @('-all', '-latest', '-products', '*', '-version', '[17.0,18.0)', '-format', 'json')
        )

        $candidates = @()
        foreach ($query in $queries) {
            $json = & $vswhereExe @query 2>$null
            if ([string]::IsNullOrWhiteSpace($json)) {
                continue
            }

            $items = $json | ConvertFrom-Json
            if ($items -isnot [System.Array]) {
                $items = @($items)
            }

            foreach ($item in $items) {
                $path = [string]$item.installationPath
                if ([string]::IsNullOrWhiteSpace($path)) {
                    continue
                }

                $candidates += [PSCustomObject]@{
                    Path = $path.Trim()
                    Version = ([string]$item.installationVersion).Trim()
                    IsComplete = [bool]$item.isComplete
                    IsLaunchable = [bool]$item.isLaunchable
                }
            }
        }

        $candidates = $candidates | Group-Object Path | ForEach-Object { $_.Group[0] }
        foreach ($candidate in $candidates) {
            $candidate | Add-Member -NotePropertyName HasVcvars -NotePropertyValue (Test-Path (Join-Path $candidate.Path 'VC\Auxiliary\Build\vcvars64.bat'))
        }
        $candidates = $candidates | Sort-Object -Property @(
            @{ Expression = { [int]$_.HasVcvars }; Descending = $true },
            @{ Expression = { [int]$_.IsComplete }; Descending = $true },
            @{ Expression = { [int]$_.IsLaunchable }; Descending = $true }
        )

        foreach ($candidate in $candidates) {
            if ($candidate.HasVcvars) {
                $generatorInstance = if ([string]::IsNullOrWhiteSpace($candidate.Version)) {
                    $candidate.Path
                }
                else {
                    "$($candidate.Path),version=$($candidate.Version)"
                }

                return [PSCustomObject]@{
                    Path = $candidate.Path
                    GeneratorInstance = $generatorInstance
                    IsComplete = $candidate.IsComplete
                    IsLaunchable = $candidate.IsLaunchable
                }
            }
        }

        if ($candidates.Count -gt 0) {
            $fallback = $candidates[0]
            $generatorInstance = if ([string]::IsNullOrWhiteSpace($fallback.Version)) {
                $fallback.Path
            }
            else {
                "$($fallback.Path),version=$($fallback.Version)"
            }

            return [PSCustomObject]@{
                Path = $fallback.Path
                GeneratorInstance = $generatorInstance
                IsComplete = $fallback.IsComplete
                IsLaunchable = $fallback.IsLaunchable
            }
        }

        return $null
    }
    catch {
        return $null
    }
}

function Import-VsBuildEnvironment {
    param(
        [Parameter(Mandatory = $true)][string]$VsInstance
    )

    $vcvars = Join-Path $VsInstance 'VC\Auxiliary\Build\vcvars64.bat'
    if (-not (Test-Path $vcvars)) {
        return $false
    }

    $envLines = & cmd.exe /d /s /c "`"$vcvars`" >nul && set"
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to import VS build environment from: $vcvars"
    }

    foreach ($line in $envLines) {
        $separatorIndex = $line.IndexOf('=')
        if ($separatorIndex -le 0) {
            continue
        }

        $name = $line.Substring(0, $separatorIndex)
        $value = $line.Substring($separatorIndex + 1)
        [System.Environment]::SetEnvironmentVariable($name, $value)
    }

    return $true
}

function Set-StableWindowsSdkEnvironment {
    $sdkRoot = 'C:\Program Files (x86)\Windows Kits\10\'
    if (-not (Test-Path $sdkRoot)) {
        return
    }

    # Keep MSBuild's UCRT/SDK roots consistent to avoid mixed Program Files paths.
    [System.Environment]::SetEnvironmentVariable('UCRTContentRoot', $sdkRoot)
    [System.Environment]::SetEnvironmentVariable('UniversalCRTSdkDir', $sdkRoot)
    [System.Environment]::SetEnvironmentVariable('WindowsSdkDir', $sdkRoot)
}

Push-Location $projectRoot
try {
    if ($Clean -and (Test-Path 'build')) {
        Remove-Item -Recurse -Force 'build'
    }

    $vsInstanceInfo = Get-VsGeneratorInstanceInfo
    Write-Host 'Configuring CMake project...'
    Set-StableWindowsSdkEnvironment
    if ($null -eq $vsInstanceInfo -or [string]::IsNullOrWhiteSpace($vsInstanceInfo.Path)) {
        & $cmakeExe --preset vs2022-x64
    }
    else {
        Write-Host "Using Visual Studio instance: $($vsInstanceInfo.Path)"
        if ((-not $vsInstanceInfo.IsComplete) -or (-not $vsInstanceInfo.IsLaunchable)) {
            Write-Warning "Selected Visual Studio instance is not complete/launchable. If configure still fails, run scripts/repair_vs_cpp_workload.ps1 as administrator."
        }
        $null = Import-VsBuildEnvironment -VsInstance $vsInstanceInfo.Path
        & $cmakeExe --preset vs2022-x64 -DCMAKE_GENERATOR_INSTANCE="$($vsInstanceInfo.GeneratorInstance)"
    }
    if ($LASTEXITCODE -ne 0) {
        if ($null -eq $vsInstanceInfo -or [string]::IsNullOrWhiteSpace($vsInstanceInfo.Path)) {
            throw "CMake configure failed, exit code: $LASTEXITCODE (and no Visual Studio 2022 instance detected via vswhere)."
        }

        Write-Host "Preset configure failed, retrying with explicit Visual Studio instance: $($vsInstanceInfo.Path)"
        if (Test-Path 'build') {
            Remove-Item -Recurse -Force 'build'
        }
        & $cmakeExe -S . -B build -G "Visual Studio 17 2022" -A x64 -DCMAKE_GENERATOR_INSTANCE="$($vsInstanceInfo.GeneratorInstance)"
        if ($LASTEXITCODE -ne 0) {
            throw "CMake configure failed after fallback retry, exit code: $LASTEXITCODE"
        }
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
