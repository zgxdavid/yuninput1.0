param(
    [string]$OutputPath = "./docs/mode_samples.tsv",
    [int]$Limit = 10
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $projectRoot $OutputPath
}

$paths = @{
    'zhengma-large' = Join-Path $projectRoot 'data/zhengma-large.dict'
    'zhengma-large-pinyin' = Join-Path $projectRoot 'data/zhengma-pinyin.dict'
}

function Get-TopMap([string]$path) {
    $map = @{}
    foreach ($lineRaw in Get-Content -Path $path -Encoding UTF8) {
        $line = $lineRaw.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) {
            continue
        }

        $parts = $line -split '\s+'
        if ($parts.Count -lt 2) {
            continue
        }

        $code = $parts[0].ToLowerInvariant()
        $text = $parts[1]
        if (-not $map.ContainsKey($code)) {
            $map[$code] = New-Object System.Collections.Generic.List[string]
        }

        if ($map[$code].Count -lt 8 -and -not $map[$code].Contains($text)) {
            [void]$map[$code].Add($text)
        }
    }

    return $map
}

$maps = @{}
foreach ($k in $paths.Keys) {
    $maps[$k] = Get-TopMap $paths[$k]
}

$allCodes = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $maps.Keys) {
    foreach ($c in $maps[$k].Keys) {
        [void]$allCodes.Add($c)
    }
}

$diff = @()
foreach ($c in $allCodes) {
    if ($c.Length -lt 2 -or $c.Length -gt 5) {
        continue
    }

    $a = if ($maps['zhengma-large'].ContainsKey($c)) { ($maps['zhengma-large'][$c] -join '|') } else { '' }
    $d = if ($maps['zhengma-large-pinyin'].ContainsKey($c)) { ($maps['zhengma-large-pinyin'][$c] -join '|') } else { '' }

    if ($a -ne $d) {
        $diff += [PSCustomObject]@{
            code = $c
            large = if ($a) { $a } else { '<none>' }
            hybrid = if ($d) { $d } else { '<none>' }
            len = $c.Length
        }
    }
}

$selected = $diff | Sort-Object len, code | Select-Object -First $Limit

$dir = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

@("code`tzhengma-large`tzhengma-large-pinyin") | Set-Content -Path $OutputPath -Encoding UTF8
foreach ($row in $selected) {
    "$($row.code)`t$($row.large)`t$($row.hybrid)" | Add-Content -Path $OutputPath -Encoding UTF8
}

Write-Host "Generated: $OutputPath"
