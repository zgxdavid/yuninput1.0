param(
    [string]$SourcePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'third_party/fcitx/zhengma-pinyin.txt'),
    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'data/zhengma-pinyin.dict')
)

$ErrorActionPreference = 'Stop'

$resolvedSource = (Resolve-Path $SourcePath).Path
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Parent $resolvedOutput
if (-not [string]::IsNullOrWhiteSpace($outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$utf8 = [System.Text.UTF8Encoding]::new($false)
$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
$rows = [System.Collections.Generic.List[string]]::new()

foreach ($line in [System.IO.File]::ReadLines($resolvedSource, $utf8)) {
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }

    $trimmed = $line.Trim()
    if ($trimmed.StartsWith('#')) {
        continue
    }

    if ($trimmed -notmatch '^@([A-Za-z]+)\s+(\S+)$') {
        continue
    }

    $code = $Matches[1].ToLowerInvariant()
    $text = $Matches[2]
    if ($text.Length -ne 1) {
        continue
    }

    $key = "$code`t$text"
    if ($seen.Add($key)) {
        $rows.Add("$code $text")
    }
}

if ($rows.Count -eq 0) {
    throw "No pinyin single-character rows were extracted from $resolvedSource"
}

[System.IO.File]::WriteAllLines($resolvedOutput, $rows, $utf8)

$rulesPath = $resolvedOutput + '.rules'
if (Test-Path $rulesPath) {
    Remove-Item -Path $rulesPath -Force
}

Write-Host "Generated $($rows.Count) rows to $resolvedOutput"