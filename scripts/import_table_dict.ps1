param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [string]$OutputPath,

    [string]$SourceEncoding = 'utf-8',

    [switch]$Append,

    [switch]$KeepPromptEntries
)

$ErrorActionPreference = 'Stop'

$SourcePath = (Resolve-Path $SourcePath).Path
if (-not $OutputPath) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $OutputPath = Join-Path (Split-Path -Parent $PSScriptRoot | Join-Path -ChildPath 'data') ($name + '.dict')
}

if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path (Get-Location) $OutputPath
}

$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$sourceTextEncoding = [System.Text.Encoding]::GetEncoding($SourceEncoding)
$outputDir = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$rulesOutputPath = $OutputPath + '.rules'

$metadataLines = [System.Collections.Generic.List[string]]::new()
$constructPhrase = $null
$ruleLines = [System.Collections.Generic.List[string]]::new()

function Test-CodeToken {
    param([string]$Token)

    return $Token -match '^[\^]?[A-Za-z]+$'
}

function Convert-CodeToken {
    param([string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token)) {
        return $null
    }

    if ($Token.StartsWith('@')) {
        return $null
    }

    return ($Token -replace '^\^', '').ToLowerInvariant()
}

$rows = [System.Collections.Generic.List[string]]::new()
$inDataBlock = $false
$inRuleBlock = $false

foreach ($rawLine in [System.IO.File]::ReadLines($SourcePath, $sourceTextEncoding)) {
    $line = $rawLine.Trim()
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }

    if ($line.StartsWith('###') -or $line.StartsWith('#')) {
        continue
    }

    if ($line -match '^ConstructPhrase=(.+)$') {
        $constructPhrase = $Matches[1].Trim()
        continue
    }

    if ($line -match '^BEGIN_TABLE$') {
        $inDataBlock = $true
        $inRuleBlock = $false
        continue
    }

    if ($line -match '^\[.*\]$') {
        if ($line -match '^(\[Rule\]|\[组词规则\])$') {
            $inRuleBlock = $true
            $inDataBlock = $false
            continue
        }

        if ($line -match '^(\[Data\]|\[数据\])$' -or ($line -match '^\[.*\]$' -and $line -notmatch '[A-Za-z]' -and $line -notmatch '规则')) {
            $inDataBlock = $true
            $inRuleBlock = $false
        }
        continue
    }

    if ($inRuleBlock -and $line -match '^[A-Za-z][0-9]+=') {
        $ruleLines.Add($line)
        continue
    }

    if ($line -match '^END_TABLE$' -or $line -match '^END_TABlE$' -or $line -match '^BEGIN_GOUCI$') {
        break
    }

    if (-not $inDataBlock) {
        continue
    }

    $parts = $line -split '\s+'
    if ($parts.Count -lt 2) {
        continue
    }

    $first = $parts[0]
    $second = $parts[1]
    $third = if ($parts.Count -ge 3) { $parts[2] } else { $null }

    if (-not $KeepPromptEntries -and $first.StartsWith('&')) {
        continue
    }

    $code = $null
    $text = $null
    if ((Test-CodeToken $first) -and -not (Test-CodeToken $second)) {
        $code = Convert-CodeToken $first
        $text = $second
    } elseif ((Test-CodeToken $second) -and -not (Test-CodeToken $first)) {
        $code = Convert-CodeToken $second
        $text = $first
    } else {
        continue
    }

    if ([string]::IsNullOrWhiteSpace($code) -or [string]::IsNullOrWhiteSpace($text)) {
        continue
    }

    if ($third -and $third -match '^\d+$') {
        $rows.Add("$code $text $third")
    } else {
        $rows.Add("$code $text")
    }
}

if ($rows.Count -eq 0) {
    throw "No usable table rows found in source file: $SourcePath"
}

if (-not [string]::IsNullOrWhiteSpace($constructPhrase)) {
    $metadataLines.Add("# yuninput:construct_phrase=$constructPhrase")
}
foreach ($ruleLine in $ruleLines) {
    $metadataLines.Add("# yuninput:rule:$ruleLine")
}

$allLines = [System.Collections.Generic.List[string]]::new()
foreach ($metadataLine in $metadataLines) {
    $allLines.Add($metadataLine)
}
foreach ($row in $rows) {
    $allLines.Add($row)
}

if ($Append -and (Test-Path $OutputPath)) {
    [System.IO.File]::AppendAllLines($OutputPath, $rows, $utf8NoBom)
} else {
    [System.IO.File]::WriteAllLines($OutputPath, $allLines, $utf8NoBom)
}

if ($metadataLines.Count -gt 0) {
    [System.IO.File]::WriteAllLines($rulesOutputPath, $metadataLines, $utf8NoBom)
} elseif (Test-Path $rulesOutputPath) {
    Remove-Item -Force $rulesOutputPath
}

Write-Host "Imported $($rows.Count) rows to $OutputPath"