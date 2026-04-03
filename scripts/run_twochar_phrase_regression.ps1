param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$SingleDictRelativePath = 'data/zhengma-single.dict',
    [string]$PhraseSourceRelativePath = 'data/zhengma-all.dict',
    [string]$OutputRelativePath = 'diagnostics/twochar_phrase_regression_latest.tsv',
    [int]$SampleSize = 200,
    [string[]]$FocusPhrases = @('新泾', '华益')
)

$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom
chcp 65001 | Out-Null

function Resolve-RepoPath([string]$basePath, [string]$relativePath) {
    return [System.IO.Path]::GetFullPath((Join-Path $basePath $relativePath))
}

function Is-CodeToken([string]$token) {
    return -not [string]::IsNullOrWhiteSpace($token) -and $token -cmatch '^[A-Za-z]+$'
}

function Try-ParseDictionaryEntry([string]$line) {
    if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) {
        return $null
    }

    $parts = $line -split '\s+'
    if ($parts.Count -lt 2) {
        return $null
    }

    $first = $parts[0]
    $second = $parts[1]
    $firstIsCode = Is-CodeToken $first
    $secondIsCode = Is-CodeToken $second
    if ($firstIsCode -eq $secondIsCode) {
        return $null
    }

    if ($firstIsCode) {
        return [PSCustomObject]@{
            Code = $first.ToLowerInvariant()
            Text = $second
        }
    }

    return [PSCustomObject]@{
        Code = $second.ToLowerInvariant()
        Text = $first
    }
}

function Get-CodeChar([string]$code, [int]$index) {
    if ([string]::IsNullOrEmpty($code)) {
        throw 'Empty code is not allowed.'
    }

    if ($index -le 0) {
        return $code.Substring(0, 1)
    }

    if ($code.Length -eq 1) {
        return 'v'
    }

    $safeIndex = [Math]::Min($index, $code.Length - 1)
    return $code.Substring($safeIndex, 1)
}

function Get-ExpectedTwoCharCodes([string[]]$firstCodes, [string[]]$secondCodes) {
    $codes = @{}
    foreach ($firstCode in $firstCodes) {
        foreach ($secondCode in $secondCodes) {
            for ($firstTailIndex = 1; $firstTailIndex -le 3; $firstTailIndex++) {
                for ($secondTailIndex = 1; $secondTailIndex -le 3; $secondTailIndex++) {
                    $phraseCode =
                        (Get-CodeChar $firstCode 0) +
                        (Get-CodeChar $firstCode $firstTailIndex) +
                        (Get-CodeChar $secondCode 0) +
                        (Get-CodeChar $secondCode $secondTailIndex)
                    $codes[$phraseCode] = $true
                }
            }
        }
    }

    return @($codes.Keys | Sort-Object)
}

function Invoke-PhraseProbe([string]$probePath, [string]$dictPath, [string]$phrase) {
    $escapedProbe = '"' + $probePath + '"'
    $escapedDict = '"' + $dictPath + '"'
    $escapedPhrase = '"' + $phrase + '"'
    $command = "$escapedProbe $escapedDict --phrase $escapedPhrase 2>nul"
    $lines = @(cmd.exe /d /c $command)
    if ($LASTEXITCODE -ne 0) {
        return @()
    }

    return @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
}

$repo = [System.IO.Path]::GetFullPath($RepoRoot)
$probePath = Resolve-RepoPath $repo 'build/Release/yuninput_sort_probe.exe'
$singleDictPath = Resolve-RepoPath $repo $SingleDictRelativePath
$phraseSourcePath = Resolve-RepoPath $repo $PhraseSourceRelativePath
$outputPath = Resolve-RepoPath $repo $OutputRelativePath

foreach ($requiredPath in @($probePath, $singleDictPath, $phraseSourcePath)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Missing required path: $requiredPath"
    }
}

$outputDir = Split-Path $outputPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$charCodeMap = @{}
foreach ($line in Get-Content -Path $singleDictPath -Encoding UTF8) {
    $entry = Try-ParseDictionaryEntry $line
    if ($null -eq $entry -or $entry.Text.Length -ne 1) {
        continue
    }

    if (-not $charCodeMap.ContainsKey($entry.Text)) {
        $charCodeMap[$entry.Text] = @()
    }
    if ($charCodeMap[$entry.Text] -notcontains $entry.Code) {
        $charCodeMap[$entry.Text] += $entry.Code
    }
}

$compatibleCharCodeMap = @{}
foreach ($key in $charCodeMap.Keys) {
    $codes = @($charCodeMap[$key])
    if ($codes.Count -eq 0) {
        continue
    }

    $compatibleCharCodeMap[$key] = @($codes | Sort-Object Length, @{ Expression = { $_ } } -Unique)
}

$phrases = New-Object System.Collections.ArrayList
$seenPhrases = @{}

foreach ($phrase in $FocusPhrases) {
    if ([string]::IsNullOrWhiteSpace($phrase) -or $phrase.Length -ne 2) {
        continue
    }
    if (-not $seenPhrases.ContainsKey($phrase)) {
        $seenPhrases[$phrase] = $true
        [void]$phrases.Add($phrase)
    }
}

foreach ($line in Get-Content -Path $phraseSourcePath -Encoding UTF8) {
    if ($phrases.Count -ge ($FocusPhrases.Count + $SampleSize)) {
        break
    }

    $entry = Try-ParseDictionaryEntry $line
    if ($null -eq $entry -or $entry.Text.Length -ne 2) {
        continue
    }

    if ($seenPhrases.ContainsKey($entry.Text)) {
        continue
    }
    $seenPhrases[$entry.Text] = $true

    $firstChar = $entry.Text.Substring(0, 1)
    $secondChar = $entry.Text.Substring(1, 1)
    if (-not $compatibleCharCodeMap.ContainsKey($firstChar) -or -not $compatibleCharCodeMap.ContainsKey($secondChar)) {
        continue
    }

    [void]$phrases.Add($entry.Text)
}

$rows = New-Object System.Collections.Generic.List[string]
$rows.Add('phrase`tstatus`tchar1_codes`tchar2_codes`texpected_count`tactual_count`tmissing_codes')
$failures = New-Object System.Collections.Generic.List[object]
$skippedNoBuild = New-Object System.Collections.Generic.List[object]

foreach ($phrase in $phrases) {
    $firstChar = $phrase.Substring(0, 1)
    $secondChar = $phrase.Substring(1, 1)
    $firstCodes = @($compatibleCharCodeMap[$firstChar])
    $secondCodes = @($compatibleCharCodeMap[$secondChar])
    $expectedCodes = @(Get-ExpectedTwoCharCodes $firstCodes $secondCodes)
    $actualCodes = @(Invoke-PhraseProbe $probePath $singleDictPath $phrase | Sort-Object -Unique)

    if ($actualCodes.Count -eq 0) {
        $rows.Add(
            ($phrase + "`t" +
            ('no-build' + "`t") +
            (($firstCodes -join ',') + "`t") +
            (($secondCodes -join ',') + "`t") +
            ($expectedCodes.Count.ToString() + "`t") +
            ('0' + "`t")))
        $skippedNoBuild.Add([PSCustomObject]@{
            Phrase = $phrase
            FirstCodes = $firstCodes -join ','
            SecondCodes = $secondCodes -join ','
        }) | Out-Null
        continue
    }

    $actualSet = @{}
    foreach ($actualCode in $actualCodes) {
        $actualSet[$actualCode.Trim()] = $true
    }

    $missingCodes = New-Object System.Collections.Generic.List[string]
    foreach ($expectedCode in $expectedCodes) {
        if (-not $actualSet.ContainsKey($expectedCode)) {
            $missingCodes.Add($expectedCode) | Out-Null
        }
    }

    $rows.Add(
        ($phrase + "`t" +
        ('checked' + "`t") +
        (($firstCodes -join ',') + "`t") +
        (($secondCodes -join ',') + "`t") +
        ($expectedCodes.Count.ToString() + "`t") +
        ($actualCodes.Count.ToString() + "`t") +
        ($missingCodes -join ',')))

    if ($missingCodes.Count -gt 0) {
        $failures.Add([PSCustomObject]@{
            Phrase = $phrase
            FirstCodes = $firstCodes -join ','
            SecondCodes = $secondCodes -join ','
            MissingCodes = $missingCodes -join ','
        }) | Out-Null
    }
}

Set-Content -Path $outputPath -Value $rows -Encoding UTF8
Write-Host "Two-char regression snapshot written: $outputPath"
Write-Host ("Phrases checked: {0}" -f $phrases.Count)
Write-Host ("No-build samples skipped: {0}" -f $skippedNoBuild.Count)

if ($failures.Count -eq 0) {
    Write-Host 'Two-char regression check passed.'
    exit 0
}

Write-Host ("Two-char regression check failed: {0} phrases with missing codes." -f $failures.Count)
$preview = $failures | Select-Object -First 20
foreach ($failure in $preview) {
    Write-Host ("FAIL {0} char1={1} char2={2} missing={3}" -f $failure.Phrase, $failure.FirstCodes, $failure.SecondCodes, $failure.MissingCodes)
}
if ($failures.Count -gt $preview.Count) {
    Write-Host ("... truncated, remaining failing phrases: {0}" -f ($failures.Count - $preview.Count))
}
exit 1