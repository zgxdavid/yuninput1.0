param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputRelativePath = 'diagnostics/textservice_behavior_regression_latest.tsv',
    [switch]$SkipWildcardProbe
)

$ErrorActionPreference = 'Stop'

chcp 65001 | Out-Null
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

function Resolve-RepoPath([string]$basePath, [string]$relativePath) {
    return [System.IO.Path]::GetFullPath((Join-Path $basePath $relativePath))
}

function Test-AllPatterns([string]$content, [string[]]$patterns) {
    foreach ($pattern in $patterns) {
        if ($content -notmatch $pattern) {
            return $false
        }
    }

    return $true
}

$repo = [System.IO.Path]::GetFullPath($RepoRoot)
$textServicePath = Resolve-RepoPath $repo 'src/TextService.cpp'
$outputPath = Resolve-RepoPath $repo $OutputRelativePath

if (-not (Test-Path $textServicePath)) {
    throw "TextService source not found: $textServicePath"
}

$textServiceContent = Get-Content -Path $textServicePath -Raw -Encoding UTF8

$cases = @(
    [PSCustomObject]@{
        Id = 'auto_commit'
        Title = 'auto commit'
        CheckType = 'source-guard'
        Patterns = @(
            'constexpr\s+size_t\s+kAutoCommitCodeLength\s*=\s*4;',
            'shouldAutoCommitDuringTyping\s*=\s*!pinyinQueryMode;',
            'TryFindUniqueExactCommitCandidateIndex'
        )
        ForbiddenPatterns = @()
        Note = '4-code overflow and unique-exact auto-commit paths exist'
    },
    [PSCustomObject]@{
        Id = 'code_plus_punctuation'
        Title = 'code + punctuation'
        CheckType = 'source-guard'
        Patterns = @(
            'if \(IsPunctuationVirtualKey\(key\) \|\| digitProducesSymbol\)',
            'TryBuildPunctuationCommitText\(',
            'punctuation after composition commit='
        )
        ForbiddenPatterns = @()
        Note = 'punctuation commit-after-composition path exists'
    },
    [PSCustomObject]@{
        Id = 'page_boundary_double_hit'
        Title = 'page boundary double hit'
        CheckType = 'source-guard'
        Patterns = @(
            'pageBoundaryHitCount_ \+= 1;',
            'if \(pageBoundaryHitCount_ < 2\)',
            'keydown-page-boundary code='
        )
        ForbiddenPatterns = @()
        Note = 'second-hit commit path at page boundary exists'
    },
    [PSCustomObject]@{
        Id = 'tab_cross_page'
        Title = 'tab cross page'
        CheckType = 'source-guard'
        Patterns = @(
            'if \(tabNavigation_ && key == VK_TAB && !compositionCode_\.empty\(\)\)',
            'if \(totalPages > 0 && pageIndex_ \+ 1 < totalPages\)',
            'pageIndex_ \+= 1;'
        )
        ForbiddenPatterns = @()
        Note = 'tab and shift+tab cross-page navigation path exists'
    },
    [PSCustomObject]@{
        Id = 'pinyin_query_zhengma'
        Title = 'pinyin query zhengma'
        CheckType = 'source-guard'
        Patterns = @(
            'dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin',
            'auto buildDisplayCode = \[this, pinyinFallbackMode\]',
            'GetSingleCharZhengmaCodeHint\(text\)',
            'codeColumnHeader = dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin \? L"??" : L""'
        )
        ForbiddenPatterns = @()
        Note = 'pinyin zhengma-hint and code-column-header paths exist'
    },
    [PSCustomObject]@{
        Id = 'wildcard_query'
        Title = 'wildcard query'
        CheckType = 'source-guard'
        Patterns = @(
            'yuninput::IsSupportedWildcardCodePattern\(compositionCode_\)',
            'yuninput::ExpandWildcardCodePattern\(compositionCode_, wildcardExpandedCodes\)',
            'kWildcardFastExpandedCodeLimit',
            'emptyHintText =\s*\r?\n\s*pageCandidates\.empty\(\) && hasWildcardChar'
        )
        ForbiddenPatterns = @()
        Note = 'wildcard validate/expand/hint paths exist'
    },
    [PSCustomObject]@{
        Id = 'session_auto_phrase_fastpath'
        Title = 'session auto phrase fast path'
        CheckType = 'source-guard'
        Patterns = @(
            'Session/helper auto phrases must still be merged on fast-path refreshes\.',
            'if \(!pinyinFallbackMode && compositionCode_\.size\(\) >= 4\)'
        )
        ForbiddenPatterns = @(
            '!fastPathSatisfied && !pinyinFallbackMode && compositionCode_\.size\(\) >= 4'
        )
        Note = 'session/helper merge path is not blocked by fastPathSatisfied gate'
    }
)

$outputDir = Split-Path $outputPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$rows = New-Object System.Collections.Generic.List[string]
$rows.Add('case_id`tcase_title`tstatus`tcheck_type`tnote')
$failures = New-Object System.Collections.Generic.List[string]

foreach ($case in $cases) {
    $ok = Test-AllPatterns -content $textServiceContent -patterns $case.Patterns
    if ($ok -and $case.ForbiddenPatterns.Count -gt 0) {
        foreach ($forbiddenPattern in $case.ForbiddenPatterns) {
            if ($textServiceContent -match $forbiddenPattern) {
                $ok = $false
                break
            }
        }
    }

    $status = if ($ok) { 'pass' } else { 'fail' }
    $rows.Add("$($case.Id)`t$($case.Title)`t$status`t$($case.CheckType)`t$($case.Note)")
    if (-not $ok) {
        $failures.Add($case.Id) | Out-Null
    }
}

if (-not $SkipWildcardProbe) {
    $wildcardScriptPath = Resolve-RepoPath $repo 'scripts/run_wildcard_regression.ps1'
    if (Test-Path $wildcardScriptPath) {
        $wildcardOutput = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $wildcardScriptPath 2>&1)
        $wildcardExitCode = $LASTEXITCODE
        $wildcardPassed = $wildcardExitCode -eq 0
        $wildcardStatus = if ($wildcardPassed) { 'pass' } else { 'fail' }
        $wildcardNote = if ($wildcardPassed) {
            'runtime wildcard probe passed'
        } else {
            'runtime wildcard probe failed'
        }

        $rows.Add("wildcard_probe_live`twildcard query (runtime probe)`t$wildcardStatus`truntime-probe`t$wildcardNote")
        if (-not $wildcardPassed) {
            $failures.Add('wildcard_probe_live') | Out-Null
        }

        foreach ($line in $wildcardOutput) {
            Write-Host $line
        }
    } else {
        $rows.Add('wildcard_probe_live`twildcard query (runtime probe)`tfail`truntime-probe`tmissing scripts/run_wildcard_regression.ps1')
        $failures.Add('wildcard_probe_live') | Out-Null
    }
}

Set-Content -Path $outputPath -Value $rows -Encoding UTF8
Write-Host "TextService behavior regression snapshot written: $outputPath"

if ($failures.Count -eq 0) {
    Write-Host 'TextService behavior regression check passed.'
    exit 0
}

Write-Host ("TextService behavior regression check failed: {0} cases." -f $failures.Count)
foreach ($caseId in $failures) {
    Write-Host ("FAIL {0}" -f $caseId)
}
exit 1
