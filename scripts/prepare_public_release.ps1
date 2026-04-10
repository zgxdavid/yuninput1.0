param(
    [string]$VersionTag = '1.0',
    [string]$PublicRoot = '',
    [string]$PrivateArchiveRoot = ''
)

$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot

if ([string]::IsNullOrWhiteSpace($PublicRoot)) {
    $PublicRoot = Join-Path $workspaceRoot ("yuninput_public_" + $VersionTag)
}

if ([string]::IsNullOrWhiteSpace($PrivateArchiveRoot)) {
    $PrivateArchiveRoot = Join-Path $workspaceRoot 'yuninput_private_archive'
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$publicSourceDir = Join-Path $PublicRoot 'source'
$publicArtifactsDir = Join-Path $PublicRoot 'artifacts'
$privateSnapshotDir = Join-Path $PrivateArchiveRoot ("snapshot_" + $stamp)

$msiPath = Join-Path $projectRoot ("Yuninput" + $VersionTag + '.msi')
$sourceListPath = Join-Path $projectRoot ('docs\public-source-file-list-' + $VersionTag + '.txt')

function Remove-IfExists {
    param([string]$Path)

    if (Test-Path $Path) {
        Remove-Item -Recurse -Force $Path
    }
}

function Ensure-Dir {
    param([string]$Path)

    New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

Ensure-Dir $PublicRoot
Ensure-Dir $PrivateArchiveRoot
Ensure-Dir $privateSnapshotDir
Ensure-Dir $publicArtifactsDir

foreach ($privateDirName in @('diagnostics', '_recover_zip')) {
    $src = Join-Path $projectRoot $privateDirName
    if (Test-Path $src) {
        Copy-Item -Recurse -Force $src (Join-Path $privateSnapshotDir $privateDirName)
    }
}

Remove-IfExists $publicSourceDir
Ensure-Dir $publicSourceDir

$excludeDirs = @(
    '.git',
    '.vs',
    'build',
    'dist',
    'diagnostics',
    '_recover_zip',
    '__pycache__'
)

$excludeFiles = @(
    '*.suo',
    '*.user',
    '*.VC.db',
    '*.VC.opendb',
    '*.obj',
    '*.ilk',
    '*.exp',
    '*.lib',
    '*.pdb',
    '*.log',
    '*.dll',
    '*.exe',
    '*.msi',
    '*.wixpdb',
    'data\\yuninput_user.dict',
    'yuninput_user.dict',
    'yuninput_setup.msi',
    'yuninput_setup.wixpdb',
    ('Yuninput' + $VersionTag + '.msi')
)

$robocopyArgs = @(
    $projectRoot,
    $publicSourceDir,
    '/E',
    '/R:1',
    '/W:1',
    '/NFL',
    '/NDL',
    '/NJH',
    '/NJS',
    '/NP'
)

if ($excludeDirs.Count -gt 0) {
    $robocopyArgs += '/XD'
    $robocopyArgs += $excludeDirs
}

if ($excludeFiles.Count -gt 0) {
    $robocopyArgs += '/XF'
    $robocopyArgs += $excludeFiles
}

& robocopy @robocopyArgs | Out-Null
$rc = $LASTEXITCODE
if ($rc -ge 8) {
    throw "Robocopy failed with exit code: $rc"
}

if (Test-Path $msiPath) {
    Copy-Item -Force $msiPath (Join-Path $publicArtifactsDir ("Yuninput" + $VersionTag + '.msi'))
}

$allPublicFiles = Get-ChildItem -Recurse -File $publicSourceDir |
    Sort-Object FullName |
    ForEach-Object {
        $_.FullName.Substring($publicSourceDir.Length).TrimStart('\\') -replace '\\', '/'
    }

Ensure-Dir (Split-Path -Parent $sourceListPath)

$header = @(
    "# Yuninput Public Source List (v$VersionTag)",
    "GeneratedAt: $([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))",
    "SourceRoot: $publicSourceDir",
    "ArtifactRoot: $publicArtifactsDir",
    "PrivateSnapshot: $privateSnapshotDir",
    ''
)

($header + $allPublicFiles) | Set-Content -Encoding utf8 $sourceListPath

$publicSourceListPath = Join-Path $publicSourceDir ('docs\public-source-file-list-' + $VersionTag + '.txt')
Ensure-Dir (Split-Path -Parent $publicSourceListPath)
Copy-Item -Force $sourceListPath $publicSourceListPath

$publicReadme = Join-Path $PublicRoot 'PUBLIC_RELEASE_README.txt'
@(
    "Yuninput public release prepared.",
    "",
    "Public source folder:",
    $publicSourceDir,
    "",
    "Public artifact folder:",
    $publicArtifactsDir,
    "",
    "Private snapshot folder:",
    $privateSnapshotDir,
    "",
    "Source list file:",
    $sourceListPath
) | Set-Content -Encoding utf8 $publicReadme

Get-Item $sourceListPath | Select-Object FullName, Length, LastWriteTime | Format-List
if (Test-Path (Join-Path $publicArtifactsDir ("Yuninput" + $VersionTag + '.msi'))) {
    Get-Item (Join-Path $publicArtifactsDir ("Yuninput" + $VersionTag + '.msi')) | Select-Object FullName, Length, LastWriteTime | Format-List
}