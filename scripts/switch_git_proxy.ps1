param(
    [ValidateSet('AutoPac', 'Set', 'Direct')]
    [string]$Mode = 'AutoPac',
    [string]$Proxy,
    [ValidateSet('local', 'global')]
    [string]$Scope = 'local',
    [switch]$Test
)

$ErrorActionPreference = 'Stop'

function Get-PacUrl {
    $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
    $item = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    return $item.AutoConfigURL
}

function Get-ProxyFromPac {
    param([Parameter(Mandatory = $true)][string]$PacUrl)

    $pac = curl.exe -s --max-time 20 $PacUrl
    if ([string]::IsNullOrWhiteSpace($pac)) {
        throw "PAC download failed: $PacUrl"
    }

    $match = [regex]::Match($pac, 'PROXY\s+([A-Za-z0-9._-]+:\d+)', 'IgnoreCase')
    if (-not $match.Success) {
        throw 'No PROXY host:port found in PAC file.'
    }

    return $match.Groups[1].Value
}

function Set-GitProxy {
    param(
        [Parameter(Mandatory = $true)][string]$ProxyValue,
        [Parameter(Mandatory = $true)][string]$TargetScope
    )

    git config --$TargetScope http.proxy "http://$ProxyValue"
    git config --$TargetScope https.proxy "http://$ProxyValue"
}

function Clear-GitProxy {
    param([Parameter(Mandatory = $true)][string]$TargetScope)

    git config --$TargetScope --unset-all http.proxy 2>$null
    git config --$TargetScope --unset-all https.proxy 2>$null
}

if ($Mode -eq 'AutoPac') {
    $pacUrl = Get-PacUrl
    if ([string]::IsNullOrWhiteSpace($pacUrl)) {
        throw 'AutoConfigURL is empty. Use -Mode Set -Proxy <host:port> instead.'
    }
    $resolvedProxy = Get-ProxyFromPac -PacUrl $pacUrl
    Set-GitProxy -ProxyValue $resolvedProxy -TargetScope $Scope
    Write-Host "Git proxy set from PAC: $resolvedProxy ($Scope)"
}
elseif ($Mode -eq 'Set') {
    if ([string]::IsNullOrWhiteSpace($Proxy)) {
        throw 'Use -Proxy <host:port> with -Mode Set.'
    }
    Set-GitProxy -ProxyValue $Proxy -TargetScope $Scope
    Write-Host "Git proxy set: $Proxy ($Scope)"
}
else {
    Clear-GitProxy -TargetScope $Scope
    Write-Host "Git proxy cleared ($Scope)"
}

if ($Test) {
    Write-Host 'Testing remote access...'
    git ls-remote https://github.com/zgxdavid/yuninput1.0.git HEAD
}

Write-Host 'Current proxy config:'
git config --$Scope --get-regexp '^http\.proxy|^https\.proxy'