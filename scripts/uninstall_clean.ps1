param(
	[string]$InstallRoot = "$env:LOCALAPPDATA\\yuninput",
	[switch]$SkipElevation,
	[switch]$ForceElevation
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$inVsCodeHost = -not [string]::IsNullOrWhiteSpace($env:VSCODE_PID)
if (-not $isAdmin -and -not $SkipElevation -and $inVsCodeHost) {
	$SkipElevation = $true
	Write-Warning 'Detected non-admin VS Code terminal. Auto-enabling -SkipElevation to avoid hidden UAC prompt blocking.'
}

if (-not $isAdmin -and -not $SkipElevation -and -not $ForceElevation) {
	$SkipElevation = $true
	Write-Warning 'Non-admin session detected. Auto-enabling -SkipElevation to avoid hidden UAC prompt blocking.'
}

if (-not $isAdmin -and -not $SkipElevation) {
	$argumentList = @(
		'-NoProfile',
		'-ExecutionPolicy',
		'Bypass',
		'-File',
		$PSCommandPath,
		'-SkipElevation'
	)

	if ($InstallRoot -ne "$env:LOCALAPPDATA\\yuninput") {
		$argumentList += '-InstallRoot'
		$argumentList += $InstallRoot
	}

	Write-Host 'Requesting administrator permission...'
	try {
		$proc = Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argumentList -Wait -PassThru
		exit $proc.ExitCode
	}
	catch {
		Write-Warning "Elevation failed or was canceled. Continuing with current-user cleanup only: $($_.Exception.Message)"
	}
}

$unregisterScript = Join-Path $PSScriptRoot 'unregister_ime.ps1'
$dllPath = Join-Path $InstallRoot 'bin\yuninput.dll'
$clsid = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'
$profileGuid = '{47DE2FB1-F5E4-4CF8-AB2F-8F7A761731B2}'
$tipToken = '0804:' + $clsid + $profileGuid

function Stop-InputProcesses {
	param([int]$RetryCount = 3)

	$processNames = @('ctfmon', 'TextInputHost')
	for ($i = 0; $i -lt $RetryCount; $i++) {
		foreach ($name in $processNames) {
			Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
		}

		Start-Sleep -Milliseconds 250
	}
}

function Remove-CurrentUserTipToken {
	param([string]$Token)

	try {
		$languageList = Get-WinUserLanguageList
		$updated = $false
		foreach ($language in $languageList) {
			if ($null -eq $language.InputMethodTips) {
				continue
			}

			$keep = @()
			foreach ($tip in $language.InputMethodTips) {
				if (-not [string]::Equals([string]$tip, $Token, [System.StringComparison]::OrdinalIgnoreCase)) {
					$keep += [string]$tip
				}
			}

			if ($keep.Count -ne $language.InputMethodTips.Count) {
				$language.InputMethodTips.Clear()
				foreach ($tip in $keep) {
					[void]$language.InputMethodTips.Add($tip)
				}
				$updated = $true
			}
		}

		if ($updated) {
			Set-WinUserLanguageList -LanguageList $languageList -Force
		}
	}
	catch {
		Write-Warning "Failed to update current user language list: $($_.Exception.Message)"
	}
}

function Set-RegistryAdminAclRecursively {
	param(
		[Parameter(Mandatory = $true)]
		[string]$RegistryPath
	)

	if (-not (Test-Path $RegistryPath)) {
		return
	}

	$allTargets = @()
	$allTargets += $RegistryPath
	$children = Get-ChildItem -Path $RegistryPath -Recurse -ErrorAction SilentlyContinue |
		Sort-Object { $_.Name.Length } -Descending
	$allTargets += $children.PSPath

	foreach ($target in $allTargets) {
		try {
			$acl = Get-Acl -Path $target
			$owner = [System.Security.Principal.NTAccount]::new('Administrators')
			$acl.SetOwner($owner)
			Set-Acl -Path $target -AclObject $acl

			$updatedAcl = Get-Acl -Path $target
			$rule = [System.Security.AccessControl.RegistryAccessRule]::new(
				'Administrators',
				[System.Security.AccessControl.RegistryRights]::FullControl,
				[System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
				[System.Security.AccessControl.PropagationFlags]::None,
				[System.Security.AccessControl.AccessControlType]::Allow
			)
			$updatedAcl.SetAccessRule($rule)
			Set-Acl -Path $target -AclObject $updatedAcl
		}
		catch {
			Write-Warning "Failed to update ACL on ${target}: $($_.Exception.Message)"
		}
	}
}

function Remove-RegistryKeyForce {
	param(
		[Parameter(Mandatory = $true)]
		[string]$RegistryPath,
		[Parameter(Mandatory = $true)]
		[string]$RegExePath
	)

	if (-not (Test-Path $RegistryPath)) {
		return
	}

	if ($isAdmin -and $RegistryPath -like 'Registry::HKLM:*') {
		Set-RegistryAdminAclRecursively -RegistryPath $RegistryPath
	}

	try {
		Remove-Item -Path $RegistryPath -Recurse -Force -ErrorAction Stop
		return
	}
	catch {
		Write-Warning "Remove-Item failed for ${RegistryPath}: $($_.Exception.Message)"
	}

	$regNative = $RegExePath.Replace('Registry::', '').Replace('HKCU:', 'HKCU').Replace('HKLM:', 'HKLM')
	& reg.exe delete $regNative /f | Out-Null
}

Stop-InputProcesses

if (Test-Path $dllPath) {
	try {
		& $unregisterScript -DllPath $dllPath
	}
	catch {
		Write-Warning "Unregister script returned error: $($_.Exception.Message)"
	}
}

Get-ChildItem -Path (Join-Path $InstallRoot 'bin') -Filter 'yuninput_*.dll' -File -ErrorAction SilentlyContinue | ForEach-Object {
	try {
		& $unregisterScript -DllPath $_.FullName
	}
	catch {
	}
}

Remove-CurrentUserTipToken -Token $tipToken

# Explicit registry cleanup for stale TIP/CLSID residues.
$registryTargets = @(
	'Registry::HKCU:\Software\Microsoft\CTF\TIP\' + $clsid,
	'Registry::HKCU:\Software\Classes\CLSID\' + $clsid
)

if ($isAdmin) {
	$registryTargets += @(
		'Registry::HKLM:\Software\Microsoft\CTF\TIP\' + $clsid,
		'Registry::HKLM:\Software\Classes\CLSID\' + $clsid
	)
}

foreach ($path in $registryTargets) {
	Remove-RegistryKeyForce -RegistryPath $path -RegExePath $path
}

# Restart text service process so keyboard list refreshes immediately.
Stop-InputProcesses
Start-Process ctfmon.exe | Out-Null

if (Test-Path $InstallRoot) {
	Get-ChildItem -Path $InstallRoot -Recurse -Force -ErrorAction SilentlyContinue |
		Sort-Object FullName.Length -Descending |
		ForEach-Object {
			$itemPath = $_.FullName
			try {
				Remove-Item -LiteralPath $itemPath -Recurse -Force -ErrorAction Stop
			}
			catch {
				Write-Warning "Failed to remove ${itemPath}: $($_.Exception.Message)"
			}
		}

	try {
		Remove-Item -LiteralPath $InstallRoot -Recurse -Force -ErrorAction Stop
	}
	catch {
		Write-Warning "Install root not fully removed: $($_.Exception.Message)"
	}
}

Write-Host ''
if ($isAdmin) {
	Write-Host 'Uninstall completed. If keyboard entry still exists, remove it from language settings.'
}
else {
	Write-Host 'Current-user cleanup completed. Machine-wide registry entries, if any, still require an elevated uninstall.'
}
Start-Process 'ms-settings:regionlanguage'
