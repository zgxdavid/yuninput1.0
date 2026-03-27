param(
	[string]$InstallRoot = "$env:LOCALAPPDATA\\yuninput"
)

$ErrorActionPreference = "Stop"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
	$argumentList = @(
		'-NoProfile',
		'-ExecutionPolicy',
		'Bypass',
		'-File',
		$PSCommandPath
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
		Write-Error "Elevation failed or was canceled: $($_.Exception.Message)"
		exit 1
	}
}

$unregisterScript = Join-Path $PSScriptRoot 'unregister_ime.ps1'
$dllPath = Join-Path $InstallRoot 'bin\yuninput.dll'
$clsid = '{6DE9AB40-3BA8-4B77-8D8F-233966E1C102}'

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

	Set-RegistryAdminAclRecursively -RegistryPath $RegistryPath

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

if (Test-Path $dllPath) {
	try {
		& $unregisterScript -DllPath $dllPath
	}
	catch {
		Write-Warning "Unregister script returned error: $($_.Exception.Message)"
	}
}

# Explicit registry cleanup for stale TIP/CLSID residues.
$registryTargets = @(
	'Registry::HKCU:\Software\Microsoft\CTF\TIP\' + $clsid,
	'Registry::HKLM:\Software\Microsoft\CTF\TIP\' + $clsid,
	'Registry::HKLM:\Software\Classes\CLSID\' + $clsid
)

foreach ($path in $registryTargets) {
	Remove-RegistryKeyForce -RegistryPath $path -RegExePath $path
}

# Restart text service process so keyboard list refreshes immediately.
Get-Process -Name ctfmon -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Process ctfmon.exe | Out-Null

if (Test-Path $InstallRoot) {
	Remove-Item -Recurse -Force $InstallRoot
}

Write-Host ''
Write-Host 'Uninstall completed. If keyboard entry still exists, remove it from language settings.'
Start-Process 'ms-settings:regionlanguage'
