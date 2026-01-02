# ---------------------------------------------------------------------------------------------
#   Copyright (c) Microsoft Corporation. All rights reserved.
#   Licensed under the MIT License. See License.txt in the project root for license information.
# ---------------------------------------------------------------------------------------------

# Prevent installing more than once per session
if ((Test-Path variable:global:__VSCodeState) -and $null -ne $Global:__VSCodeState.OriginalPrompt) {
	return;
}

# Disable shell integration when the language mode is restricted
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
	return;
}

$Global:__VSCodeState = @{
	OriginalPrompt = $function:Prompt
	LastHistoryId = -1
	IsInExecution = $false
	EnvVarsToReport = @()
	Nonce = $null
	IsStable = $null
	IsA11yMode = $null
	IsWindows10 = $false
}

# Store the nonce in a regular variable and unset the environment variable. It's by design that
# anything that can execute PowerShell code can read the nonce, as it's basically impossible to hide
# in PowerShell. The most important thing is getting it out of the environment.
$Global:__VSCodeState.Nonce = $env:VSCODE_NONCE
$env:VSCODE_NONCE = $null

$Global:__VSCodeState.IsStable = $env:VSCODE_STABLE
$env:VSCODE_STABLE = $null

$Global:__VSCodeState.IsA11yMode = $env:VSCODE_A11Y_MODE
$env:VSCODE_A11Y_MODE = $null

$__vscode_shell_env_reporting = $env:VSCODE_SHELL_ENV_REPORTING
$env:VSCODE_SHELL_ENV_REPORTING = $null
if ($__vscode_shell_env_reporting) {
	$Global:__VSCodeState.EnvVarsToReport = $__vscode_shell_env_reporting.Split(',')
}
Remove-Variable -Name __vscode_shell_env_reporting -ErrorAction SilentlyContinue

$osVersion = [System.Environment]::OSVersion.Version
$Global:__VSCodeState.IsWindows10 = $IsWindows -and $osVersion.Major -eq 10 -and $osVersion.Minor -eq 0 -and $osVersion.Build -lt 22000
Remove-Variable -Name osVersion -ErrorAction SilentlyContinue

if ($env:VSCODE_ENV_REPLACE) {
	$Split = $env:VSCODE_ENV_REPLACE.Split(":")
	foreach ($Item in $Split) {
		$Inner = $Item.Split('=', 2)
		[Environment]::SetEnvironmentVariable($Inner[0], $Inner[1].Replace('\x3a', ':'))
	}
	$env:VSCODE_ENV_REPLACE = $null
}
if ($env:VSCODE_ENV_PREPEND) {
	$Split = $env:VSCODE_ENV_PREPEND.Split(":")
	foreach ($Item in $Split) {
		$Inner = $Item.Split('=', 2)
		[Environment]::SetEnvironmentVariable($Inner[0], $Inner[1].Replace('\x3a', ':') + [Environment]::GetEnvironmentVariable($Inner[0]))
	}
	$env:VSCODE_ENV_PREPEND = $null
}
if ($env:VSCODE_ENV_APPEND) {
	$Split = $env:VSCODE_ENV_APPEND.Split(":")
	foreach ($Item in $Split) {
		$Inner = $Item.Split('=', 2)
		[Environment]::SetEnvironmentVariable($Inner[0], [Environment]::GetEnvironmentVariable($Inner[0]) + $Inner[1].Replace('\x3a', ':'))
	}
	$env:VSCODE_ENV_APPEND = $null
}

# Register Python shell activate hooks
# Prevent multiple activation with guard
if (-not $env:VSCODE_PYTHON_AUTOACTIVATE_GUARD) {
	$env:VSCODE_PYTHON_AUTOACTIVATE_GUARD = '1'
	if ($env:VSCODE_PYTHON_PWSH_ACTIVATE -and $env:TERM_PROGRAM -eq 'vscode') {
		$activateScript = $env:VSCODE_PYTHON_PWSH_ACTIVATE
		Remove-Item Env:VSCODE_PYTHON_PWSH_ACTIVATE

		try {
			Invoke-Expression $activateScript
			$Global:__VSCodeState.OriginalPrompt = $function:Prompt
		}
		catch {
			$activationError = $_
			Write-Host "`e[0m`e[7m * `e[0;103m VS Code Python powershell activation failed with exit code $($activationError.Exception.Message) `e[0m"
		}
	}
}

function Global:__VSCode-Escape-Value([string]$value) {
	# NOTE: In PowerShell v6.1+, this can be written `$value -replace '…', { … }` instead of `[regex]::Replace`.
	# Replace any non-alphanumeric characters.
	[regex]::Replace($value, "[$([char]0x00)-$([char]0x1f)\\\n;]", { param($match)
			# Encode the (ascii) matches as `\x<hex>`
			-Join (
				[System.Text.Encoding]::UTF8.GetBytes($match.Value) | ForEach-Object { '\x{0:x2}' -f $_ }
			)
		})
}

function Global:Prompt() {
	$FakeCode = [int]!$global:?
	# NOTE: We disable strict mode for the scope of this function because it unhelpfully throws an
	# error when $LastHistoryEntry is null, and is not otherwise useful.
	Set-StrictMode -Off
	$LastHistoryEntry = Get-History -Count 1
	$Result = ""
	# Skip finishing the command if the first command has not yet started or an execution has not
	# yet begun
	if ($Global:__VSCodeState.LastHistoryId -ne -1 -and ($Global:__VSCodeState.HasPSReadLine -eq $false -or $Global:__VSCodeState.IsInExecution -eq $true)) {
		$Global:__VSCodeState.IsInExecution = $false
		if ($LastHistoryEntry.Id -eq $Global:__VSCodeState.LastHistoryId) {
			# Don't provide a command line or exit code if there was no history entry (eg. ctrl+c, enter on no command)
			$Result += "$([char]0x1b)]633;D`a"
		}
		else {
			# Command finished exit code
			# OSC 633 ; D [; <ExitCode>] ST
			$Result += "$([char]0x1b)]633;D;$FakeCode`a"
		}
	}
	# Prompt started
	# OSC 633 ; A ST
	$Result += "$([char]0x1b)]633;A`a"
	# Current working directory
	# OSC 633 ; <Property>=<Value> ST
	$Result += if ($pwd.Provider.Name -eq 'FileSystem') { "$([char]0x1b)]633;P;Cwd=$(__VSCode-Escape-Value $pwd.ProviderPath)`a" }

	# Send current environment variables as JSON
	# OSC 633 ; EnvJson ; <Environment> ; <Nonce>
	if ($Global:__VSCodeState.EnvVarsToReport.Count -gt 0) {
		$envMap = @{}
        foreach ($varName in $Global:__VSCodeState.EnvVarsToReport) {
            if (Test-Path "env:$varName") {
                $envMap[$varName] = (Get-Item "env:$varName").Value
            }
        }
        $envJson = $envMap | ConvertTo-Json -Compress
        $Result += "$([char]0x1b)]633;EnvJson;$(__VSCode-Escape-Value $envJson);$($Global:__VSCodeState.Nonce)`a"
	}

	# Before running the original prompt, put $? back to what it was:
	if ($FakeCode -ne 0) {
		Write-Error "failure" -ea ignore
	}
	# Run the original prompt
	$OriginalPrompt += $Global:__VSCodeState.OriginalPrompt.Invoke()
	$Result += $OriginalPrompt

	# Prompt
	# OSC 633 ; <Property>=<Value> ST
	if ($Global:__VSCodeState.IsStable -eq "0") {
		$Result += "$([char]0x1b)]633;P;Prompt=$(__VSCode-Escape-Value $OriginalPrompt)`a"
	}

	# Write command started
	$Result += "$([char]0x1b)]633;B`a"
	$Global:__VSCodeState.LastHistoryId = $LastHistoryEntry.Id
	return $Result
}

# Report prompt type
if ($env:STARSHIP_SESSION_KEY) {
	[Console]::Write("$([char]0x1b)]633;P;PromptType=starship`a")
}
elseif ($env:POSH_SESSION_ID) {
	[Console]::Write("$([char]0x1b)]633;P;PromptType=oh-my-posh`a")
}
elseif ((Test-Path variable:global:GitPromptSettings) -and $Global:GitPromptSettings) {
	[Console]::Write("$([char]0x1b)]633;P;PromptType=posh-git`a")
}

if ($Global:__VSCodeState.IsA11yMode -eq "1") {
	if (-not (Get-Module -Name PSReadLine)) {
		$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
		$specialPsrlPath = Join-Path $scriptRoot 'psreadline'
		Import-Module $specialPsrlPath
		if (Get-Module -Name PSReadLine) {
			Set-PSReadLineOption -EnableScreenReaderMode
		}
	}
}

# Only send the command executed sequence when PSReadLine is loaded, if not shell integration should
# still work thanks to the command line sequence
$Global:__VSCodeState.HasPSReadLine = $false
if (Get-Module -Name PSReadLine) {
	$Global:__VSCodeState.HasPSReadLine = $true
	[Console]::Write("$([char]0x1b)]633;P;HasRichCommandDetection=True`a")

	$Global:__VSCodeState.OriginalPSConsoleHostReadLine = $function:PSConsoleHostReadLine
	function Global:PSConsoleHostReadLine {
		$CommandLine = $Global:__VSCodeState.OriginalPSConsoleHostReadLine.Invoke()
		$Global:__VSCodeState.IsInExecution = $true

		# Command line
		# OSC 633 ; E [; <CommandLine> [; <Nonce>]] ST
		$Result = "$([char]0x1b)]633;E;"
		$Result += $(__VSCode-Escape-Value $CommandLine)
		# Only send the nonce if the OS is not Windows 10 as it seems to echo to the terminal
		# sometimes
		if ($Global:__VSCodeState.IsWindows10 -eq $false) {
			$Result += ";$($Global:__VSCodeState.Nonce)"
		}
		$Result += "`a"

		# Command executed
		# OSC 633 ; C ST
		$Result += "$([char]0x1b)]633;C`a"

		# Write command executed sequence directly to Console to avoid the new line from Write-Host
		[Console]::Write($Result)

		$CommandLine
	}

	# Set ContinuationPrompt property
	$Global:__VSCodeState.ContinuationPrompt = (Get-PSReadLineOption).ContinuationPrompt
	if ($Global:__VSCodeState.ContinuationPrompt) {
		[Console]::Write("$([char]0x1b)]633;P;ContinuationPrompt=$(__VSCode-Escape-Value $Global:__VSCodeState.ContinuationPrompt)`a")
	}
}

# Set IsWindows property
if ($PSVersionTable.PSVersion -lt "6.0") {
	# Windows PowerShell is only available on Windows
	[Console]::Write("$([char]0x1b)]633;P;IsWindows=$true`a")
}
else {
	[Console]::Write("$([char]0x1b)]633;P;IsWindows=$IsWindows`a")
}

# Set always on key handlers which map to default VS Code keybindings
function Set-MappedKeyHandler {
	param ([string[]] $Chord, [string[]]$Sequence)
	try {
		$Handler = Get-PSReadLineKeyHandler -Chord $Chord | Select-Object -First 1
	}
 catch [System.Management.Automation.ParameterBindingException] {
		# PowerShell 5.1 ships with PSReadLine 2.0.0 which does not have -Chord,
		# so we check what's bound and filter it.
		$Handler = Get-PSReadLineKeyHandler -Bound | Where-Object -FilterScript { $_.Key -eq $Chord } | Select-Object -First 1
	}
	if ($Handler) {
		Set-PSReadLineKeyHandler -Chord $Sequence -Function $Handler.Function
	}
}

function Set-MappedKeyHandlers {
	Set-MappedKeyHandler -Chord Ctrl+Spacebar -Sequence 'F12,a'
	Set-MappedKeyHandler -Chord Alt+Spacebar -Sequence 'F12,b'
	Set-MappedKeyHandler -Chord Shift+Enter -Sequence 'F12,c'
	Set-MappedKeyHandler -Chord Shift+End -Sequence 'F12,d'
}

if ($Global:__VSCodeState.HasPSReadLine) {
	Set-MappedKeyHandlers
}

# SIG # Begin signature block
# MIIu7wYJKoZIhvcNAQcCoIIu4DCCLtwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDkJj4XYymVNoKe
# bz8AekMe0RT8g36wAeqyNJEIBnwuiaCCFBcwggYxMIIEGaADAgECAhMzAAAAOb0d
# 20f6nVYAAAEAAAA5MA0GCSqGSIb3DQEBCwUAMIGHMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMTEwLwYDVQQDEyhNaWNyb3NvZnQgTWFya2V0cGxh
# Y2UgUHJvZHVjdGlvbiBDQSAyMDExMB4XDTI1MDYxOTE4NTQxNFoXDTI2MDYxNzE4
# NTQxNFowdDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEeMBwG
# A1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkqhkiG9w0BAQEFAAOC
# AQ8AMIIBCgKCAQEAtdIdYPZGr5OPIrWkmxTpozsk8RnWxnTYkMnchCH1uxSHAAEG
# RgzhGbD7s32XpJ+XX2hgepvkkQuwVCAWJCw568PSrXoChWYBJVDonwZSphS48Diz
# NuoqhZmp61pqt+p+94Xbli9gtBqk1RBxuEaX6SHbG6JfbIKeOsR6iiMvXsuLoDP4
# G6xf1iBhodE43l7I6mcHnSnmtnrUZCDgrOYN0C4nnf8tExsfcNQ69oX2CyXd3k/2
# f3YFGatRzFjMWY7DKZ5OoUW/igxm7TgspmXgue/JpY1yuO2eOG/Us1LjeNDCa6x9
# UwRMQf3J8X29TQlesnOapdGzP3PRGAGo3/wqlwIDAQABo4IBpjCCAaIwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFBxpSPRkCMknX2x4/vfIokCCS6/5MFQG
# A1UdEQRNMEukSTBHMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRp
# b25zIExpbWl0ZWQxFjAUBgNVBAUTDTIyOTk3OSs1MDUzMDYwHwYDVR0jBBgwFoAU
# nqf5oCNwnxHFaeOhjQr68bD01YAwbAYDVR0fBGUwYzBhoF+gXYZbaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0JTIwTWFya2V0cGxh
# Y2UlMjBQcm9kdWN0aW9uJTIwQ0ElMjAyMDExKDEpLmNybDB5BggrBgEFBQcBAQRt
# MGswaQYIKwYBBQUHMAKGXWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljcm9zb2Z0JTIwTWFya2V0cGxhY2UlMjBQcm9kdWN0aW9uJTIwQ0El
# MjAyMDExKDEpLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBZ
# QQS8MMpdKn3gQxAKmAS6isRYEP2Tz9k/l/zIRoSthAYXhpe5DtlKxvuQQcpEzBuw
# G5M89QaiCZaFVfosu9fXh2R4FK4RCJDwm/zV0PWBY7ZBfjgX2+RljYffHJ5GY7YX
# VV4WIh9hMw220sQ0xg3+BCM9qHjUKx1fDUZINcJLTU/dDJQ1cLKvSYISLA9VlC0F
# dArywD+SCpvL0TUQmRY2kw4VzRx3fP0aPLhbDVcUQc9P9Wuwx/dp+2faji+Ug2aD
# y9CgDHqNGQAP9MhZpQYnr78q/s4BHDxdGmTjfvvgBvIrmIM4nq0F921G5CSFatbi
# srUEN2M+Jk3jMfWciiVTvv2y/wE0MXt+R+hJhyhqlObMJSIr0tFles7fD2sBs+5m
# ccBsMCwt7S+gqtlBr5J/9yn07XSkKJlZSAiDx0MvqWVyBBGvu8QDv4saUjNOG7nH
# JSfpaP66cF3shogzmGad2T+il7p0kmvvTgEX2RQ64x252Hur2DlEmYE1GE5b8bLb
# +7rXwFbsdzR+Gytz/MmKQK5c1LaHJT21guU+AbeRKAPbgs6UajiLBPiGamqhG+Y6
# u8VvH71qKGNBrO+3+Fa3IeGh2R/JdedigfKUQX7NyA8wPlLx3HM4+kQJqZIF2dL1
# N0v4QYELCuZ6a8sxFZFslkLb0zQpVewg91s5rblo6DCCBtcwggS/oAMCAQICCmES
# RKIAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eSAyMDExMB4XDTExMDMyODIxMDkzOVoXDTMxMDMyODIxMTkz
# OVowfTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEnMCUGA1UE
# AxMeTWljcm9zb2Z0IE1hcmtldFBsYWNlIFBDQSAyMDExMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAubUaSwGYVsE3MAnPfvmozUhAB3qxBABgJRW1vDp4
# +tVinXxD32f7k1K89JQ6zDOgS/iDgULC+yFK1K/1Qjac/0M7P6c8v5LSjnWGlERL
# a/qY32j46S7SLQcit3g2jgoTTO03eUG+9yHZUTGV/FJdRYB8uXhrznJBa+Y+yGwi
# QKF+m6XFeBH/KORoKFx+dmMoy9EWJ/m/o9IiUj2kzm9C691+vZ/I2w0Bj93W9SPP
# kV2PCNHlzgfIAoeajWpHmi38Wi3xZHonkzAVBHxPsCBppOoNsWvmAfUM7eBthkSP
# vFruekyDCPNEYhfGqgqtqLkoBebXLZCOVybF7wTQaLvse60//3P003icRcCoQYgY
# 4NAqrF7j80o5U7DkeXxcB0xvengsaKgiAaV1DKkRbpe98wCqr1AASvm5rAJUYMU+
# mXmOieV2EelY2jGrenWe9FQpNXYV1NoWBh0WKoFxttoWYAnF705bIWtSZsz08ZfK
# 6WLX4GXNLcPBlgCzfTm1sdKYASWdBbH2haaNhPapFhQQBJHKwnVW2iXErImhuPi4
# 5W3MVTZ5D9ASshZx69cLYY6xAdIa+89Kf/uRrsGOVZfahDuDw+NI183iAyzC8z/Q
# Rt2P32LYxP0xrCdqVh+DJo2i4NoE8Uk1usCdbVRuBMBQl/AwpOTq7IMvHGElf65C
# qzUCAwEAAaOCAUswggFHMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBQPU8s/
# FmEl/mCJHdO5fOiQrbOU0TAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNV
# HQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQF
# TuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29m
# dC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
# MjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
# MjIuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCjuZmM8ZVNDgp9wHsL4RY8KJ8nLinv
# xFTphNGCrxaLknkYG5pmMhVlX+UB/tSiW8W13W60nggz9u5xwMx7v/1t/Tgm6g2b
# rVyOKI5A7u6/2SIJwkJKFw953K0YIKVT28w9zl8dSJnmRnyR0G86ncWbF6CLQ6A6
# lBQ9o2mTGVqDr4m35WKAnc6YxUUM1y74mbzFFZr63VHsCcOp3pXWnUqAY1rb6Q6N
# X1b3clncKqLFm0EjKHcQ56grTbwuuB7pMdh/IFCJR01MQzQbDtpEisbOeZUi43YV
# AAHKqI1EO9bRwg3frCjwAbml9MmI4utMW94gWFgvrMxIX+n42RBDIjf3Ot3jkT6g
# t3XeTTmO9bptgblZimhERdkFRUFpVtkocJeLoGuuzP93uH/Yp032wzRH+XmMgujf
# Zv+vnfllJqxdowoQLx55FxLLeTeYfwi/xMSjZO2gNven3U/3KeSCd1kUOFS3AOrw
# Z0UNOXJeW5JQC6Vfd1BavFZ6FAta1fMLu3WFvNB+FqeHUaU3ya7rmtxJnzk29DeS
# qXgGNmVSywBS4NajI5jJIKAA6UhNJlsg8CHYwUOKf5ej8OoQCkbadUxXygAfxCfW
# 2YBbujtI+PoyejRFxWUjYFWO5LeTI62UMyqfOEiqugoYjNxmQZla2s4YHVuqIC34
# R85FQlg9pKQBsDCCBwMwggTroAMCAQICEzMAAABVyAZrOCOXKQkAAAAAAFUwDQYJ
# KoZIhvcNAQELBQAwfTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEnMCUGA1UEAxMeTWljcm9zb2Z0IE1hcmtldFBsYWNlIFBDQSAyMDExMB4XDTIx
# MDkwOTIyNDIzMFoXDTMwMDkwOTIyNTIzMFowgYcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMTAvBgNVBAMTKE1pY3Jvc29mdCBNYXJrZXRwbGFj
# ZSBQcm9kdWN0aW9uIENBIDIwMTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDHfQ3P+L0El1S6JNYAz70y3e1i7EZAYcCDVXde/nQdpOKtVr6H4QkBkROv
# 7HBxY0U8lR9C3bUUZKn6CCcN3v3bQuYKu1Ff2G4nIIr8a1cB4iOU8i4YSN7bRr+5
# LvD5hyCfJHqXrJe5LRRGjws5aRAxYuGhQ3ypWPEZYfrIXmmYK+e+udApgxahHUPB
# qcbI2PT1PpkKDgqR7hyzW0CfWzRUwh+YoZpsVvDaEkxcHQe/yGJB5BluYyRm5K9z
# +YQqBvYJkNUisTE/9OImnaZqoujkEuhM5bBV/dNjw7YN37OcBuH0NvlQomLQo+V7
# PA519HVVE1kRQ8pFad6i4YdRWpj/+1yFskRZ5m7y+dEdGyXAiFeIgaM6O1CFrA1L
# bMAvyaZpQwBkrT/etC0hw4BPmW70zSmSubMoHpx/UUTNo3fMUVqx6r2H1xsc4aXT
# pPN5IxjkGIQhPN6h3q5JC+JOPnsfDRg3Ive2Q22jj3tkNiOXrYpmkILk7v+4XUxD
# Erdc/WLZ3sbF27hug7HSVbTCNA46scIqE7ZkgH3M7+8aP3iUBDNcYUWjO1u+P1Q6
# UUzFdShSbGbKf+Z3xpqlwdxQq9kuUahACRQLMFjRUfmAqGXUdMXECRaFPTxl6SB/
# 7IAcuK855beqNPcexVEpkSZxZJbnqjKWbyTk/GA1abW8zgfH2QIDAQABo4IBbzCC
# AWswEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUeBlfau2VIfkw
# k2K+EoAD6hZ05ccwHQYDVR0OBBYEFJ6n+aAjcJ8RxWnjoY0K+vGw9NWAMBkGCSsG
# AQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAG
# AQH/AgEAMB8GA1UdIwQYMBaAFA9Tyz8WYSX+YIkd07l86JCts5TRMFcGA1UdHwRQ
# ME4wTKBKoEiGRmh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL01pY01hclBDQTIwMTFfMjAxMS0wMy0yOC5jcmwwWwYIKwYBBQUHAQEETzBN
# MEsGCCsGAQUFBzAChj9odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRz
# L01pY01hclBDQTIwMTFfMjAxMS0wMy0yOC5jcnQwDQYJKoZIhvcNAQELBQADggIB
# ACY4RaglNFzKOO+3zgazCsgCvXca79D573wDc0DAj6KzBX9m4rHhAZqzBkfSWvan
# LFilDibWmbGUGbkuH0y29NEoLVHfY64PXmXcBWEWd1xK4QxyKx2VVDq9P9494Z/v
# Xy9OsifTP8Gt2UkhftAQMcvKgGiAHtyRHda8r7oU4cc4ITZnMsgXv6GnMDVuIk+C
# q0Eh93rgzKF2rJ1sJcraH/kgSkgawBYYdJlXXHTkOrfEPKU82BDT5h8SGsXVt5L1
# mwRzjVQRLs1FNPkA+Kqyz0L+UEXJZWldNtHC79XtYh/ysRov4Yu/wLF+c8Pm15IC
# n8EYJUL4ZKmk9ZM7ZcaUV/2XvBpufWE2rcMnS/dPHWIojQ1FTToqM+Ag2jZZ33fl
# 8rJwnnIF/Ku4OZEN24wQLYsOMHh6WKADxkXJhiYUwBe2vCMHDVLpbCY7CbPpQdtB
# YHEkto0MFADdyX50sNVgTKboPyCxPW6GLiR5R+qqzNRzpYru2pTsM6EodSTgcMbe
# aDZI7ssnv+NYMyWstE1IXQCUywLQohNDo6H7/HNwC8HtdsGd5j0j+WOIEO5PyCbj
# n5viNWWCUu7Ko6Qx68NuxHf++swe9YQhufh0hzJnixidTRPkBUgYQ6xubG6I5g/2
# OO1BByOu9/jt5vMTTvctq2YWOhUjoOZPe53eYSzjvNydMYIaLjCCGioCAQEwgZ8w
# gYcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMTAvBgNVBAMT
# KE1pY3Jvc29mdCBNYXJrZXRwbGFjZSBQcm9kdWN0aW9uIENBIDIwMTECEzMAAAA5
# vR3bR/qdVgAAAQAAADkwDQYJYIZIAWUDBAIBBQCggbAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIKLfJwfcCS1FOTtHeWToQOlMlH1iAdZOhG3Hy+B38Hs7MEQGCisG
# AQQBgjcCAQwxNjA0oBCADgBWAFMAIABDAG8AZABloSCAHmh0dHBzOi8vY29kZS52
# aXN1YWxzdHVkaW8uY29tLzANBgkqhkiG9w0BAQEFAASCAQCJ/6Z/Z57ixRtbtYug
# 0BIVLuGGR3qUKmqSLF3CuWZLU0dYPpmoZEZwkC34LUxco8ZIAWBjolUAowJ/oDc9
# 6LQ+hRLKf7XdWlni4pkdRJJuEVKLo8ARkymnGgwtOqH5nqlDvpnQSXGnOH9We9H1
# l/1ShFhCidtXEQ3qjKUX56OpJFXcpsiZM3BAPpLhNKtTs32rGBsiCSS/2iB+uup/
# LI7RSEJRbOvjrxYOlloG0Gpqee2JeoZf4bI/6TP84FdGXivfsgljUW+8nbjHI6IP
# 7qBVa0YJdindW90+bkuDGgfL30GoC+Ijf07rccCg+x52hmByjw9/3ru2qwXyHd1W
# t1OloYIXrDCCF6gGCisGAQQBgjcDAwExgheYMIIXlAYJKoZIhvcNAQcCoIIXhTCC
# F4ECAQMxDzANBglghkgBZQMEAgEFADCCAVoGCyqGSIb3DQEJEAEEoIIBSQSCAUUw
# ggFBAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIByZew7Z3A3IDQyV
# c8XvvFqIZrvCe4aHB9pEfoYInDQHAgZpQpfLDgEYEzIwMjUxMjE3MTQ0MjA5LjI1
# OFowBIACAfSggdmkgdYwgdMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGlt
# aXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjZCMDUtMDVFMC1EOTQ3MSUw
# IwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIR+jCCBygwggUQ
# oAMCAQICEzMAAAIRRRg5m0PP/GwAAQAAAhEwDQYJKoZIhvcNAQELBQAwfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjUwODE0MTg0ODEzWhcNMjYxMTEz
# MTg0ODEzWjCB0zELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEt
# MCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMScw
# JQYDVQQLEx5uU2hpZWxkIFRTUyBFU046NkIwNS0wNUUwLUQ5NDcxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQDPubszEB0vlVrKuTuOwyjcaeE3zmS0cJkS8RyPgEhxwcp3
# 80oLu4++lfl2E7rdbpUzmILGSUbypB5VWs9oq+Px1hgkLsM23g03deVV0L++i94m
# 48+FMn+7tf6liZXap6FNU844HX+Gma3nVLODFlzMx2cWX5fZ7U+C61IDkICH39fP
# k1bQLGdhXPyDRWnGD4GrfZqaS1FevybcFISBSzyOBZE9XM8cRzOluGWgYYR8dpE6
# YeFUoio34mEzB4SNTY1czZbqGbfaP9Af8j8pao019hyEdobTEmWNVNihQo+lxAO6
# Ef11AoSC8bGPZTn/cWrV6bh07oiHTibpH623GvpjyhEkf1mFnexyIUEi9mHsTZgV
# c6M/gwbJtLKVBM8MQUC0ceCmSyR4RSGw8NH1W9ZaF6SFDHepdoAqH4CQubP+GkTd
# 7TL5Ego7YBESNQskAqB/5H1Cc2+ox4yTP08auOyKOpYbMHaTYk3JpRgqVuZDB45p
# uwKKiJjZ8luKaNXIUAaTkB5h11QXG8kaBFUIfsF4E8oCrsww6ZIJM4xnRLDrPI3H
# hSGHljS4nRk6hMqcHcp9039tr94ocV4SGLdaoB/NPGLLSsy+Gx+xdkrvOhyWppG9
# WXxDjwnXvj57KuLKlj0eFT6iGCJiLi5AYMNV1MN4oO2gL+EPYKf4BHPATWsV8QID
# AQABo4IBSTCCAUUwHQYDVR0OBBYEFGJ9RQPA6eohy99vnf7JXQRmfs5wMB8GA1Ud
# IwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCGTmh0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMFRp
# bWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4wXAYI
# KwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMv
# TWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeA
# MA0GCSqGSIb3DQEBCwUAA4ICAQCkQp2cx4ghSJTo9q1n+puvCIPNhQwpFzMLgGn9
# djVL02Ycj7Zzd1ynAfZI6YN928giq3uZGuC8E9g68n0K1lLl54iuw5sLRvSCApO/
# bCtOBYb6qS2o0USFB6Kl1RE0s3ry4cCbl53AHK13WTDLmvoH3eSXEOyV06ZVa3D+
# eCPuSc3T2a4KbCvXsmewwVygg38fn2z7VFg3tWJ3j7uePwVy9jL2ttk4yd0HOxOK
# iwXUz5owglfaTcRUVWy4Mvv9Hmmkj1ODt5ZA5Yoxkc92wDdmpbMO6EmpPOgVJBKG
# dl6cL7Gr/P0GEc8UVtS1+MCgboQM+NJAlheaiCNrw4RrX3HCeHfBW594/5yT7/SD
# E2LuD6Q7pZo6bTnYXiyIPzGLpS/vkvvv3yUe89OFzEceyBeoxjn3Z3XBSh/e0v94
# NpDRSGdgJTzIaRTZcmdy042cEoC9REC9/aqIhYOPgulybTMDtW6h+4lHVOm7Jzmn
# WNrnZs1kEFWoA7DIOECapawlcCNheeywL98mR57fXgWH4YjIyC8A9FJyCpFmpXXp
# 1MFi+h77DWf/Baz/JJNSzEPDhP8AhNy7k8CwucJWkCsOsUtFMXK6354dSgbpRhl+
# Pz9Gy5DjYg2x7Wlv9w+bsbaVwsm2QgpPzTG8HUuJo289MFURyY1K8VQzTGtdldxh
# zFVeJjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZIhvcN
# AQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAw
# BgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEw
# MB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDk
# 4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25PhdgM/9c
# T8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPFdvWG
# UNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6Gnsz
# rYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBpDco2
# LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50ZuyjLV
# wIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3EXzTd
# EonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0
# gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1qGFph
# AXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ+QuJ
# YfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PAPBXb
# GjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkwEgYJ
# KwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxGNSnP
# EP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARVMFMw
# UQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAKBggr
# BgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYw
# DwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoY
# xDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYB
# BQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
# cGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG9w0B
# AQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0xM7U5
# 18JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmCVgAD
# sAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449xvNo
# 32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wMnosZ
# iefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDSPeZK
# PmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2dY3RI
# LLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxnGSgk
# ujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9
# af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokLjzba
# ukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL6Xu/
# OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNVMIIC
# PQIBATCCAQGhgdmkgdYwgdMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGlt
# aXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjZCMDUtMDVFMC1EOTQ3MSUw
# IwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4D
# AhoDFQArKnyrZV2ACrVUaTN3s9nBXrM1zaCBgzCBgKR+MHwxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA7O0WSzAiGA8yMDI1MTIx
# NzExNDUxNVoYDzIwMjUxMjE4MTE0NTE1WjBzMDkGCisGAQQBhFkKBAExKzApMAoC
# BQDs7RZLAgEAMAYCAQACATgwBwIBAAICFQwwCgIFAOzuZ8sCAQAwNgYKKwYBBAGE
# WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDAN
# BgkqhkiG9w0BAQsFAAOCAQEAUSSg/ljYwajpdv5SmE+fIhfA1nT0ploBdMDGAUYh
# w+4RNkzhj0DZV2z7v1yjpUJ16wVOSnEHOzqsgtWqEFF8FqFyPEyLftPGlhvXoIaN
# o3iHnlknf8FRs0ZCkHICzRpnsy3s3VNbCSyAIwxvM0y1qzH2rA96Pz8y4h9Rxfef
# V0+mhSLLgaQ/no8U3hY7qz3EBm8NgtKsz1gr1M8OL25EEn/lI3sLmMo4o7ZHwRs5
# /8h8N5s/qaBqnxsDHHLdSK53LlQaDWTuM4sE/abIzy5mGpwG3/VMwAEYEVVuWPz6
# d5snP+zp2Kv5Hob3uQOQJcbHIEJjdGwImzWQr0oYmoiQ2TGCBA0wggQJAgEBMIGT
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACEUUYOZtDz/xsAAEA
# AAIRMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQ
# AQQwLwYJKoZIhvcNAQkEMSIEIA9dhgJa8yVMIfCL0M4AytmaQioP1nvRik8DAjEy
# xFhfMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgLK0zqZrvh06tWlxcL5YY
# xfKdp1AjTQhF/zlixzQzJrcwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAhFFGDmbQ8/8bAABAAACETAiBCDCp4Ox30qYftlcDtUS9M4+
# ReEEO1Rf/ExzqJCJ3STXZzANBgkqhkiG9w0BAQsFAASCAgAMiWydIE8gD0iZc+tB
# Hyxfe8nCEo0FlS/Rh2j8JZdgd/h1ihe4emqbs4nS56RFYRH5foIJTBGeJ+ojUwMN
# D2HOC62eIOQ905bjtTXELujMItAwfyU9hs0Ww8UtKVmbultre6KOXjLWwMDsYtFY
# httaulKOxZ9EXR3THsF/LC4/6zuvlHVO1oDIFb2KwzDJRdGWmEyzJu4aTK3Co2r7
# w5K8OY9cOFJ5Tstj8+u7ikrV7NnntQxLI6gk6RzRpBGCs+8AqILdAj9oia4OqTk9
# HD8PyZ/o/+lt4+6x2jlTe+pwtEG52sVWlQhHVrS1obAuY5AdRYH4vwBclFuo08LX
# kpLzoL3u0hjMBBsL1rTwTGxmCSRgFPzHMRKeR+kwGYecKpKxLMzYoB4EPMafjVSW
# iY5Hz91DX3ZcbC0jDBYCGdtmHIDFvIwOnuDc6ZMq0iiACvcOSoCnQEIw2+GgBLW8
# GK74ZJFOO21GONM6mqiAv7pf/1sqM+63Ks7RIGJvPrVR4P6/CqZ2vQ/UQfmWM538
# auUA6u6lJNiBpaTcLc11bCUsDo/qjarwWknuMkHFXptQKP6ffdns6H5R8Kfcmqir
# QSAtnSDt3slXLocC2/P+dgYQSRlGhI7jbPS2uO9+fXg01luEeVq51QMKL2VFa74d
# lYD0vmCI9Yh9UUc1PfjXfl6sMg==
# SIG # End signature block
