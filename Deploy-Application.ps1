<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableUAC = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableCodeSigningCertificate = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableDisplayingPreviousUsernames = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableInternetExplorerAcademicSettings = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableInternetExplorerFacultyStaffSettings = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableWindowsConsumerFeatures = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableLoginLegalNotice = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableStartMenuLogoffButton = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableVPN = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableSupportInformation = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnableLegacyDataRestore = $false,
	[Parameter(Mandatory=$false)]
	[switch]$EnablePowerSettings = $false,
	[Parameter(Mandatory=$false)]
	[switch]$SetTaskbar = $false,
	[Parameter(Mandatory=$false)]
	[switch]$SetExecutionPolicy = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch { Write-Error -Message "Unable to set the PowerShell Execution Policy to Bypass for this process." }

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'MSU Denver'
	[string]$appName = 'Windows Settings'
	[string]$appVersion = '2.2.0'
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '06/05/2017'
	[string]$appScriptAuthor = 'Jordan Hamilton/Michael Reuther/Quan Tran'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''

	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.9'
	[string]$deployAppScriptDate = '02/12/2017'
	[hashtable]$deployAppScriptParameters = $psBoundParameters

	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Installation tasks here>
			#  Verify that settings were specified on the command-line
			#If ((-not $DisableUAC) -and (-not $DisableDisplayingPreviousUsernames) -and (-not $DisableWindowsConsumerFeatures) -and (-not $EnableInternetExplorerAcademicSettings) -and (-not $EnableInternetExplorerFacultyStaffSettings) -and (-not $EnableLoginLegalNotice) -and (-not $EnableStartMenuLogoffButton) -and (-not $EnableSupportInformation) -and (-not $EnableVPN) -and (-not $EnableLegacyDataRestore) -and (-not $EnablePowerSettings) -and (-not $EnableCodeSigningCertificate) -and (-not $SetTaskbar)) {
				#Show-InstallationPrompt -Message 'No settings were specified' -ButtonRightText 'OK' -Icon 'Error'
				#Exit-Script -ExitCode 9
			#}
			## Enumerate allowed sites for Internet Explorer
			$AllowedSites = @("*.ahec.edu","*.mathxl.com","*.pearsoncmg.com","*.kaltura.com","*.ecollege.com","*.msudenver.edu","*.myitlab.com","*.wimba.com","*.auraria.edu","*.accuplacer.org","*.college-assist.org","*.pearsonmylabandmastering.com","*.coursecompass.com","*.readingplus.com","*.blackboard.com","*.pearsoned.com")
			## Define registry paths
			$DisableDisplayingPreviousUsernamesKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
			$DisableUACKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
			$EnableLoginLegalNoticeKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
			$EnableStartMenuLogoffButtonKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Explorer"
			$EnableSupportInformationKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"

		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		## Handle Zero-Config MSI Installations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
		}

		## <Perform Installation tasks here>
		If ($DisableUAC) {
			Set-RegistryKey -Key $DisableUACKey -Name "EnableLUA" -Value "0" -Type "DWord"
			$mainExitCode = 3010
		}
		If ($DisableDisplayingPreviousUsernames) {
			Set-RegistryKey -Key $DisableDisplayingPreviousUsernamesKey -Name "dontdisplaylastusername" -Value "1" -Type "DWord"
		}
		If ($DisableWindowsConsumerFeatures) {
			Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value "1" -Type "DWord"
		}
		If ($EnableInternetExplorerAcademicSettings) {
			Write-Log -Message "Disabling Internet Explorer first run customization..." -Severity 1 -Source $deployAppScriptFriendlyName
			Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value "1" -Type "DWord"
			[scriptblock]$HKCUFacultyStaffRegistrySettings = {
				Write-Log -Message "Setting Internet Explorer homepage..." -Severity 1 -Source $deployAppScriptFriendlyName
				Set-RegistryKey -Key "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "Start Page" -Value "http://msudenver.edu/studenthub" -Type "String" -SID $UserProfile.SID
				Foreach ($AllowedSite in $AllowedSites)	{
					Write-Log -Message "Allowing Internet Explorer pop-ups for ${AllowedSite}..." -Severity 1 -Source $deployAppScriptFriendlyName
					Set-RegistryKey -Key "HKCU\SOFTWARE\Microsoft\Internet Explorer\New Windows\Allow" -Name "$AllowedSite" -Value (0x00,0x00) -Type "Binary" -SID $UserProfile.SID
				}
			}
			Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCUFacultyStaffRegistrySettings
		}
		If ($EnableInternetExplorerFacultyStaffSettings) {
			Write-Log -Message "Disabling Internet Explorer first run customization..." -Severity 1 -Source $deployAppScriptFriendlyName
			Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value "1" -Type "DWord"
			[scriptblock]$HKCUFacultyStaffRegistrySettings = {
				Write-Log -Message "Setting Internet Explorer homepage..." -Severity 1 -Source $deployAppScriptFriendlyName
				Set-RegistryKey -Key "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "Start Page" -Value "http://msudenver.edu/facstaff" -Type "String" -SID $UserProfile.SID
				Foreach ($AllowedSite in $AllowedSites)	{
					Write-Log -Message "Allowing Internet Explorer pop-ups for ${AllowedSite}..." -Severity 1 -Source $deployAppScriptFriendlyName
					Set-RegistryKey -Key "HKCU\SOFTWARE\Microsoft\Internet Explorer\New Windows\Allow" -Name "$AllowedSite" -Value (0x00,0x00) -Type "Binary" -SID $UserProfile.SID
				}
			}
			Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCUFacultyStaffRegistrySettings
		}
		If ($EnableLoginLegalNotice) {
			Set-RegistryKey -Key $EnableLoginLegalNoticeKey -Name "legalnoticecaption" -Value "Metropolitan State University of Denver" -Type "String"
			Set-RegistryKey -Key $EnableLoginLegalNoticeKey -Name "legalnoticetext" -Value "This workstation is managed by MSU Denver Information Technology Services. By successfully logging into Metropolitan State University of Denver's System, you are agreeing to both the letter and spirit of the MSU Denver Information Security Appropriate Use Policy. Access to this system is restricted to authorized users only. Unauthorized access to this system or misuse of its information may result in disciplinary action including termination of employee and/or student status. Such conduct may also violate criminal laws, which carry severe penalties and are vigorously prosecuted. Information Technology Services routinely monitors system logs and network activity." -Type "String"
		}
		If ($EnableStartMenuLogoffButton) {
			Set-RegistryKey -Key $EnableStartMenuLogoffButtonKey -Name "PowerButtonAction" -Value "1" -Type "DWord"
		}
		If ($EnableSupportInformation) {
			Copy-File -Path "$dirFiles\OEMLogo.bmp" -Destination "$envWinDir" -ContinueOnError $true
			Set-RegistryKey -Key $EnableSupportInformationKey -Name "Logo" -Value "$envWinDir\OEMLogo.bmp" -Type "String"
			Set-RegistryKey -Key $EnableSupportInformationKey -Name "Manufacturer" -Value "Metropolitan State University of Denver" -Type "String"
			Set-RegistryKey -Key $EnableSupportInformationKey -Name "SupportHours" -Value "Monday - Friday, 8am - 5pm | Visit Admin 475 or West 243" -Type "String"
			Set-RegistryKey -Key $EnableSupportInformationKey -Name "SupportURL" -Value "http://msudenver.edu/technology" -Type "String"
			Set-RegistryKey -Key $EnableSupportInformationKey -Name "SupportPhone" -Value "303-352-7548 | 24/7 Phone Support" -Type "String"
			Write-Log -Message "Adding Support Information for ${envOSName} (${envOSVersionMajor}.$envOSVersionMinor)" -Severity 1 -Source $deployAppScriptFriendlyName
			Switch -Wildcard ($envOSVersion) {
				6.1.* {Set-RegistryKey -Key $EnableSupportInformationKey -Name "Model" -Value "Windows 7 ${currentDate}" -Type "String"; break}
				6.2.* {Set-RegistryKey -Key $EnableSupportInformationKey -Name "Model" -Value "Windows 8 ${currentDate}" -Type "String"; break}
				6.3.* {Set-RegistryKey -Key $EnableSupportInformationKey -Name "Model" -Value "Windows 8.1 ${currentDate}" -Type "String"; break}
				10.0.* {Set-RegistryKey -Key $EnableSupportInformationKey -Name "Model" -Value "Windows 10 ${currentDate}" -Type "String"; break}
				default {break}
			}
		}

		If ($EnableVPN) {
			Copy-File -Path "$dirFiles\rasphone.pbk" -Destination "$envProgramData\Microsoft\Network\Connections\Pbk"
		}

		If ($EnableLegacyDataRestore) {
			Write-Log -Message "Creating directory structure for data restoration" -Severity 1 -Source $deployAppScriptFriendlyName
			New-Folder -Path "${envSystemDrive}\Data"
			New-Folder -Path "${envSystemRoot}\MSUDenver"
			Copy-File -Path "${dirFiles}\DataRestore.cmd" -Destination "${envSystemRoot}\MSUDenver"
			[scriptblock]$RunOnce = {
				Write-Log -Message "Setting data Restore to run once on logon..." -Severity 1 -Source $deployAppScriptFriendlyName
				Set-RegistryKey -Key "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "DataRestore" -Value "${envSystemRoot}\MSUDenver\DataRestore.cmd" -Type "String" -SID $UserProfile.SID
			}
			Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $RunOnce
		}

		If ($EnablePowerSettings) {
			$exitCode = Execute-Process -Path "$envSystem32Directory\powercfg.exe" -Parameters "/CHANGE standby-timeout-ac 0" -WindowStyle "Hidden" -PassThru
			If ($exitCode.ExitCode -ne "0") {
				$mainExitCode = $exitCode.ExitCode
			}
		}

		If ($SetTaskbar) {
			#Set local cache for files
			Copy-File -Path "$dirFiles\PinItem\*" -Destination "$envWindir\MSUDenver\PinItem"
			Set-ActiveSetup -StubExePath "$envWindir\system32\cscript.exe" -Arguments "//B //Nologo $envWindir\MSUDenver\PinItem\PinItem.vbs /item:`"$envCommonStartMenuPrograms\Google Chrome.lnk`" /taskbar" -Description "Google Chrome" -Key "Add Chrome" -Version "1"
			Set-ActiveSetup -StubExePath "$envWindir\system32\cscript.exe" -Arguments "//B //Nologo $envWindir\MSUDenver\PinItem\PinItem.vbs /item:`"$envCommonStartMenuPrograms\Mozilla Firefox.lnk`" /taskbar" -Description "Mozilla Firefox" -Key "Add Firefox" -Version "1"
			Set-ActiveSetup -StubExePath "$envWindir\system32\cscript.exe" -Arguments "//B //Nologo $envWindir\MSUDenver\PinItem\PinItem.vbs /item:`"$envCommonStartMenuPrograms\Windows Media Player.lnk`" /unpin /taskbar" -Description "Windows Media Player" -Key "Remove WMP" -Version "1"
		}

		If ($SetExecutionPolicy) {
			## Setting execution policy to RemoteSigned
			Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" -Name "ExecutionPolicy" -Value "RemoteSigned" -Type "String"
			## Importing MSU Denver code signing certificate
			Execute-Process -Path "$envSystem32Directory\certutil.exe" -Parameters "-addstore TrustedPublisher `"$dirFiles\MSUDCodeSigningCertificate.cer`"" -WindowStyle "Hidden"
		}

		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>

		## Display a message at the end of the install
		If (-not $useDefaultMsi) {}
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Uninstallation tasks here>


		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}

		# <Perform Uninstallation tasks here>


		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'

		## <Perform Post-Uninstallation tasks here>


	}

	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}

# SIG # Begin signature block
# MIIU4wYJKoZIhvcNAQcCoIIU1DCCFNACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCB6a+iHjGAqw6d
# 7N+V10u4n+Vy8/2QjReuZWbouyotZqCCD4cwggQUMIIC/KADAgECAgsEAAAAAAEv
# TuFS1zANBgkqhkiG9w0BAQUFADBXMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xv
# YmFsU2lnbiBudi1zYTEQMA4GA1UECxMHUm9vdCBDQTEbMBkGA1UEAxMSR2xvYmFs
# U2lnbiBSb290IENBMB4XDTExMDQxMzEwMDAwMFoXDTI4MDEyODEyMDAwMFowUjEL
# MAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMT
# H0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzIwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQCU72X4tVefoFMNNAbrCR+3Rxhqy/Bb5P8npTTR94ka
# v56xzRJBbmbUgaCFi2RaRi+ZoI13seK8XN0i12pn0LvoynTei08NsFLlkFvrRw7x
# 55+cC5BlPheWMEVybTmhFzbKuaCMG08IGfaBMa1hFqRi5rRAnsP8+5X2+7UulYGY
# 4O/F69gCWXh396rjUmtQkSnF/PfNk2XSYGEi8gb7Mt0WUfoO/Yow8BcJp7vzBK6r
# kOds33qp9O/EYidfb5ltOHSqEYva38cUTOmFsuzCfUomj+dWuqbgz5JTgHT0A+xo
# smC8hCAAgxuh7rR0BcEpjmLQR7H68FPMGPkuO/lwfrQlAgMBAAGjgeUwgeIwDgYD
# VR0PAQH/BAQDAgEGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFEbYPv/c
# 477/g+b0hZuw3WrWFKnBMEcGA1UdIARAMD4wPAYEVR0gADA0MDIGCCsGAQUFBwIB
# FiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAzBgNVHR8E
# LDAqMCigJqAkhiJodHRwOi8vY3JsLmdsb2JhbHNpZ24ubmV0L3Jvb3QuY3JsMB8G
# A1UdIwQYMBaAFGB7ZhpFDZfKiVAvfQTNNKj//P1LMA0GCSqGSIb3DQEBBQUAA4IB
# AQBOXlaQHka02Ukx87sXOSgbwhbd/UHcCQUEm2+yoprWmS5AmQBVteo/pSB204Y0
# 1BfMVTrHgu7vqLq82AafFVDfzRZ7UjoC1xka/a/weFzgS8UY3zokHtqsuKlYBAIH
# MNuwEl7+Mb7wBEj08HD4Ol5Wg889+w289MXtl5251NulJ4TjOJuLpzWGRCCkO22k
# aguhg/0o69rvKPbMiF37CjsAq+Ah6+IvNWwPjjRFl+ui95kzNX7Lmoq7RU3nP5/C
# 2Yr6ZbJux35l/+iS4SwxovewJzZIjyZvO+5Ndh95w+V/ljW8LQ7MAbCOf/9RgICn
# ktSzREZkjIdPFmMHMUtjsN/zMIIEnzCCA4egAwIBAgISESHWmadklz7x+EJ+6RnM
# U0EUMA0GCSqGSIb3DQEBBQUAMFIxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9i
# YWxTaWduIG52LXNhMSgwJgYDVQQDEx9HbG9iYWxTaWduIFRpbWVzdGFtcGluZyBD
# QSAtIEcyMB4XDTE2MDUyNDAwMDAwMFoXDTI3MDYyNDAwMDAwMFowYDELMAkGA1UE
# BhMCU0cxHzAdBgNVBAoTFkdNTyBHbG9iYWxTaWduIFB0ZSBMdGQxMDAuBgNVBAMT
# J0dsb2JhbFNpZ24gVFNBIGZvciBNUyBBdXRoZW50aWNvZGUgLSBHMjCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBALAXrqLTtgQwVh5YD7HtVaTWVMvY9nM6
# 7F1eqyX9NqX6hMNhQMVGtVlSO0KiLl8TYhCpW+Zz1pIlsX0j4wazhzoOQ/DXAIlT
# ohExUihuXUByPPIJd6dJkpfUbJCgdqf9uNyznfIHYCxPWJgAa9MVVOD63f+ALF8Y
# ppj/1KvsoUVZsi5vYl3g2Rmsi1ecqCYr2RelENJHCBpwLDOLf2iAKrWhXWvdjQIC
# KQOqfDe7uylOPVOTs6b6j9JYkxVMuS2rgKOjJfuv9whksHpED1wQ119hN6pOa9PS
# UyWdgnP6LPlysKkZOSpQ+qnQPDrK6Fvv9V9R9PkK2Zc13mqF5iMEQq8CAwEAAaOC
# AV8wggFbMA4GA1UdDwEB/wQEAwIHgDBMBgNVHSAERTBDMEEGCSsGAQQBoDIBHjA0
# MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0
# b3J5LzAJBgNVHRMEAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7
# MDkwN6A1oDOGMWh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vZ3MvZ3N0aW1lc3Rh
# bXBpbmdnMi5jcmwwVAYIKwYBBQUHAQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8v
# c2VjdXJlLmdsb2JhbHNpZ24uY29tL2NhY2VydC9nc3RpbWVzdGFtcGluZ2cyLmNy
# dDAdBgNVHQ4EFgQU1KKESjhaGH+6TzBQvZ3VeofWCfcwHwYDVR0jBBgwFoAURtg+
# /9zjvv+D5vSFm7DdatYUqcEwDQYJKoZIhvcNAQEFBQADggEBAI+pGpFtBKY3IA6D
# lt4j02tuH27dZD1oISK1+Ec2aY7hpUXHJKIitykJzFRarsa8zWOOsz1QSOW0zK7N
# ko2eKIsTShGqvaPv07I2/LShcr9tl2N5jES8cC9+87zdglOrGvbr+hyXvLY3nKQc
# MLyrvC1HNt+SIAPoccZY9nUFmjTwC1lagkQ0qoDkL4T2R12WybbKyp23prrkUNPU
# N7i6IA7Q05IqW8RZu6Ft2zzORJ3BOCqt4429zQl3GhC+ZwoCNmSIubMbJu7nnmDE
# Rqi8YTNsz065nLlq8J83/rU9T5rTTf/eII5Ol6b9nwm8TcoYdsmwTYVQ8oDSHQb1
# WAQHsRgwggbIMIIFsKADAgECAhN/AAAAIhO6jvua86/0AAEAAAAiMA0GCSqGSIb3
# DQEBCwUAMGIxEzARBgoJkiaJk/IsZAEZFgNlZHUxGTAXBgoJkiaJk/IsZAEZFglt
# c3VkZW52ZXIxFTATBgoJkiaJk/IsZAEZFgV3aW5hZDEZMBcGA1UEAxMQd2luYWQt
# Vk1XQ0EwMS1DQTAeFw0xNjA1MjcyMTI0MDJaFw0xODA1MjcyMTI0MDJaMIG/MQsw
# CQYDVQQGEwJVUzERMA8GA1UECBMIQ29sb3JhZG8xDzANBgNVBAcTBkRlbnZlcjEw
# MC4GA1UEChMnTWV0cm9wb2xpdGFuIFN0YXRlIFVuaXZlcnNpdHkgb2YgRGVudmVy
# MSgwJgYDVQQLEx9JbmZvcm1hdGlvbiBUZWNobm9sb2d5IFNlcnZpY2VzMTAwLgYD
# VQQDEydNZXRyb3BvbGl0YW4gU3RhdGUgVW5pdmVyc2l0eSBvZiBEZW52ZXIwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCxCPUOmGXq89WCOBso0z5QIApw
# EosnzQeoI9zP+n8wEb7BEA//+UTmjIZHe3jP0dF6C7EFhx2FcZxs8XQgSH5bnwor
# rkLMa1FzcP2GlcNE5F+ms1zk5Bp2x2nsMOcx+12h9A6eU+JR3nXfWFwkNfvOAKrj
# 1mo4BO5TEvx4DtrVBYFli+0JGnALa1Hd7A68nYtG743FPbioQn8EQSnDr+Jjtd8l
# vujd9I5IQPptiU3inmcoaG+UFz8HKu7QS/mOLpoz/kjbSShxdNF0mcFmowg8WYMu
# f8f1trOtsmWJ3lpyroKek8Ie9oOnKw3And2dOgqWxVXnfLEhW8b6PElvZc73AgMB
# AAGjggMXMIIDEzAOBgNVHQ8BAf8EBAMCBaAwEwYDVR0lBAwwCgYIKwYBBQUHAwEw
# GwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUxu8skV6twX8T
# i5hj8XjbzTUYeqgwHwYDVR0jBBgwFoAUbmigb8ibDuAf063cjbVhC57XDzQwggEo
# BgNVHR8EggEfMIIBGzCCARegggEToIIBD4aBxWxkYXA6Ly8vQ049d2luYWQtVk1X
# Q0EwMS1DQSgxKSxDTj1WTVdDQTAxLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBT
# ZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPXdpbmFkLERD
# PW1zdWRlbnZlcixEQz1lZHU/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNl
# P29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50hkVodHRwOi8vVk1XQ0Ew
# MS53aW5hZC5tc3VkZW52ZXIuZWR1L0NlcnRFbnJvbGwvd2luYWQtVk1XQ0EwMS1D
# QSgxKS5jcmwwggE+BggrBgEFBQcBAQSCATAwggEsMIG6BggrBgEFBQcwAoaBrWxk
# YXA6Ly8vQ049d2luYWQtVk1XQ0EwMS1DQSxDTj1BSUEsQ049UHVibGljJTIwS2V5
# JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz13aW5h
# ZCxEQz1tc3VkZW52ZXIsREM9ZWR1P2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RD
# bGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MG0GCCsGAQUFBzAChmFodHRwOi8v
# Vk1XQ0EwMS53aW5hZC5tc3VkZW52ZXIuZWR1L0NlcnRFbnJvbGwvVk1XQ0EwMS53
# aW5hZC5tc3VkZW52ZXIuZWR1X3dpbmFkLVZNV0NBMDEtQ0EoMSkuY3J0MCEGCSsG
# AQQBgjcUAgQUHhIAVwBlAGIAUwBlAHIAdgBlAHIwDQYJKoZIhvcNAQELBQADggEB
# AIpoMvUtE1iFHSbi7X/M9a+JBPpiAQZzEbq70is1mzdosSVTMN7QoWk4WzHCJBpX
# Oh7cvBrTLf0m4EqJ7OwPY43ZW7MycOjgtk393CaCzr9BiEDjWzJf8r5bDDCodEFm
# dodj3/el8nV4HapjiGnJKrhg0b3xRjPP4cvjtBltbqO7tngkpDu+m63X68aC3wrt
# XwJulfsGeTbd0v4hkji9GCTpLT92mkJyJE04SA/thv4F7yNx1W5XCEWswZeGLiR5
# 9C5AlUm1WrhjAaoyxabDJWfljV//qk+TeoC5CNQ7ZkqdxFBYPc5d2UdkmmiK76D+
# qaobXtlVJ9wRYfFoOaUb5dQxggSyMIIErgIBATB5MGIxEzARBgoJkiaJk/IsZAEZ
# FgNlZHUxGTAXBgoJkiaJk/IsZAEZFgltc3VkZW52ZXIxFTATBgoJkiaJk/IsZAEZ
# FgV3aW5hZDEZMBcGA1UEAxMQd2luYWQtVk1XQ0EwMS1DQQITfwAAACITuo77mvOv
# 9AABAAAAIjANBglghkgBZQMEAgEFAKBmMBgGCisGAQQBgjcCAQwxCjAIoAKAAKEC
# gAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJKoZIhvcNAQkEMSIEIFW9
# riA0DnwlrLAjY4IYNLlbDLyAlDhbDN+MwbX9Os9eMA0GCSqGSIb3DQEBAQUABIIB
# AGUV5ODwBArFcukjzCkDxk66mqSmcQjzGksdqglxlgSzTlNvi9IiCOZV7JcNBnSX
# FDfaTDsVtYZ6Xdz3cKiEzY+4TNUNC2V78YQdeyFKLMu4C21UTNEwafZKxGH91aOZ
# YMxjQjTMw+AlhiDsovFJO6PMSP7fAL2g8D2gRcoqSYYbMPlwTVWnE3ARPgW6N3t4
# kxOBuBbKoDhN5CG2ZKWwHG+rycId0dsxekjY7zH/8uVRIXSjC+5Zsoi493nZk0J8
# EzKfOe1GZFpfYOHJGTYMIR4wlwLvGF98oUG28fAnIFFZljZBE/HpKR41V+mQXzUe
# 1H4Euj9FF8V2tP17hmjTuYChggKiMIICngYJKoZIhvcNAQkGMYICjzCCAosCAQEw
# aDBSMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEoMCYG
# A1UEAxMfR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBHMgISESHWmadklz7x
# +EJ+6RnMU0EUMAkGBSsOAwIaBQCggf0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEH
# ATAcBgkqhkiG9w0BCQUxDxcNMTcwNjAyMDI1MjUzWjAjBgkqhkiG9w0BCQQxFgQU
# xBzkhwi+zD81NVNYQ49s8S7lMOswgZ0GCyqGSIb3DQEJEAIMMYGNMIGKMIGHMIGE
# BBRjuC+rYfWDkJaVBQsAJJxQKTPseTBsMFakVDBSMQswCQYDVQQGEwJCRTEZMBcG
# A1UEChMQR2xvYmFsU2lnbiBudi1zYTEoMCYGA1UEAxMfR2xvYmFsU2lnbiBUaW1l
# c3RhbXBpbmcgQ0EgLSBHMgISESHWmadklz7x+EJ+6RnMU0EUMA0GCSqGSIb3DQEB
# AQUABIIBAHnXmCCR0ZMBghPnPOTmylaVvfoXr5lpYcrCt9KuGiFecMrsRwjoShnr
# xOvsvxxgk+5kAZE/U6PtoX5Bf1WzDSHKoxjGCCRPFxykzCaBbVcQY4IT2VKfC1zS
# Wrkfvwdzcwc4bziTKnTd53crlAZeJK6ZnuETBJlJJa5lfJymvAdLl+Npzf+uUz9z
# YT9T2xbXQ4dYwViX1MMfxmn9VjIyxCkzdZ3zWTmJwHVyH63k4hx947S0LjMAFayB
# XfV+xqN9naQcDTkAp2LMpg/JwjR4dsnhzJAixU31oQyg7imb4SfYW6THvN1tyGai
# sEhvZ0Zi71k2SBinq4Lzhnx5hMjzLm4=
# SIG # End signature block
