<#
------------------------------------------------------------------
AUTHOR:
Brian Gade

DESCRIPTION:
Installs Office 2016

PARAMETERS:  
FullSuite - switch
ExcelPowerPointOnly - switch
UninstallOffice - switch

SAMPLE SYNTAX:
.\Install_Office2016.ps1 -FullSuite

CHANGE LOG:
v1.0.0, Brian Gade, 4/2/18 - Original Version
------------------------------------------------------------------
#>
<#

.SYNOPSIS
Author: Brian Gade

.DESCRIPTION
Removes Office 2013, Office 2016, and Office 365.  Then installs Office 2016.

.PARAMETER Install
Select this parameter to install this software

.PARAMETER Uninstall
Select this parameter to uninstall this software

.EXAMPLE
.\Install_Office2016.ps1 -Install

#>
# DEFINE PARAMETERS ----------------------------------------------
Param(
	[Parameter(ParameterSetName='Install')]
		[switch] $Install,
	[Parameter(ParameterSetName='Uninstall')]
		[switch] $Uninstall
)
# END DEFINE PARAMETERS ------------------------------------------
# DEFINE FUNCTIONS -----------------------------------------------

Function RemoveOffice2013
{
	$UninstallParentPath = "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller"
	$UninstallPath = "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\setup.exe"
	$CD = (Get-Location).Path
	If((Test-Path $UninstallPath) -and ((Get-ChildItem $UninstallPath).VersionInfo.ProductVersion -LIKE '15.*'))
	{
		# Remove Office 2013 Portuguese Language Pack
		Write-Host "Removing Office 2013 Portuguese Language Pack" -ForegroundColor Green
		cmd /c msiexec /x '{90150000-001F-0416-0000-0000000FF1CE}' /quiet /norestart
		
		# Remove Office 2013 Korean Language Pack
		Write-Host "Removing Office 2013 Korean Language Pack" -ForegroundColor Green
		cmd /c msiexec /x '{90150000-001F-0412-0000-0000000FF1CE}' /quiet /norestart
		
		# Remove Office 2013 Japanese Language Pack
		Write-Host "Removing Office 2013 Japanese Language Pack" -ForegroundColor Green
		cmd /c msiexec /x '{90150000-001F-0411-0000-0000000FF1CE}' /quiet /norestart
		
		# Remove Office 2013 Italian Language Pack
		Write-Host "Removing Office 2013 Italian Language Pack" -ForegroundColor Green
		cmd /c msiexec /x '{90150000-001F-0410-0000-0000000FF1CE}' /quiet /norestart
		
		# Remove Office 2013 German Language Pack
		Write-Host "Removing Office 2013 German Language Pack" -ForegroundColor Green
		cmd /c msiexec /x '{90150000-001F-0407-0000-0000000FF1CE}' /quiet /norestart
		
		# Remove Office 2013 Chinese Language Pack
		Write-Host "Removing Office 2013 Chinese Language Pack" -ForegroundColor Green
		cmd /c msiexec /x '{90150000-001F-0804-0000-0000000FF1CE}' /quiet /norestart
		
		# Remove Office 2013
		Write-Host "Removing Office 2013 ProPlus" -ForegroundColor Green
		copy .\Uninstall_Office2013ProPlus.xml $UninstallParentPath\
		Start-Process -FilePath $UninstallPath -ArgumentList "/uninstall ProPlus /config .\Uninstall_Office2013ProPlus.xml" -Wait
	}
}

# END DEFINE FUNCTIONS -------------------------------------------
# START TRANSCRIPTING --------------------------------------------
$TranscriptLogFile = '.\Install_Office2016.log'
Start-Transcript -Path $TranscriptLogFile -Append -NoClobber
# SCRIPT BODY ----------------------------------------------------

# If FullSuite is selected, install the full Office 2016 suite
If($Install)
{
	Write-Host "Removing Office 2013" -ForegroundColor Green
	RemoveOffice2013
	Write-Host "Removing Office 2016" -ForegroundColor Green
	cmd /c setup.exe /uninstall ProPlus /config .\Uninstall_Office2016ProPlus.xml
	Write-Host "Removing Office 365" -ForegroundColor Green
	cmd /c setup_365.exe /configure .\Uninstall_O365.xml
	Write-Host "Installing Office 2016" -ForegroundColor Green
	cmd /c setup.exe /adminfile .\FullSuite.msp
}
ElseIf($Uninstall)
{
	Write-Host "Removing Office 2013" -ForegroundColor Green
	RemoveOffice2013
	Write-Host "Removing Office 2016" -ForegroundColor Green
	cmd /c setup.exe /uninstall ProPlus /config .\Uninstall_Office2016ProPlus.xml
	Write-Host "Removing Office 365" -ForegroundColor Green
	cmd /c setup_365.exe /configure .\Uninstall_O365.xml
}

# END SCRIPT BODY ------------------------------------------------
# STOP TRANSCRIPTING ---------------------------------------------
Stop-Transcript
# END SCRIPT -----------------------------------------------------