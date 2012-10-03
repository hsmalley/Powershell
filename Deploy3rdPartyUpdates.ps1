<# 
	.Synopsis 
   		Update 3rd party software on Network computers
	.Description 
		This script deploy's MSI & MSP Updates to computers in a list pulled from AD. Any systems that are offline at the time will be placed in $Skipped.
	.Notes
    	NAME: Deploy3rdPartyUpdates.ps1
    	AUTHOR: Hugh Smalley
	.Errors
		* If running script over again in the same session you will get the following error. "Add-PSSnapin : Cannot add Windows PowerShell snap-in Quest.ActiveRoles.ADManagement because it is already added. Verify the name of the snap-in and try again."
		* ReturnValue : 2 - Means locked files
		* HRESULT: 0x800706BA - Means a connection can not be made to the system, e.g. system is shutting down or account doesn't have access to the system.
#> 
<# ::: Updates ::: #>
$FlashUpdateAX = "\\DEPLOYMENTSEVER\deploy$\Applications\Adobe Flash Active X\install_flash_player_10_active_x.msi"
$FlashUpdatePL = "\\DEPLOYMENTSEVER\deploy$\Applications\Adobe Flash Plugin\install_flash_player_10_plugin.msi"
$JavaUpdate = "\\DEPLOYMENTSEVER\deploy$\Applications\Oracle Java x86\jre1.6.0_24.msi"
$QuickTimeUpdate = "\\DEPLOYMENTSEVER\deploy$\Applications\Apple QuickTime\QuickTime.msi"
<# ::: Get Domain Admin Credential ::: #>
$Cred = Get-Credential
<# ::: Needed for AD access. You need to have the Quest AD Powershell tools installed. ::: #>
Add-PSSnapin Quest.ActiveRoles.ADManagement
<# ::: Pull Computers from AD ::: Change OU if needed. ::: #>
$Computers = Get-QADComputer -searchroot 'domain.local/Workstations/Deployment'
$Skipped = @()
foreach($Computer in $Computers)
{	
	$ObjComputerName = New-Object PSObject
	$ObjComputerName = $computer.name
	$System = $ObjComputerName
<# ::: If the computer is not responding, record that we skipped it and continue. We can review this collection after the script completes. ::: #>
	if(-not (Test-Connection -Quiet $System -Count 1))
	{
		$Skipped += $System
	}
<# ::: Do Updates ::: #>
	(Get-WMIObject -Class Win32_Product -ComputerName $System -List -Credential $Cred).Install($FlashUpdateAX,$null,$true)
	(Get-WMIObject -Class Win32_Product -ComputerName $System -List -Credential $Cred).Install($FlashUpdatePL,$null,$true)
	(Get-WMIObject -Class Win32_Product -ComputerName $System -List -Credential $Cred).Install($JavaUpdate,$null,$true)
#	(Get-WMIObject -Class Win32_Product -ComputerName $System -List -Credential $Cred).Install($QuickTimeUpdate,$null,$true) <# Causes some systems to reboot #>
#	(Get-WMIObject -Class Win32_Process -ComputerName $System -List -Credential $Cred).Create("cmd.exe /c filehere.exe /switches")
}
<# ::: Show Systems Skipped ::: #>
"Computer Name:" + $Skipped | Out-File -Append -FilePath $env:USERPROFILE\Desktop\Report.txt
