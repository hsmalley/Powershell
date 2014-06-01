###################
# Locate PST Files
###################
# Created By Hugh Smalley for US Epperson
# http://google.com/profiles/hsmalley
# Protected Under CC-BY-SA - http://creativecommons.org/licenses/by-sa/3.0/
###################
# :Function: 
###################
#	This script will locate PST files under the specified drive and report on their locations. 
#	Also if need this script can move the PST to a specified location.
###################
# :Change Log:
###################
# 	Ver 1.0
#		Basic Searching and Reporting created
# 	Ver 1.1
#		Added Ablity to copy files to server.
###################

Add-PSSnapin Quest.ActiveRoles.ADManagement
$Computers = Get-QADComputer -searchroot 'domain.local/Workstations/Deployment'
$Skipped = @()
$Laptops = @()
foreach($Computer in $Computers)
{	
	$ObjComputerName = New-Object PSObject
	$ObjComputerName = $computer.name
	$System = $ObjComputerName
	$Path = "\\" + $System + "\C$\Documents and Settings"
	<# ::: If the computer is not responding, record that we skipped and log it. ::: #>
	if (-not (Test-Connection -Quiet $System -Count 1))
		{
			$Skipped += $System
		}
	<# ::: If the path does not exist or we do not have access to skip the system and log it. Some times Techs isn't a memeber of the administrators group ::: #>
	if (-not (Test-Path $Path))
		{
			$Skipped += $System
		}	
	<# ::: Test to see if the system is a laptop if it is skip and log it. ::: #>
	if ((Get-WmiObject -Class Win32_ComputerSystem -computer $System).PCSystemType -eq 2)
		{
			$Laptops += $System
		}
	<# ::: Find the PST Files under the user profile that are greater than 100MB ::: #>
	Get-ChildItem $Path -Include *.pst -Recurse | Where-Object { $_.Length -gt 100MB } | Format-Table -AutoSize -Property FullName | Out-File -Append -NoClobber -FilePath $Env:USERPROFILE\Desktop\PSTLIST.TXT
}

$Laptops | Out-File -Append -NoClobber -FilePath $Env:USERPROFILE\Desktop\LAPTOPS.TXT
$Skipped | Out-File -Append -NoClobber -FilePath $Env:USERPROFILE\Desktop\SKIPPED.TXT
