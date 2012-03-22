Clear-Host
Add-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue #This is needed to pull computers from AD
$Skipped = @()	#Create Array for Skipped Computers
$Finished = @()	#Create Array for Finished Computers
$Offline = @()	#Create Array for Offline Computers
$Computers = Get-QADComputer -SearchScope Subtree -SearchRoot "OU=Workstations,DC=local,dc=com"	#Get Computers from AD
FOREACH ($Computer in $Computers)	#Process Computers
	{
		$ObjComputerName = New-Object PSObject
		$ObjComputerName = $Computer.name
		$System = $ObjComputerName
		IF (Test-Connection -ComputerName $System -Quiet -Count 1)	#Tests to see if computer is online
			{
				IF (Test-Path "\\$System\C$\Users\Public")	#This would indicate Windows 7 OR Vista
					{
						Copy-Item -Path "C:\Users\Public\Training" -Destination "\\$System\C$\Users\Public" -Recurse -Force
						Write-Host "Finished Win7 System $System"
						$Finished += $System
					}
				ELSE 
					{ 	
						IF (Test-Path "\\$System\C$\Documents and Settings\All Users")
							{
								Copy-Item -Path "C:\Users\Public\Training" -Destination "\\$System\C$\Documents and Settings\All Users" -Recurse -Force
								Write-Host "Finished WinXP System $System"
								$Finished += $System
							}
					}
			}
		ELSE
			{
				Write-Host "System $System is offline"
				$Skipped += $System
			}
	}
FOREACH ($System in $Skipped)	#Process Skipped Computers
	{
		Write-Host "Retrying Offline System $System"
		IF (Test-Connection -ComputerName $System -Quiet -Count 1)
			{
				IF (Test-Path "\\$System\C$\Users\Public")
					{
						Copy-Item -Path "C:\Users\Public\Training" -Destination "\\$System\C$\Users\Public" -Recurse -Force
						Write-Host "Finished Win7 System $System"
						$Finished += $System
					}
				ELSE 
					{ 	
						IF (Test-Path "\\$System\C$\Documents and Settings\All Users")
							{
								Copy-Item -Path "C:\Users\Public\Training" -Destination "\\$System\C$\Documents and Settings\All Users" -Recurse -Force
								Write-Host "Finished WinXP System $System"
								$Finished += $System
							}
					}
			}
		ELSE
			{
				Write-Host "System $System is Offline"
				$Offline +=$System
			}
	}
Write-Host "Offline Systems:	$Offline"
Write-Host "Finished Systems:	$Finished"
