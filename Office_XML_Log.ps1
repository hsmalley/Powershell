Clear-Host
# Add Quest CMDLETS
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Get Computers
$Computers = Get-QADComputer -SearchRoot "OU=Workstations,DC=Domain,DC=Local"
# Create XML Document
$XMLDocument = New-Object System.XML.xmlDataDocument
# Make Root Tag
$XMLRoot = $XMLDocument.CreateElement('OfficeCounts')
# Get Rid of the console output
[void]$XMLDocument.AppendChild($XMLRoot)
# Get Time Stamp
$Date = Get-Date
# Start Processing Computers
$Computers | ForEach-Object {
	# CLEAR VARIABLES
	$Description = $null
	$Computer = $null
	$Office = $null
	$User = $null
	$Version = $null
	# Name Object in Pipe
	$Computer = $_
	# Setup XML Fields
	$XMLComputer = $XMLDocument.CreateElement('Computer')
	$XMLComputer.SetAttribute('WorkStation',$XMLComputer.WorkStation)
	$XMLComputer.SetAttribute('User',$XMLComputer.User)
	$XMLComputer.SetAttribute('Office',$XMLComputer.Office)
	$XMLComputer.SetAttribute('Description',$XMLComputer.Description)
	<# REAL WORK STARTS #>
	IF ((Test-Connection -ComputerName $Computer.Name -Quiet) -eq $true) {
		Write-Host $Computer.Name
		$User = (Get-WMiObject -Class Win32_ComputerSystem -ComputerName $Computer.Name).Username
		IF ($User -eq $null) {$User = "No User Logged On"}
		ELSE {
			$User = Get-QADUser -Identity $User
			$Description = $User.Name + " - " + $User.Office
			Write-Host $Description
			$User = $User.LogonName
			Write-Host $User
		}
		<# Start Office Search #>
		$Version = 0
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer.Name)
  		$RegKey = $Reg.OpenSubKey('Software\Microsoft\Office').GetSubKeyNames()
		$Regkey | ForEach-Object {
			IF ($_ -match '(\d+)\.') {
      			IF ([int]$Matches[1] -gt $Version) {
        		$Version = $Matches[1]
      			}
    		}
  		}
		IF ($Version) {
			IF ($Version -gt 0) {$Office = $Version}
			IF (($Version -ige "10") -and ($Version -lt 14)) {$Office = "Office XP"}
			IF ($Version -ige "14") {$Office = "Office 2010"}
    	} ELSE {$Office = "Not Found"}
		Write-Host $Office
		<# End Office Search #>
	} ELSE {
		$User = "OFFLINE"
		$OFFICE = "OFFLINE"
		IF ($Computer.Description -eq $null) {$Description = "_"}
		ELSE {$Description = $Computer.Description}
	}
	<# REAL WORK IS DONE #>
	# Add data in XML Object
	$XMLComputer.WorkStation = $Computer.Name
	$XMLComputer.User = $User
	$XMLComputer.Office = $Office
	$XMLComputer.Description = $Description
	#Save XML Object
 	[void]$XMLRoot.AppendChild($XMLComputer)
}
# And, lastly save the XML document out to disk
$File = "Path\Here"
$Number = $Date.TimeOfDay.Hours + $Date.TimeOfDay.Minutes
$File = $File + "Computers" + $Number + ".xml"
$XMLDocument.Save($File)