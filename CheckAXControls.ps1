<# 
	.Synopsis 
   		Check ActiveX Controls Workstations.
	.Description 
		This script checks the \Windows\Downloaded Program Files\ directory for ocx, dll, & exe files and enums their properties into an excel sheet.
	.Notes
    	NAME: CheckAXControls.ps1
    	AUTHOR: Hugh Smalley
	.Errors
		* If running script over again in the same session you will get the following error. "Add-PSSnapin : Cannot add Windows PowerShell snap-in Quest.ActiveRoles.ADManagement because it is already added. Verify the name of the snap-in and try again."
#>

<# ::: Get Domain Admin Credential ::: #>
# $Cred = Get-Credential
<# ::: Needed for AD access. You need to have the Quest AD Powershell tools installed. ::: #>
Add-PSSnapin Quest.ActiveRoles.ADManagement
<# ::: Pull Computers from AD ::: Change OU if needed. ::: #>
$Computers = Get-QADComputer -searchroot 'DOMAIN.LOCAL/Workstations'
$Skipped = @()
<# ::: Create Excel Sheet :::  #>
$erroractionpreference = "SilentlyContinue"
$a = New-Object -comobject Excel.Application
$a.visible = $True
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "File Name"
$c.Cells.Item(1,3) = "Product Name"
$c.Cells.Item(1,4) = "Version"
$c.Cells.Item(1,5) = "Description"
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$intRow = @()
$intRow = 2
foreach($Computer in $Computers)
{	
	$ObjComputerName = New-Object PSObject
	$ObjComputerName = $computer.name
	$System = $ObjComputerName
<# ::: If the computer is not responding, record that we skipped it and continue. We can review this collection after the script completes. ::: #>
	if(-not (Test-Connection -Quiet $System -Count 1))
		{
		$Skipped += $System #Left in so a text file could be created if needed
		$c.Cells.Item($intRow,1) = "$System"
		$c.Cells.Item($intRow,2) = "Offline"
		$c.Cells.Item($intRow,3) = "Offline"
		$c.Cells.Item($intRow,4) = "Offline"
		$c.Cells.Item($intRow,5) = "Offline"
		$intRow = $intRow + 1
		}
	if((Test-Connection -Quiet $System -Count 1))
	{
	$Path = "\\" + $System + "\C$\Windows\Downloaded Program Files"
		if(-not (Test-Path $Path))
			{
			$Skipped += $System #Left in so a text file could be created if needed
			$c.Cells.Item($intRow,1) = "$System"
			$c.Cells.Item($intRow,2) = "Access Denied"
			$c.Cells.Item($intRow,3) = "Access Denied"
			$c.Cells.Item($intRow,4) = "Access Denied"
			$c.Cells.Item($intRow,5) = "Access Denied"
			$intRow = $intRow + 1
			}	
		if((Test-Path $Path))
			{
			$OCX = Get-Item $Path\*.ocx
			foreach ($File in $OCX) 
				{
				$c.Cells.Item($intRow,1) = $System
				$c.Cells.Item($intRow,2) = $File.VersionInfo.OriginalFileName
				$c.Cells.Item($intRow,3) = $File.VersionInfo.ProductName
				$c.Cells.Item($intRow,4) = $File.VersionInfo.FileVersion
				$c.Cells.Item($intRow,5) = $File.VersionInfo.FileDescription
				$intRow = $intRow + 1
				}
			$DLL = Get-Item $Path\*.dll
			foreach ($File in $DLL)
				{
				$c.Cells.Item($intRow,1) = $System
				$c.Cells.Item($intRow,2) = $File.VersionInfo.OriginalFileName
				$c.Cells.Item($intRow,3) = $File.VersionInfo.ProductName
				$c.Cells.Item($intRow,4) = $File.VersionInfo.FileVersion
				$c.Cells.Item($intRow,5) = $File.VersionInfo.FileDescription
				$intRow = $intRow + 1
				}		
			$EXE = Get-Item $Path\*.exe
			foreach ($File in $EXE) 
				{
				$c.Cells.Item($intRow,1) = $System
				$c.Cells.Item($intRow,2) = $File.VersionInfo.OriginalFileName
				$c.Cells.Item($intRow,3) = $File.VersionInfo.ProductName
				$c.Cells.Item($intRow,4) = $File.VersionInfo.FileVersion
				$c.Cells.Item($intRow,5) = $File.VersionInfo.FileDescription
				$intRow = $intRow + 1
				}
			}
		}
}
$d.EntireColumn.AutoFit()
