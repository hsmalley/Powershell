#			 DLO SCRUBBER         
#	 	CREATED BY HUGH 10/2009    
#
# You will need to following files:
# PSEXEC.EXE
# DLO_UPDATES.TXT
# DLOSCRUB.CMD
# This will remove and older DLO and Install a new one in its place on a given list.

#Choose DLO Version Here - TODO Allow User to input version number
$dlover = "3.10.338.7401"
#Computer List - TODO Allow user to choose list or choose to grab laptops from AD.
$colComputers = Get-Content C:\ISO\DLO\DLO_UPDATES.TXT
foreach ($strComputer in $colComputers)
 {
  $Path = "\\"+ $strComputer + "\C$\Program Files\Symantec\Backup Exec\dlo\dloclientu.exe"
  $File = get-item $Path
  if ($File.VersionInfo.FileVersion -ne $dlover) 
  {
   #Path to PSEXE & DLOSCRUB.CMD - TODO Allow user to choose where they are located.
   $DLOSCRUB = "C:\ISO\DLO\PSEXEC.EXE \\"+ $strComputer + " -S -C C:\ISO\DLO\DLOSCRUB.CMD"
   Invoke-Expression $DLOSCRUB
  }
}
#Start Creating Spreadsheet of systems that have or need DLO.

$erroractionpreference = "SilentlyContinue"
$a = New-Object -comobject Excel.Application
$a.visible = $True
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "DLO Version"
$c.Cells.Item(1,3) = "Report Time Stamp"
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$intRow = 2
foreach ($strComputer in $colComputers)
	{
        $c.Cells.Item($intRow,1)  = $strComputer
        Function GetFileInfo
    	{
			$Path = "\\"+ $strComputer + "\C$\Program Files\Symantec\Backup Exec\dlo\dloclientu.exe"
			$File = get-item $Path
			$c.Cells.Item($intRow,2)  = $File.VersionInfo.FileVersion
		}
		GetFileInfo
		$c.Cells.Item($intRow,3) = Get-date
		$intRow = $intRow + 1
	}
$d.EntireColumn.AutoFit()
