$erroractionpreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "DLO Version"
$c.Cells.Item(1,3) = "Report Time Stamp"

$d = $c.DOMAIN.LOCALdRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$colComputers = get-content c:\iso\dlo\keepscan.txt

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
