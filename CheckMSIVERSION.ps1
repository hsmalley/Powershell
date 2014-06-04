$erroractionpreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "File Name"
$c.Cells.Item(1,3) = "Version"
$c.Cells.Item(1,4) = "Report Time Stamp"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$colComputers = get-content C:\Temp\Machinelist.txt

foreach ($strComputer in $colComputers)
{
$c.Cells.Item($intRow,1)  = $strComputer

Function GetFileInfo
{

$Path = "\\"+ $strComputer + "\C$\Windows\System32\msi.dll"

$File = get-item $Path

$c.Cells.Item($intRow,2)  = $File.Name
$c.Cells.Item($intRow,3)  = $File.VersionInfo.Productversion
}

GetFileInfo

$c.Cells.Item($intRow,4) = Get-date


$intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()
