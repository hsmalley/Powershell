$colComputers = get-content c:\iso\dlo\DLO_Updates.txt
foreach ($strComputer in $colComputers)
{
{
$Path = "\\"+ $strComputer + "\C$\Program Files\Symantec\Backup Exec\dlo\dloclientu.exe"
$File = get-item $Path
}
Write-Output $File.VersionInfo | Select {Get-Date}, {$strComputer},FileVersion | Out-GridView
}
