Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
$Cred = Get-Credential
$Computers = Get-QADComputer -SearchRoot "OU=Workstations,DC=DOMAIN,DC=LOCAL" -Credential $Cred

$Computers | ForEach-Object {
$Computer = $_.name
$Office = C:\Users\Hsmalley\Get-InstalledApp.ps1 -ComputerName $Computer -AppName "Microsoft Office XP Standard*" | Select ComputerName,AppName
$Office | Export-Csv -Append -Path C:\Users\Hsmalley\Office.csv -NoTypeInformation
$Office = C:\Users\Hsmalley\Get-InstalledApp.ps1 -ComputerName $Computer -AppName "Microsoft Office XP Professional*" | Select ComputerName,AppName
$Office | Export-Csv -Append -Path C:\Users\Hsmalley\Office.csv -NoTypeInformation
$Office = C:\Users\Hsmalley\Get-InstalledApp.ps1 -ComputerName $Computer -AppName "Microsoft Office * 2010" | Select ComputerName,AppName
$Office | Export-Csv -Append -Path C:\Users\Hsmalley\Office.csv -NoTypeInformation
}

$Username = (Get-WMiObject -Class Win32_ComputerSystem -ComputerName $ComputerName).Username
