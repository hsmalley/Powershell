Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
$Cred = Get-Credential
Connect-QADService -Service 'domaincontroler.local' -Credential $Cred
$User = Read-Host "Account to Unlock:"
Unlock-QADUser -Identity $User
