<#

Set Out of Office User

This script will change a user's out of office message.

This Script Requires the Quest Active Directory cmdlets.

#>

Clear-Host
$Random = Get-Random
$Cred = Get-Credential
#Import-Module ActiveDirectory -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
#Import-Module PSCX -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
Connect-QADService -Service 'DC.Domain.Local' -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch10/PowerShell/ -Authentication Kerberos -Credential $Cred
Import-PSSession $Session -AllowClobber

$OOOUser = Read-Host "Enter user's alias/opID to have the Out Of Office Message Set:"
$OOOMessage = Read-Host "Enter the Out Of Office Message:"
Set-MailboxAutoReplyConfiguration -Identity $OOOUser -autoreplystate enabled -InternalMessage $OOOMessage -ExternalMessage $OOOMessage
