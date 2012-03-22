<#

Terminate User

This script will follow Standard User Termination.

	*	Export PST				*	Move SMTPs
	*	Wipe Mobile Device		*	Set Out Of Office
	*	Remove Title & Office	*	Hide in GAL

This Script Requires the Quest Active Directory cmdlets.


WIP:
	SMTPs Transfer
	Log to a file e.g., establish accountablity and documentation
	GUI
#>

Clear-Host
$Random = Get-Random
$Cred = Get-Credential
#Import-Module ActiveDirectory -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
#Import-Module PSCX -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
Connect-QADService -Service 'dc.local' -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchangeserver/PowerShell/ -Authentication Kerberos -Credential $Cred
Import-PSSession $Session -AllowClobber

Function YoureFired {
Clear-Host
	$Tech = Read-Host "Enter your alias/opID:"
	$TermUser = Read-Host "Enter user's alias/opID to be terminated:"
	$WipeMobile = Read-Host "Should the user's mobile device be wiped, does not include blackberries (Y/N):"
	#$SmtpTransfer = Read-Host "Do you want to transfer the SMTPs to another user? (Y/N):"
	$ExportToPST = Read-Host "Do you want to export the user's mailbox to a pst now? (Y/N):"

	#Functions to Run
	
	DoADTerms
	WipeMobile
	DoEmail

	$DoAnother = Read-Host "Would you like to do another one? (Y/N):"
		IF ($DoAnother -ieq "Y") {YoureFired}
}

Function WipeMobile {
Clear-Host
	Write-Warning "Remember, this DOES NOT WIPE BLACKBERRY DEVICES"
	Write-Warning "The Error that follows means there are not Active Sync Devices `
	Cannot bind argument to parameter 'Identity' because it is null."
	Sleep -Seconds "7"
	
	IF ($WipeMobile -ieq "Y") {
			$ASDevices = (Get-ActiveSyncDevice -mailbox $TermUser | Format-Table Identity)
			ForEach ($device in $ASDevices)
			{Clear-ActiveSyncDevice -Identity $device -NotificationEmailAddresses "$Tech@domain.local" -Verbose}
		}
	ELSE { 
		ForEach ($device in $ASDevices)
		{remove-activesyncdevice -identity $device -Verbose}
		}
	}

Function DoADTerms {
Clear-Host	
	# Native AD Commands, Needs Server with Active Directory Web Services on it.
	# Set-ADUser -Identity $TermUser -Credential $Cred -Server 'dc.local' -Title '' -Office ''
	# Disable-ADAccount -Identity $TermUser -Credential $Cred -Server 'dc.local'
	
	# Quest AD cmdlets - Note: Without Active Roles Server these commands are run as your user account.
	Set-QADUser -Identity $TermUser -title '' -Office ''
	Disable-QADUser -Identity $TermUser -Credential $Cred
	}

Function DoEmail {
Clear-Host
	Set-Mailbox -Identity $TermUser -hiddenfromaddresslistsenabled $True
	Set-MailboxAutoReplyConfiguration -Identity $TermUser -autoreplystate enabled -InternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with US." -ExternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with US, however your email has been forwarded to someone who can help you."
	
	$Emails = Get-Mailbox -Identity $TermUser 
	$Emails.PrimarySmtpAddress = $Emails.PrimarySmtpAddress -replace "@domin.local",".12345@domain.local"
	Set-Mailbox -Identity $TermUser -PrimarySmtpAddress $Emails.PrimarySmtpAddress -EmailAddressPolicyEnabled $false
	
	<#WIP:
	#Get SMTPS From User, Move Change the name XXX@DOMAIN.LOCAL TO XXX_123456@DOMAIN.LOCAL Then Move XXX@DOMAIN.LOCAL to another user.
	#Grab user name & name and make them into seperate vars
	
	IF ($SmtpTransfer -ieq "Y") 
		{
			$SmtpTransferID = Read-Host "Enter user's alias/opID to transfer the SMTPs to:"
			$STMPs = ($Emails.EmailAddresses | Select-String -CaseSensitive "smtp") -replace "smtp:","" 
			$SMTPTransfterUser = Get-Mailbox -Identity $SmtpTransferID
			# $Emails.EmailAddresses = ($Emails.EmailAddresses | Select-String -CaseSensitive "smtp") -replace "smtp:",""
			$Emails.EmailAddresses += $STMP
			Set-Mailbox -Identity $SmtpTransferID -EmailAddresses $Emails.EmailAddresses -WhatIf
		} 
	ELSE 
		{
			Write-Warning "SMTPs will be appened with .123456 e.g., FIRST.LAST.123456@DOMAIN.LOCAL or SHORTNAME.123456@DOMAIN.LOCAL"
			$Emails.EmailAddresses = $Emails.EmailAddresses -replace "smtp:","" -replace "@DOMAIN.LOCAL",".123456@DOMAIN.LOCAL"
		}
	#>
	
	IF ($ExportToPST -ieq "Y")
	{Write-Warning "Depending on the amount of email this might take a while."
	Sleep -Seconds "5"
	$ExportTo = Read-Host "Enter the path for the pst file, must be a UNC path, e.g. \\NAS\Misc\User.PST:"
	New-MailboxExportRequest -Mailbox $TermUser -FilePath $ExportTo
    #Write-Host "Waiting for Export To Complete Displaying Status in 30 seconds"
	#Sleep -Seconds "30"
	#Get-MailboxExportRequest | Get-MailboxExportRequestStatistics
	}
} 

#In the words of The Donald
YoureFired
