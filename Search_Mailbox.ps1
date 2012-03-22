#region Setup
#Connect to Exchange with different credentials
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://EXCHANGESERVER/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -AllowClobber
<#
Function Get-Mails {

	#Setup Variables
	$User1 = Read-Host -Prompt "Enter User ID to search"
	$User2 = Read-Host -Prompt "Enter User to send mail to"
	$Search = Read-Host -Prompt "Enter Subject to search for"

	#search mailbox with USER and Subject of Yummy Pies then create a folder in user's mailbox. To search for keywoards remove "Subject"
	Search-Mailbox -Identity "$User1" -SearchQuery "Subject:$Search" -TargetMailbox "$User2" -TargetFolder "$Search"
	$Again = Read-Host -Prompt "Do another? (Y/N):"
}
#endregion Setup

Get-Mails
If ($Again -ieq "Y") {Get-Mails} ELSE {}
#>
