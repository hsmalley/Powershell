Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue # Add the required snap in
$OldEMail = "sams-tasty-pies.com" # Set the email address you want to change here.
$NewEMail = "sals-fantastic-hams.org" # Put the new email address you want to use here.
$Users = Get-QADUser -SizeLimit 10000 #Get all AD User Objects. NOTE: The size limit is needed.
$Users | Where-Object -Property "Email" -Like "*$OldEMail*" | ForEach-Object { # Find the AD accounts with the old email
	$EMail = $_.Email # Create the email variable
	$EMail = $EMail -replace $OldEMail,$NewEMail # Replace the old email with new email.
	Set-QADUser -Identity $_ -Email $Email # Change it on the account
}
