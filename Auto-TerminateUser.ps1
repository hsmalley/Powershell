<# This is a auto termination script. It's going to royaly screw an account. So make sure you want that. Okay?
# You really shouldn't monkey around with this as it assumes the following:
#	0. YOU KNOW WHAT YOU'RE DOING!
#	1. The account running it has exchange admin access
#	2. The computer running is set to EST.
#	3. The account running it has access to the IT share.
#	4. You have the Quest Active Roles AD Cmdlets installed.
#	5. 0-4
#>
# Add Quest CMDLETS
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Get Timestamp
$Date = Get-Date
# Setup Email fields
$SMTP = New-Object Net.Mail.SMTPClient
$SMTP.Host = "mailserver.domain.local"
$SMTP.TargetName = "mailserver.domain.local"
$EmailFrom = "ITTechs@domain.local"
$EmailTO = "ITTechs@domain.local"
$Subject = "User Termination - $Env:COMPUTERNAME"
# Start crafting the message.
$Message = New-Object System.Net.Mail.MailMessage $EmailFrom, $EmailTO
$Message.Subject = $Subject
$Message.IsBodyHTML = $true
# CSS Style for Email.
$Style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$Style = $Style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$Style = $Style + "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
$Style = $Style + "TD{border: 1px solid black; padding: 5px;}"
$Style = $Style + "</style>"
# Generate a Random Password for the User
Function Generate-Password {
	$Password = $null
	$Random = $null
	#Set up random number generator
	$Random = New-Object System.Random
	#Generate a new 10 character password
	1..10 | ForEach { $Password = $Password + [char]$Random.next(33,127) }
	$Password
}
#Connect to Exchange Server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://ExchServer1/PowerShell/"
Import-PSSession $Session -AllowClobber
# Import the list of people heading to Château d'If
$Users = Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermUsers.csv"
$Users | ForEach-Object {
	# Setup Variables
	$User = $_
	$TimeZone = $User.TimeZone
	$TermUser = $User.OPID
	$TermDate = $User.Date
	# DATE CHECK
	IF ($TermDate -ieq (Get-Date -Format yyyyMMdd)) {
		# TIME ZONES
		# The task runs an hour after these times. So we're safe with this.
		IF (($TimeZone -ieq "EST") -and ($Date.Hour -gt "16")) {$RunTerm = "YES"}
		ELSEIF (($TimeZone -ieq "CST") -and ($Date.Hour -gt "17")) {$RunTerm = "YES"}
		ELSEIF (($TimeZone -ieq "MNT") -and ($Date.Hour -gt "18")) {$RunTerm = "YES"}
		ELSEIF (($TimeZone -ieq "PST") -and ($Date.Hour -gt "19")) {$RunTerm = "YES"}
		ELSE {$RunTerm = "NO"}
	}
	# Run the termination if it's time.
	IF (($RunTerm -eq "YES") -and ((Get-QADObject -Identity $TermUser | Get-QADMemberOf) -notmatch "Disabled Users")) {
		# Generate a Random Password for the User
		$Password = Generate-Password
		# Convert OPID to User Account information for Quest AD Tools
		$ADTermUser = Get-QADUser -Identity $TermUser
		# Move User to the Disabled User OU
		Move-QADObject -Identity $ADTermUser -NewParentContainer (Get-QADObject "Users - Disabled")
		# Refresh ADTermUser
		$ADTermUser = Get-QADUser -Identity $TermUser
		# Try to correct possible human error.
		$TermUser = $ADTermUser.LogonName
		# Get the Display Name for the User
		$ADTermUserDisplayName = $ADTermUser.DisplayName
		# Get the User's Manager. If the user doesn't have a manager (OR has an invalid manager) a manager will be assigned to them.
		IF ((Get-QADObject -Identity $ADTermUser.Manager) -eq $null) {$ADTermUserManager = Get-QADUser -Identity "disabledusermanager"}
		ELSE {$ADTermUserManager = Get-QADObject -Identity $ADTermUser.Manager}
		# Set the User's Manager's Logon Name
		$TermUserManager = $ADTermUserManager.LogonName
		# Get the Manager's Display Name
		$ADTermUserManagerDisplayName = $ADTermUserManager.DisplayName
		#Set Expire Date
		$Expire = $Date.AddMonths(13)
		<# START CSV LOG #>
		$Log =  New-Object PSObject
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Run Date" -Value $Date -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Remove After" -Value $Expire -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "User ID" -Value $TermUser -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "User Name" -Value $ADTermUserDisplayName -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Manager ID" -Value $TermUserManager -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Manager" -Value $ADTermUserManagerDisplayName -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Tech" -Value "AUTOMAGICAL" -Force
		$TermLog = @()
		$TermLog += $Log
		$LogPath = "\\FileServer\ITSHARE\User Termination Log\TermLog.csv"
		$TermLog | Export-Csv -Append -Force -NoTypeInformation $LogPath
		<# END CSV LOG #>
		# Disable AD Account and add to Disabled Users Group
		Disable-QADUser -Identity $ADTermUser -Confirm:$false
		Add-QADGroupMember -Identity "Disabled Users" -Member $ADTermUser -Confirm:$false
		# Set Description
		$Description = ($AdtermUser.Description + " - Terminated On: $Date")
		# Removes Title, Office, Phones, Fax, changes password to random number, set description, and change primary group to Disabled Users.
		# It's not on a single line for a reason, go ahead and test the wheels fate; if that's what you're into...
		Set-QADUser -Identity $ADTermUser -Title '' -Office '' -PhoneNumber '' -MobilePhone '' -Pager '' -Fax ''  -Description $Description -Confirm:$false
		Set-QADUser -Identity $ADTermUser -UserMustChangePassword $true -UserPassword $Password -Confirm:$false
		Set-QADUser -Identity $ADTermUser -AccountExpires $Expire -Confirm:$false
		Set-QADUser -Identity $ADTermUser -objectAttributes @{PrimaryGroupID='19791'} -Confirm:$false #GroupID should be a disable users security group
		# Remove Groups
		$ADTermUser.MemberOf | Remove-QADGroupMember -Member $ADTermUser -Confirm:$false
		# Hide the user from the GAL and set forwarding of the user's mail to the manager.
		Set-Mailbox -Identity $TermUser -hiddenfromaddresslistsenabled $True -ForwardingAddress $TermUserManager
		# Set out of office reply aka Let's people who who got canned.
		Set-MailboxAutoReplyConfiguration -Identity $TermUser -autoreplystate enabled -InternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with LUA." -ExternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with LUA, however your email has been forwarded to someone who can help you."
		# Give Manager Full Access
		Add-MailboxPermission $TermUser -User $TermUserManager -AccessRights fullaccess
		# Punches their mail clients and phones in the face.
		Set-CASMailbox -Identity $TermUser -OwaEnabled $false -EcpEnabled $false -EwsEnabled $false -MapiEnabled $false -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false -Confirm:$false
		Set-CASMailbox -Identity $TermUser -MapiBlockOutlookRpcHttp $true -EwsAllowMacOutlook $false -EwsAllowOutlook $false -EwsAllowEntourage $false -Confirm:$false
		# Banishes their phones.
		Get-ActivesyncDevice -Mailbox $TermUser | Remove-ActiveSyncDevice -Confirm:$false
		# Send an email that the termination is done.
		$Body = "AD Termination for $TermUser has completed. <br>"
		$Body += "Script Name:	Auto_Terminate_User.ps1 <br>"
		$Body += "Running on: $Env:COMPUTERNAME <br>"
		$Body += "<hr>"
		$Body += "<b> USER REPORT </b><br>"
		$Body += Get-QADUser -Identity $TermUser | Select-Object -Property DisplayName,AccountIsDisabled,UserMustChangePassword,PrimaryGroupId,Manager,AccountExpirationStatus | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER MAILBOX REPORT </b><br>"
		$Body += Get-CASMailbox $TermUser | Select-Object DisplayName,MapiEnabled,MAPIBlockOutlookRpcHttp | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER DEVICE REPORT - <i>Should be empty</i></b><br>"
		$Body += Get-ActivesyncDevice -Mailbox $TermUser | Select-Object DeviceType,DeviceId,DeviceUserAgent | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TO TERM </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermLog.csv" | ConvertTo-Html
		$Body += "<hr>"
		$Message.Body = ConvertTo-Html -Head $Style -Body $Body
		$SMTP.Send($Message)
	}
	# Clean Up - This part checks to make sure the user account was terminated correctly and enable MAPI on the mailbox.
	# If the user is a member of Disabled Users, account is disabled, mapi is disabled on mailbox, and it's been more than 1 day.
	IF (($RunTerm -eq "YES") -and ((Get-QADObject -Identity $TermUser | Get-QADMemberOf) -match "Disabled Users") -and ((Get-QADObject -Identity $TermUser).get_AccountIsDisabled() -eq $true) -and ($Date.AddDays(1) -gt ((Get-QADObject -Identity $TermUser).get_ModificationDate())) -and ((Get-CASMailbox -Identity $TermUser).MAPIEnabled -eq $false)) {
		# Enable MAPI
		Set-CASMailbox -Identity $TermUser -MapiEnabled $true -MapiBlockOutlookRpcHttp $false -Confirm:$false
		# Start Moving The Mailbox aka let's deport them to Canada...
		New-MoveRequest -Identity $TermUser -BadItemLimit 5 -TargetDatabase "Canada" -Confirm:$false
		# Send Email
		$Body = "Clean up on $TermUser is done.<br>"
		$Body += "MAPI has been enabled on mailbox.<br>"
		$Body += "Mailbox is being deported to Canada.<br>"
		$Body += "Please check Exchange move requests for errors moving the mailbox. <br>"
		$Body += "Script Name:	Auto_Terminate_User.ps1 <br>"
		$Body += "Running on: $Env:COMPUTERNAME <hr>"
		$Body += "<b> USER MAILBOX REPORT </b><br>"
		$Body += Get-CASMailbox $TermUser | Select-Object DisplayName,MapiEnabled,MAPIBlockOutlookRpcHttp | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER REPORT </b><br>"
		$Body += Get-QADUser -Identity $TermUser | Select-Object -Property DisplayName,AccountIsDisabled,UserMustChangePassword,PrimaryGroupId,Manager,AccountExpirationStatus | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TO TERM </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermUsers.csv" | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TERMLOG </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermLog.csv" | ConvertTo-Html
		$Message.Body = ConvertTo-Html -Head $Style -Body $Body
		$SMTP.Send($Message)
	}
	# If the account is not in the disabled users or disabled run this
	ELSEIF (($RunTerm -eq "YES") -and ((Get-QADObject -Identity $TermUser | Get-QADMemberOf) -notmatch "Disabled Users") -or ((Get-QADObject -Identity $TermUser).get_AccountIsDisabled() -eq $false)) {
		# Send Email
		$Body = "Clean up on $TermUser is not done. <br><br>"
		$Body += "Either account is not terminated correctly or is not disabled.<hr>"
		$Body += "<b> USER REPORT </b><br>"
		$Body += Get-QADUser -Identity $TermUser | Select-Object -Property DisplayName,AccountIsDisabled,UserMustChangePassword,PrimaryGroupId,Manager,AccountExpirationStatus | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER MAILBOX REPORT </b><br>"
		$Body += Get-CASMailbox $TermUser | Select-Object DisplayName,MapiEnabled,MAPIBlockOutlookRpcHttp | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TO TERM </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermUsers.csv" | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TERMLOG </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermLog.csv" | ConvertTo-Html
		$Message.Body = ConvertTo-Html -Head $Style -Body $Body
		$SMTP.Send($Message)
	}
	# If the account hasn't been terminated for more than one day run this.
	ELSEIF (($RunTerm -eq "YES") -and ($Date.AddDays(1) -lt ((Get-QADObject -Identity $TermUser).get_ModificationDate()))) {
		# Send Email
		$Body = "Clean up on $TermUser is not done. <br><br>"
		$Body += "24 hours have not passed since termination.<hr>"
		$Body += "<b> USER TO TERM </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermUsers.csv" | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TERMLOG </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermLog.csv" | ConvertTo-Html
		$Message.Body = ConvertTo-Html -Head $Style -Body $Body
		$SMTP.Send($Message)
	}
	# Otherwise, there's nothing to do.
	ELSE {$RunTerm = "NO"}
	IF ($RunTerm -eq "NO") {
		<# Send Email
		$Body = "NOTHING TO DO!<hr>"
		$Body += "<b> USER TO TERM </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermUsers.csv" | ConvertTo-Html
		$Body += "<hr>"
		$Body += "<b> USER TERMLOG </b><br>"
		$Body += Import-Csv "\\FileServer\ITSHARE\User Termination Log\TermLog.csv" | ConvertTo-Html
		$Message.Body = ConvertTo-Html -Head $Style -Body $Body
		$SMTP.Send($Message)
		#>
	}
}
