<#
Terminate User

	*	Remove mobile device					*	Set Out Of Office
	*	Log terminiation details				*	Hide in GAL
	*	Disable account							*	Change account password
	*	Forward Email to Manager				*	Move account to Users - Disabled
	*	Change primary group to Disabled Users	*	Remove groups
	*	Remove the following account details:	*	Put termination date in description
		*	Phone	*	Office
		*	Mobile	*	Title
		*	Fax		*	Description

This Script Requires the Quest Active Directory cmdlets.

Change Log:
	Sep. 2012	-	Rewrote Script
	
WIP:
	Logging
	GUI - Nice to have but not needed
#>

Clear-Host
# Get User/Pass
$Cred = Get-Credential
# Add Quest CMDLETS
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Connect to DC
Connect-QADService -Service 'bdc.domain.local' -Credential $Cred
#Setup Email
$SMTPServer = '10.10.10.120';
$EmailFrom = "HelpDesk@domain.local.com";
$SMTP = New-Object Net.Mail.SmtpClient($SMTPServer);

Function YoureFired {
Clear-Host
	#Functions to Run
	SetupLogging
	DoADTerms
	RemoveMobile
	DoEmail
	SaveLog
	$DoAnother = Read-Host "Would you like to do another one? (Y/N):"
		IF ($DoAnother -ieq "Y") {YoureFired}
		ELSE {MonitorEmailMove}
	}

Function DoADTerms {
Clear-Host	
	# Convert OPID to User Account information for Quest AD Tools
	$ADTermUser = Get-QADUser -Identity $TermUser -Credential $Cred
	# Try to correct possible human error.
	$TermUser = $ADTermUser.LogonName
	# Get the Display Name for the User
	$ADTermUserDisplayName = $ADTermUser.DisplayName
	# Get the User's Manager.
	$ADTermUserManager = Get-QADObject -Identity $ADTermUser.Manager -Credential $Cred
	# Set the User's Manager's Logon Name
	$TermUserManager = $ADTermUserManager.LogonName
	# Get the Manager's Display Name
	$ADTermUserManagerDisplayName = $ADTermUserManager.DisplayName
	# Confirm Data Manager for the User
	Write-Host "The Manager for $ADTermUserDisplayName is listed as $ADTermUserManagerDisplayName"
	# Set the Data Manager IF needed
	$SetManager = Read-Host "Is $ADTermUserManagerDisplayName going to be the data manager for $ADTermUserDisplayName ?:"
	IF ($SetManager -ieq "Y")
		{
			$TermUserManager = Read-Host "Enter the Managers opID:"
			$ADTermUserManager = Get-QADObject -Identity $TermUserManager -Credential $Cred
			$TermUserManager = $ADTermUserManager.LogonName
			$ADTermUserManagerDisplayName = $ADTermUserManager.DisplayName
			Write-Warning "Setting Data Manager to $TermUserManager - $ADTermUserManagerDisplayName" -WarningAction Inquire
			# Set the new manager on the User's account
			Set-QADUser -Identity $ADTermUser -Manager $ADTermUserManager -Credential $Cred
			$ADTermUser = Get-QADUser -Identity $TermUser -Credential $Cred
		}
	# Disable AD Account
	Disable-QADUser -Identity $ADTermUser -Credential $Cred
	# Move User to the Disabled User OU
	Move-QADObject -Identity $ADTermUser -NewParentContainer (Get-QADObject "Users -  Disabled" <#Has Extra Space in AD. Is it needed for another script or Typo?#>) -Credential $Cred
	# Add to Disabled User Group
	Add-QADGroupMember -Identity "Disabled Users" -Member $ADTermUser -Credential $Cred
	# Removes Title, Office, Phones, Fax, changes password to random number, set description, change primary group to Disabled Users, and deny Dialin permission
	Set-QADUser -Identity $ADTermUser -Title '' -Office '' -PhoneNumber '' -MobilePhone '' -Pager '' -Fax '' -UserMustChangePassword $true -UserPassword (Get-Random) -Description ("Terminated User $(Get-Date)") -objectAttributes @{primaryGroupID=(Get-QADGroup 'Disabled Users').PrimaryGroupToken;msNPAllowDialin=$false} -Credential $Cred
	# Remove Groups
	$ADTermUser.MemberOf | Remove-QADGroupMember -Member $ADTermUser -Credential $Cred
	}

Function DoEmail {
Clear-Host
	# Get Exchange Server
	$ExchangeServer = Read-Host "What Exchange Server do you want to connect to? <Server/OtherServer>: "
	IF ($ExchangeServer -ieq "Server") {$ExchangeServer = "http://Serverexch10/PowerShell/"}
	IF ($ExchangeServer -ieq "OtherServeranta") {$ExchangeServer = "http://OtherServerexch10/PowerShell/"}
	#Connect to Exchange Server
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeServer -Authentication Kerberos -Credential $Cred
	Import-PSSession $Session -AllowClobber
	# Get the user's mailbox
	$Emails = Get-Mailbox -Identity $TermUser
	# We are no longer changing the SMTP address for the user.
	# $Emails.PrimarySmtpAddress = $Emails.PrimarySmtpAddress -replace "@domain.local",".1905@domain.local"
	# Hide the user from the GAL and set forwarding of the user's mail to the manager.
	Set-Mailbox -Identity $TermUser -hiddenfromaddresslistsenabled $True -ForwardingAddress $TermUserManager #-PrimarySmtpAddress $Emails.PrimarySmtpAddress -EmailAddressPolicyEnabled $false
	# Set out of office reply
	Set-MailboxAutoReplyConfiguration -Identity $TermUser -autoreplystate enabled -InternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with LUA." -ExternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with LUA, however your email has been forwarded to someone who can help you."
	# Start Moving The Mailbox
	StartMailboxMove
} 
# Start moving the data
# Adapted from http://sysadmin.flakshack.com/post/5088222801/script-to-move-exchange-2010-mailboxes
Function StartMailboxMove {
Clear-Host
	$Body = "
		Mailbox for $TermUser will start moving in 15 minutes. Please exit the Mailbox for $TermUser now.
		";

	# Submit the move requests in SUSPENDED state
	New-MoveRequest -Identity $TermUser -BadItemLimit 5 -Suspend:$true;

	# Email a notIFication to make sure everyone is out of the mailbox.
	Write-Host "Sending 15-minute until move email notIFication to $TermUser";
	$TimeOfMove = ((get-date).AddMinutes(15)).toShortTimeString();  
	$Subject = "Mailbox move starting @ $TimeOfMove";
	# Send Email to Techs
	$EmailAddress = "iitechs@domain.local"
	$SMTP.Send($EmailFrom, $EmailAddress, $Subject, $Body);
	# Send Email to The Mailbox.
	$EmailAddress = (Get-Mailbox -Identity $TermUser | Select PrimarySmtpAddress).PrimarySmtpAddress;		
	$SMTP.Send($EmailFrom, $EmailAddress, $Subject, $Body);
}

# Monitor the Move Requests
# Adapted from http://sysadmin.flakshack.com/post/5088222801/script-to-move-exchange-2010-mailboxes
Function Process-Mailboxes {
 	While ($MoveRequests = Get-MoveRequest) {
		foreach($MoveRequest in $MoveRequests) {
			$MailboxIdentity = $MoveRequest.Identity;
			$TargetDatabase = $MoveRequest.TargetDatabase;
			$MailboxAlias = $MoveRequest.Alias;
			$DisplayName = $MoveRequest.DisplayName;
			$Status = $MoveRequest.Status

			$Results = Get-MoveRequestStatistics -Identity $MailboxIdentity  | Select DisplayName, PercentComplete, BadItemsEncountered, TotalMailboxSize, TotalMailboxItemCount, TotalInProgressDuration,TotalSuspendedDuration,TotalQueuedDuration;
			$PercentComplete = $Results.PercentComplete;
			$BadItemsEncountered = $Results.BadItemsEncountered;
			$TotalMailboxSize = $Results.TotalMailboxSize;
			$Duration = $Results.TotalInProgressDuration;
			$ItemCount = $Results.TotalMailboxItemCount;
			$TimeSuspended = $Results.TotalSuspendedDuration;
			$TimeQueued = $Results.TotalQueuedDuration;

			IF (($Status -eq "InProgress") -or ($Status -eq "CompletionInProgress") -or ($Status -eq "Completed")) {
				Write-Host "$MailboxAlias	$percentComplete%	duration: $duration	Target: $TargetDatabase	BadItems: $BadItemsEncountered";			
			} ELSEIF ($Status -eq "Suspended")  {
				Write-Host "$MailboxAlias - Suspended for $timeSuspended";
			} ELSEIF ($Status -eq "Queued") {
				Write-Host "$MailboxAlias - Queued for $timeQueued";			
			} ELSE {
				Write-Warning "$MailboxAlias - $Status";
			}

			# Once the mailbox has finished moving
			IF ($Status -eq "Completed") {						
				Write-Host "Finished moving $DisplayName";
				IF ($BadItemsEncountered -ne 0) { 	# IF we encountered errors, create a log report.
					Write-Warning "Found $BadItemsEncountered BadItems on $TermUser! Creating log to debug.";
					$BodyLog = Get-MoveRequestStatistics -Identity $MailboxIdentity -IncludeReport | Format-List
					Write-Warning "Clearing job with bad items."
					Remove-MoveRequest -Identity $MailboxIdentity -Confirm:$false;
				} ELSE {
					Write-Host "Clearing successful job."				
					Remove-MoveRequest -Identity $MailboxIdentity -Confirm:$false;
					}
				# Sending Email to Techs
				Write-Host "Sending email notification to $DisplayName that their mailbox is now online."
				$EmailAddress = "ittechs@domain.local"
				$Subject = "Mailbox move for $DisplayName finished";
				$Body = "

The Mailbox has finished moving.

Mailbox Idenity   : $DisplayName
Time in queue     :	$timeQueued
Time to move data :	$duration
Your Mailbox Size :	$totalMailboxSize
Total Items Moved :	$itemCount
Corrupt Items     :	$BadItemsEncountered

If Corrupted Items were encounter please review below:

$BodyLog

";
				$SMTP = New-Object Net.Mail.SmtpClient($SMTPServer);
				$SMTP.Send($EmailFrom, $EmailAddress, $Subject, $Body);
				# Sending Email to Mailbox
				$EmailAddress = (Get-Mailbox -Identity $MailboxIdentity | Select-Object PrimarySmtpAddress).PrimarySmtpAddress
				$SMTP.Send($EmailFrom, $EmailAddress, $Subject, $Body);
			}
		}
		Write-Host "------Waiting for 60 seconds------"
		Start-Sleep 60;
	}

}

# Start the Monitor & Moves.
# Note you need to write the resume function!
Function MonitorEmailMove {
	Do {
		Process-Mailboxes;
		Write-Host "------No Suspended Mailboxes, waiting 60 seconds------"
		Start-Sleep 60;
	} Until ($runContinuously -eq $false)
}

Function RemoveMobile {
Clear-Host
	Write-Warning "Removing NOT Wiping Mobile Devices"
	$ASDevices = (Get-ActiveSyncDevice -mailbox $TermUser | Format-Table Identity)
	ForEach ($device in $ASDevices)
	{Remove-ActiveSyncDevice -identity $device -Verbose}
	}

# We no longer wipe mobile devices. This is just included in the event we need to start doing it again.
Function WipeMobile {
Clear-Host
	Write-Warning "Remember, this DOES NOT WIPE BLACKBERRY DEVICES"
	Write-Warning "The Error that follows means there are not Active Sync Devices `
	Cannot bind argument to parameter 'Identity' because it is null."
	Sleep -Seconds "7"
	$ASDevices = (Get-ActiveSyncDevice -mailbox $TermUser | Format-Table Identity)
	ForEach ($device in $ASDevices)
	{Clear-ActiveSyncDevice -Identity $device -NotIFicationEmailAddresses "ITTechs@domain.local" -Verbose}
	}

# We are not exporting mailboxes to a PST. Leaving this here in the event we need/want to.
Function ExportPST {
Clear-Host
	Write-Warning "Depending on the amount of email this might take a while."
	Sleep -Seconds "5"
	$ExportTo = Read-Host "Enter the path for the pst file, must be a UNC path, e.g. \\ServerFSIT\Misc\User.PST:"
	New-MailboxExportRequest -Mailbox $TermUser -FilePath $ExportTo
	Write-Warning "Type CheckPST to check status of mailbox export"
	}

Function CheckPST {
	Get-MailboxExportRequest | Get-MailboxExportRequestStatistics
	}

<#Does not work - We are no longer transfering SMTPs any way. Just would be nice to have working
Function SMTPTransfer {
	$SMTPTransferID = Read-Host "Enter user's alias/opID to transfer the SMTPs to:"
	$STMPs = ($Emails.EmailAddresses | Select-String -CaseSensitive "smtp") -replace "smtp:","" 
	$SMTPTransfterUser = Get-Mailbox -Identity $SMTPTransferID
	# $Emails.EmailAddresses = ($Emails.EmailAddresses | Select-String -CaseSensitive "smtp") -replace "smtp:",""
	$Emails.EmailAddresses += $STMP
	Set-Mailbox -Identity $SMTPTransferID -EmailAddresses $Emails.EmailAddresses -WhatIF
#>

<#I HAVE NO IDEA HOW TO LOG THIS. SO LET'S START WITH XML.#>

Function CreateXMLTemplate {
	# create a template XML to hold data
	$Template = @'
	<Termination version='1.0'>
		<Users>
			<OPID></OPID>
			<Name></Name>
			<Description></Description>
			<Email></Email>
			<Manager></Manager>
		</Users>
	</Termination>

'@
# We need to find a home for this file.
$Template | Out-File $home\users.xml -Encoding UTF8
	}

Function SetupLogging {
Clear-Host
	# Create XML Object for logging
	$xml = New-Object xml
	# Load XML File
	$xml.Load("$home\users.xml") #Needs to be chanaged once we have a home for this file
	# Insert XML Data into XML Object
	$Log = (@($xml.Termination.Users)[0])
	# Create New Log Entries
	$NewLog = $Log.Clone()
	}

Function SaveLog {
Clear-Host
	# Save Data to Memory
	$NewLog.OPID = $TermUser
	$NewLog.Name = $ADTermUserDisplayName
	$NewLog.Description = $ADTermUser.Description
	$NewLog.Email = $Emails.PrimarySmtpAddress
	$NewLog.Manager = $ADTermUserManagerDisplayName
	# Append Data to Log
	$xml.Termination.AppendChild($NewLog)
	# Save Log
	$xml.Save("$home\users.xml")
	}

#In the words of The Donald
#YoureFired
