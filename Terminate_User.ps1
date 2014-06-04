<#
Terminate User

	*	Remove mobile device					*	Set Out Of Office
	*	Log terminiation details				*	Hide in GAL
	*	Disable account							*	Change account password
	*	Forward Email to Manager				*	Move account to Users - Disabled
	*	Change primary group to Disabled Users	*	Remove groups
	*	Remove the following account details:	*	Put termination date in description
		*	Phone	*	Office					*	Logging - NOW WITH CSV & XML!
		*	Mobile	*	Title
		*	Fax		*	Description

This Script Requires the Quest Active Directory cmdlets.

Change Log:

WIP:
	Merge XML logs
	GUI - Nice to have but not needed
#>
Clear-Host
# Get User/Pass
$Cred = Get-Credential
# Add Quest CMDLETS
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Connect to DC
Connect-QADService -Service 'DC1.DOMAIN.LOCAL' -Credential $Cred
# Needed for Loops
$RunContinuously = $true
# Get the User
$TermUser = Read-Host "Enter Username of User to be Terminated :>"
# What Exchange Server
#$ExchangeServer = Read-Host "What Exchange Server do you want to connect to? <EXCH1/EXCH2>: "
$ExchangeServer = "EXCH1"
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
$Password = Generate-Password
# Get Timestamp
$Date = Get-Date
# Create Email Object to Send Emails
$SMTP = New-Object Net.Mail.SMTPClient
$SMTP.Host = "MAILSERVER.DOMAIN.LOCAL"
$SMTP.TargetName = "MAILSERVER.DOMAIN.LOCAL"
<# If you need to authenticate with your SMTP
$Credentials = New-Object System.Net.NetworkCredential
$Credentials.UserName = $Cred.UserName
$Credentials.Password = $Cred.GetNetworkCredential()
$SMTP.Credentials = $Credentials #>
# Convert OPID to User Account information for Quest AD Tools
$ADTermUser = Get-QADUser -Identity $TermUser -Credential $Cred
# Move User to the Disabled User OU
Move-QADObject -Identity $ADTermUser -NewParentContainer (Get-QADObject "Users - Disabled") -Credential $Cred
# Refresh ADTermUser
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
	IF (($SetManager -ieq "N") -or ($SetManager -ieq "NO")) {
		# Get the Correct Manager's Name
		$TermUserManager = Read-Host "Enter the Managers opID:"
		# Set the Term User's Manager.
		$ADTermUserManager = Get-QADObject -Identity $TermUserManager -Credential $Cred
		# Try to correct possible human error.
		$TermUserManager = $ADTermUserManager.LogonName
		# Get the Manger's Actual Name.
		$ADTermUserManagerDisplayName = $ADTermUserManager.DisplayName
		# Tell the Tech who the manager is going to be now.
		Write-Warning "Setting Data Manager to $TermUserManager - $ADTermUserManagerDisplayName" -WarningAction Inquire
		# Set the new manager on the User's account
		Set-QADUser -Identity $ADTermUser -Manager $ADTermUserManager -Credential $Cred
		# Refresh the Account Variable
		$ADTermUser = Get-QADUser -Identity $TermUser -Credential $Cred
	}
#Set Expire Date
$Expire = $Date.AddMonths(13)

<# START CSV LOG #>
# Create Object to hold data
$Log =  New-Object PSObject
# Log the Date
Add-Member -MemberType NoteProperty -InputObject $Log -Name "Run Date" -Value $Date -Force
Add-Member -MemberType NoteProperty -InputObject $Log -Name "Remove After" -Value $Expire -Force
# Log User Details
Add-Member -MemberType NoteProperty -InputObject $Log -Name "User ID" -Value $TermUser -Force
Add-Member -MemberType NoteProperty -InputObject $Log -Name "User Name" -Value $ADTermUserDisplayName -Force
#Add-Member -MemberType NoteProperty -InputObject $Log -Name "User Password" -Value $Password -Force
# Log the Manager
Add-Member -MemberType NoteProperty -InputObject $Log -Name "Manager ID" -Value $TermUserManager -Force
Add-Member -MemberType NoteProperty -InputObject $Log -Name "Manager" -Value $ADTermUserManagerDisplayName -Force
# Log the Account Used for Termination
Add-Member -MemberType NoteProperty -InputObject $Log -Name "Tech" -Value $Cred.UserName -Force
# Initialize the array
$TermLog = @()
# Attach Object to Array
$TermLog += $Log
# Save Array to File
$LogPath = "\\FILESERVER\ITSHARE\User Termination Log\TermLog.csv"
$TermLog | Export-Csv -Append -Force -NoTypeInformation $LogPath
<# END CSV LOG #>

<# START XML LOG #>
# Create XML Document
$XMLDocument = New-Object System.XML.xmlDataDocument
# Make Root Tag
$XMLRoot = $XMLDocument.CreateElement('Terminations')
# Get Rid of the console output
[Void]$XMLDocument.AppendChild($XMLRoot)
	# Setup XML Fields
	$XMLLog = $XMLDocument.CreateElement('TermUser')
	# Log the Dates
	$XMLLog.SetAttribute('TermDate',$XMLLog.TermDate)
	$XMLLog.TermDate = $Date.ToString()
	$XMLLog.SetAttribute('RemoveDate',$XMLLog.RemoveDate)
	$XMLLog.RemoveDate = $Expire.ToString()
	# Log User
	$XMLLog.SetAttribute('User',$XMLLog.User)
	$XMLLog.User = $ADTermUserDisplayName
	$XMLLog.SetAttribute('UserID',$XMLLog.UserID)
	$XMLLog.UserID = $TermUser
	$XMLLog.SetAttribute('Password',$UserDetails.Password)
	$XMLLog.Password = $Password
	# Log Manager
	$XMLLog.SetAttribute('Manager',$XMLLog.Manager)
	$XMLLog.Manager = $ADTermUserManagerDisplayName
	$XMLLog.SetAttribute('ManagerID',$XMLLog.ManagerID)
	$XMLLog.ManagerID = $TermUserManager
	$XMLLog.SetAttribute('ManagerOffice',$XMLLog.ManagerOffice)
	$XMLLog.ManagerOffice = $ADTermUserManager.Office
	$XMLLog.SetAttribute('ManagerTitle',$XMLLog.ManagerTitle)
	$XMLLog.ManagerTitle = $ADTermUserManager.Title
	$XMLLog.SetAttribute('ManagerPhone',$XMLLog.ManagerPhone)
	$XMLLog.ManagerPhone = $ADTermUserManager.PhoneNumber
	# Log Tech
	$XMLLog.SetAttribute('Tech',$XMLLog.Tech)
	$XMLLog.Tech = $Cred.UserName
 	[void]$XMLRoot.AppendChild($XMLLog)
# And, lastly save the XML Log
$Path = "\\FILESERVER\ITSHARE\User Termination Log"
$File = $ADTermUser.Sid.Value + ".XML"
$Log = $Path + "\XML\" + $File
$XMLDocument.Save($Log)
<# END XML LOG #>
# Disable AD Account
Disable-QADUser -Identity $ADTermUser -Credential $Cred
# Add to Disabled User Group
Add-QADGroupMember -Identity "Disabled Users" -Member $ADTermUser -Credential $Cred
# Set Description
$Description = ($AdtermUser.Description + " - Terminated On: $Date")
# Removes Title, Office, Phones, Fax, changes password to random number, set description, and change primary group to Disabled Users
Set-QADUser -Identity $ADTermUser -Credential $Cred -Title '' -Office '' -PhoneNumber '' -MobilePhone '' -Pager '' -Fax ''  -Description $Description
Set-QADUser -Identity $ADTermUser -Credential $Cred -UserMustChangePassword $true -UserPassword $Password
Set-QADUser -Identity $ADTermUser -Credential $Cred -AccountExpires $Expire
Set-QADUser -Identity $ADTermUser -Credential $Cred -objectAttributes @{PrimaryGroupID='19791'}
# Remove Groups
$ADTermUser.MemberOf | Remove-QADGroupMember -Member $ADTermUser -Credential $Cred
# Set Exchange Server
IF ($ExchangeServer -ieq "Exch1") {$ExchangeServer = "http://exch1/PowerShell/"}
IF ($ExchangeServer -ieq "Exch2") {$ExchangeServer = "http://exch2/PowerShell/"}
#Connect to Exchange Server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeServer -Authentication Kerberos -Credential $Cred
Import-PSSession $Session -AllowClobber
# Get the user's mailbox
$Emails = Get-Mailbox -Identity $TermUser
# Hide the user from the GAL and set forwarding of the user's mail to the manager.
Set-Mailbox -Identity $TermUser -hiddenfromaddresslistsenabled $True -ForwardingAddress $TermUserManager
# Set out of office reply
Set-MailboxAutoReplyConfiguration -Identity $TermUser -autoreplystate enabled -InternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with THE COMPANY." -ExternalMessage "Thank you for your email, unfortunately the person you have contacted is no longer with THE COMPANY, however your email has been forwarded to someone who can help you."
# Give Manager Full Access
Add-MailboxPermission $TermUser -User $TermUserManager -AccessRights fullaccess
# Punches their mail clients and phones in the face.
Set-CASMailbox -Identity $TermUser -OwaEnabled $false -EcpEnabled $false -EwsEnabled $false -MapiEnabled $false -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false
Set-CASMailbox -Identity $TermUser -MapiBlockOutlookRpcHttp $true -EwsAllowMacOutlook $false -EwsAllowOutlook $false -EwsAllowEntourage $false
# Continues flogging their phones. The switch for -ActiveSyncBlockedDeviceIDs doesn't quite work yet....
#Get-ActiveSyncDeviceStatistics –Mailbox $TermUser | select DeviceID | ForEach-Object {Set-CASMailbox -Identity $TermUser -ActiveSyncBlockedDeviceIDs $_}
# Banishes their phones.
Get-ActivesyncDevice -Mailbox $TermUser | Remove-ActiveSyncDevice
<# We no longer wipe mobile devices. This is just included in the event we need to start doing it again.
$ASDevices = (Get-ActiveSyncDevice -mailbox $TermUser | Format-Table Identity)
ForEach ($device in $ASDevices)
	{Clear-ActiveSyncDevice -Identity $device -NotIFicationEmailAddresses "ITTechs@DOMAIN.LOCAL" -Verbose}
#>
<# We are not exporting mailboxes to a PST. Leaving this here in the event we need/want to.
$ExportTo = Read-Host "Enter the path for the pst file, must be a UNC path, e.g. \\Sever\Share\User.PST:"
New-MailboxExportRequest -Mailbox $TermUser -FilePath $ExportTo
Write-Warning "Type CheckPST to check status of mailbox export"
Function CheckPST {Get-MailboxExportRequest | Get-MailboxExportRequestStatistics}
#>
# Start Moving The Mailbox
New-MoveRequest -Identity $TermUser -BadItemLimit 5 -SuspendWhenReadyToComplete:$true -TargetDatabase "Canada"

# Start the Monitor & Watch The Moves.
Function MonitorEmailMove {
	Do {Process-Mailboxes
		Write-Host "------No Suspended Mailboxes, waiting 60 seconds------"
		Start-Sleep 60
	} Until ($RunContinuously -eq $false)
}

# Monitor the Move Requests
Function Process-Mailboxes {
 	While ($MoveRequests = Get-MoveRequest) {
		foreach($MoveRequest in $MoveRequests) {
			$MailboxIdentity = $MoveRequest.Identity
			$TargetDatabase = $MoveRequest.TargetDatabase
			$MailboxAlias = $MoveRequest.Alias
			$DisplayName = $MoveRequest.DisplayName
			$Status = $MoveRequest.Status

			$Results = Get-MoveRequestStatistics -Identity $MailboxIdentity  | Select DisplayName, PercentComplete, BadItemsEncountered, TotalMailboxSize, TotalMailboxItemCount, TotalInProgressDuration,TotalSuspendedDuration,TotalQueuedDuration
			$PercentComplete = $Results.PercentComplete
			$BadItemsEncountered = $Results.BadItemsEncountered
			$TotalMailboxSize = $Results.TotalMailboxSize
			$Duration = $Results.TotalInProgressDuration
			$ItemCount = $Results.TotalMailboxItemCount
			$TimeSuspended = $Results.TotalSuspendedDuration
			$TimeQueued = $Results.TotalQueuedDuration

			IF (($Status -eq "InProgress") -or ($Status -eq "CompletionInProgress") -or ($Status -eq "Completed")) {
				Write-Host "$MailboxAlias	$percentComplete%	duration: $Duration	Target: $TargetDatabase	BadItems: $BadItemsEncountered"
			} ELSEIF ($Status -eq "Suspended")  {
				Write-Host "$MailboxAlias - Suspended for $timeSuspended"
			} ELSEIF ($Status -eq "Queued") {
				Write-Host "$MailboxAlias - Queued for $timeQueued"
			} ELSE {
				Write-Warning "$MailboxAlias - $Status"
			}

			# Once the mailbox is ready to complete
			IF ($Status -eq "AutoSuspended"){
			Write-Host "Completing $DisplayName"
			$MoveRequest | Resume-MoveRequest -Confirm:$false}

			# Once the mailbox has finished moving
			IF ($Status -eq "Completed") {
				Write-Host "Finished moving $DisplayName"
				IF ($BadItemsEncountered -ne 0) { 	# IF we encountered errors, create a log report.
					Write-Warning "Found $BadItemsEncountered BadItems on $TermUser! Creating log to debug."
					$BodyLog = Get-MoveRequestStatistics -Identity $MailboxIdentity -IncludeReport | Format-List
					Write-Warning "Clearing job with bad items."
					Remove-MoveRequest -Identity $MailboxIdentity -Confirm:$false
				} ELSE {
					Write-Host "Clearing successful job."
					Remove-MoveRequest -Identity $MailboxIdentity -Confirm:$false
					}
				# Sending Email to Techs
				Write-Host "Sending email notification to $DisplayName that their mailbox is now online."
				$EmailAddress = "ittechs@DOMAIN.LOCAL"
				$Subject = "Mailbox move for $DisplayName finished"
				$Body = "
						The Mailbox has finished moving.

						Mailbox Idenity   : $DisplayName
						Time in queue :	$timeQueued
						Time to move data :	$Duration
						Your Mailbox Size :	$totalMailboxSize
						Total Items Moved :	$itemCount
						Corrupt Items :	$BadItemsEncountered

						If Corrupted Items were encounter please review below:

						$BodyLog
						"
				$SMTP.Send("ITTechs@DOMAIN.LOCAL", "ITTechs@DOMAIN.LOCAL", $Subject, $Body)
				# Sending Email to Mailbox
				$EmailAddress = (Get-Mailbox -Identity $MailboxIdentity | Select-Object PrimarySmtpAddress).PrimarySmtpAddress
				$SMTP.Send("ITTechs@DOMAIN.LOCAL", $EmailAddress, $Subject, $Body)
			}
		}
		IF ((Get-MoveRequest) -eq $null) {$RunContinuously = $false}
Write-Host "------Waiting for 60 seconds------"
		Start-Sleep 60
	}
}
$MonitorMoves = Read-Host "Do you wish to monitor the email moves? If you need to run this again type no. "
IF (($MonitorMoves -ieq "Yes") -or ($MonitorMoves -ieq "Y")) {MonitorEmailMove}
