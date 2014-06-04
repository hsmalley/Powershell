#	This script will set up AD, User Directories, Exchange and Lync, using basic data retrived from a CSV file
# 	Requires the Active Directory module for Windows Powershell and appropriate credentials
#	CSV file with corresponding header and user(s) info:
#		Office,UserName,FirstName,LastName,Initial,Department,Role,Title,SetupSameAs,Manager,MailboxAccess,Extension,Mobile


#LOAD POWERSHELL SESSIONS
#------------------------
$exchangeserver = "exchange1.domain.com"
$Lyncserver = "lync1.domain.com.au"
$DC = "dc1.domain.com.au"
$AdminEmail = "administrator@domain.com.au"
$ScriptUser = $env:username
$usercredential= get-credential -credential domain\AdminUser

cls
write-host -foregroundcolor Green "Loading modules for AD, Exchange and Lync..."
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeserver/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PsSession $exchangesession
$lyncsession = new-pssession -connectionuri https://$Lyncserver/ocspowershell -credential $usercredential
Import-PSSession $lyncsession
import-module ActiveDirectory


#VARIABLES
#---------
#Get Script User's Name
$ScriptUserDetails = Get-ADUser $ScriptUser -Server $DC -properties givenName | select-object givenName
	$ScriptUserName = $ScriptUserDetails.givenName
	
#Email Communications
$EmailCC1 = "user1@domain.com.au"
$EmailCC2 = "user2@domain.com.au"

#LOAD USER DATA FROM CSV
#-----------------------
$InputPath = "\\domain.com.au\Scripts\Data\"  
$InputFile = $InputPath + "NewUserInfo.csv"  
Invoke-Item $inputFile
cls
write-host -foregroundcolor Green "Update user details and save CSV file. Press any key to continue"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")


#DEFINE VARIABLES
#----------------
$inputFile = Import-CSV  $inputFile 
foreach($line in $inputFile) 

	{     
	
	#Retrieve user details from CSV file
	$office = $line.Office
	$username = $line.UserName
	$firstname = $line.FirstName
	$lastname = $line.LastName
	$initial = $line.Initial
	$department = $line.Department
	$Role = $line.Role
	$title = $line.Title
	$SetupSameAs = $line.SetupSameAs
	$manager = $line.Manager
	$MailboxAccess = $line.MailboxAccess
	$Extension = $line.Extension
	$Extension = $Extension.substring($Extension.length -3, 3)
	$Mobile = $line.Mobile
	$info = ""

	cls
	
	#Prompt for any missing details
	if ($username -eq "") {$username = read-host -prompt "Please Enter Unique Username"}
	if ($office -eq "") {$office = read-host -prompt "Please Enter Office for"}
	if ($initial -eq "") {$initial = read-host -prompt "Please Enter Unique User Initials"}
	if ($department -eq "") {$department = read-host -prompt "Please Enter Department"}
	if ($Role -eq "") {$Role = read-host -prompt "Please Enter Role"}
	if ($title -eq "") {$title = read-host -prompt "Please Enter Title"}
	if ($SetupSameAs -eq "") {$SetupSameAs = read-host -prompt "Please Enter SetupSameAs"}
	if ($manager -eq "") {$manager = read-host -prompt "Please Enter thier Manager's Username"}
	if ($Extension -eq "") {$Extension = read-host -prompt "Please Enter 3 Digit Phone Extension"}
	#if ($Mobile -eq "") {$Mobile = read-host -prompt "Please Enter Mobile in (+61) 405 123 123 format"}

	#Get 'SetupSameAs' Details
	$SameAsDetails = Get-ADUser $SetupSameAs -Server $DC -properties Office,DisplayName,sAMAccountName | select-object Office,DisplayName,sAMAccountName
		$SameAsOffice = $SameAsDetails.Office
		$SameAsDisplayName = $SameAsDetails.DisplayName
		
	#Get Script User's Details
	$ScriptUserDetails = Get-ADUser $ScriptUser -Server $DC -properties mail | select-object mail
		$ScriptUserMail = $ScriptUserDetails.mail
		
	
	#Define Fixed Variables
	$umpolicy="UMDial Default Policy"
	$userou="OU=USERS,DC=domain,DC=com,DC=au"
	$companyname="My Company"
	$mailboxdatabase="MailDB1"
	$sipdomain="domain.com.au"
	$Country = "Australia"
	$DisplayName = $firstname +" " +$lastname
	$SharepointPage = "http://intranet/my/Person.aspx?accountname=domain\" +$username
	$HomePath = "\\fileserver\home$\" +$username
	$LogFile = $InputPath+"SetupLog_" +$username +".txt"
	
	#Define Generated and Localised Variables
	$Name=$Firstname+" "+$Lastname
	$accountpassword = read-host -assecurestring -prompt "Please enter temporary password for new user"
	$upn = $username+ "@domain.com.au"
	$email = $Firstname+"."+$Lastname +"@domain.com.au"

	if ($office -eq "City1") 
		{
		$EnableVoice = $True
		$Lyncserver="lync1.domain.com.au"
		$Telephone = "(+61) 3 1234 5" +$Extension
		$Street = "123 My Street"
		$City = "City"
		$State = "State"
		$PostCode = "3000"
		$ProfilePath = "\\fileserver\profiles\" +$username
		$Fax = "(+61) 3 1234 5678 "
		$IPphone = "3" +$Extension
		$teluri = "tel:+612345" +$Extension +";ext=" +$IPphone
		$info = "random text about user"
		$Operator = "1000"
		}
	elseif ($office -eq "City2") 
		{
		$EnableVoice = $True
		$Lyncserver="lync2.domain.com.au"
		$Telephone = "(+61) 3 1234 5" +$Extension
		$Street = "123 My Street"
		$City = "City"
		$State = "State"
		$PostCode = "3000"
		$ProfilePath = "\\fileserver\profiles\" +$username
		$Fax = "(+61) 3 1234 5678 "
		$IPphone = "3" +$Extension
		$teluri = "tel:+612345" +$Extension +";ext=" +$IPphone
		$info = "random text about user"
		$Operator = "1000"
		}
	elseif ($office -eq "City3") 
		{
		$EnableVoice = $True
		$Lyncserver="lync3.domain.com.au"
		$Telephone = "(+61) 3 1234 5" +$Extension
		$Street = "123 My Street"
		$City = "City"
		$State = "State"
		$PostCode = "3000"
		$ProfilePath = "\\fileserver\profiles\" +$username
		$Fax = "(+61) 3 1234 5678 "
		$IPphone = "3" +$Extension
		$teluri = "tel:+612345" +$Extension +";ext=" +$IPphone
		$info = "random text about user"
		$Operator = "1000"
		}
	elseif ($office -eq "City4") 
		{
		$EnableVoice = $True
		$Lyncserver="lync4.domain.com.au"
		$Telephone = "(+61) 3 1234 5" +$Extension
		$Street = "123 My Street"
		$City = "City"
		$State = "State"
		$PostCode = "3000"
		$ProfilePath = "\\fileserver\profiles\" +$username
		$Fax = "(+61) 3 1234 5678 "
		$IPphone = "3" +$Extension
		$teluri = "tel:+612345" +$Extension +";ext=" +$IPphone
		$info = "random text about user"
		$Operator = "1000"
		}
	else
		{
		write-host -foregroundcolor Red "Office not recognised. Quitting..."
		exit
		}
	
	
	#Optional/Future
	#	$archivedatabase="Your Database holding online archives"
	#	$retentionpolicy="Your retention policy"
	#	$dialplan="Lync Voice Dial Plan"
	#	$voicepolicy="Lync Voice Policy"
	#	$locationpolicy="Lync Location Policy"
	#	$externalaccesspolicy="Lync External Access Policy"


	Start-Transcript -path $LogFile -append
	cls
	write-host -foregroundcolor Green "New user setup starting for:" $firstname $lastname
	write-host "`r"
	
	#SETUP EXCHANGE
	#--------------
 
	#Create user and enable mailbox
	New-Mailbox -DomainController $DC -name $name -userprincipalname $upn -Alias $username -OrganizationalUnit $userou -SamAccountName $username -FirstName $FirstName -Initials $initial -LastName $LastName -Password $accountpassword -ResetPasswordOnNextLogon $true -Database $mailboxdatabase | out-null

		#OPTIONAL/FUTURE:	-Archive -ArchiveDatabase $archivedatabase -RetentionPolicy $retentionpolicy

	#pause for Exchange 
	write-host -foregroundcolor Green "New mailbox created - Pausing 10 seconds for Exchange changes"
	write-host "`r"
	Start-Sleep -s 10 

	#Update Notes info
	Get-Mailbox $username -DomainController $DC | Set-User -DomainController $DC -notes $info

	#Enable For Unified Messaging
	if ($EnableVoice -eq $true) 
	{
	Get-Mailbox $username -DomainController $DC | Enable-UMMailbox -DomainController $DC -ummailboxpolicy $umpolicy -sipresourceidentifier $email -extensions $IPphone
	Start-Sleep -s 5
	Get-Mailbox $username -DomainController $DC | Set-UMMailbox -OperatorNumber $Operator
	write-host -foregroundcolor Green "Unified Messaging Properties updated - Pausing 10 Seconds for Exchange changes"
	write-host "`r"
	Start-Sleep -s 10
	}	
	
	#Setup mailbox permissions
	if (!($MailboxAccess -eq "")) {Add-MailboxPermission -Identity $manager -User $username -AccessRights 'FullAccess'}
	
	
	#SETUP ACTIVE DIRECTORY
	#-----------------------

	#Update user properties
	Set-ADUser -server $DC -Identity $username -City $City -Company $companyname -Country "AU" -Department $department -description $title -DisplayName $DisplayName -Division $department -Fax $Fax -HomePage $SharepointPage -Manager $manager -Office $office -OfficePhone $Telephone -postalCode $PostCode -profilePath $ProfilePath -State $State -StreetAddress $Street -Title $title -add @{ipphone=$IPphone}
	
	#Update non-mandatory items
	if (!($Mobile -eq "")) {Set-ADUser -server $DC -Identity $username -Replace @{mobile=$Mobile}}
	write-host -foregroundcolor Green "Active Directory properties updated"
	write-host "`r"
	
	#Add to Groups (using SetupSameAs)
	$ds = new-object directoryServices.directorySearcher 
	$ds.filter = "(&(objectCategory=person)(objectClass=user)(samAccountName="+$SetupSameAs+"))" 
	$dn = $ds.findOne() 
	$user = [ADSI]$dn.path 
	foreach ($group in $user.memberof)
		{
		Add-ADGroupMember -Identity $group -Members $username
		#Trim group names
		$GroupTrim = $group.indexof(",OU=")
		if ($GroupTrim -gt 0)
			{
			$GroupName = $group.substring(0,$GroupTrim)
			$GroupName = $GroupName -replace ("CN="," `n")
			}
		$RPTGroups = $RPTGroups+" "+$GroupName
		}
	
	#pause for AD changes
	write-host -foregroundcolor Green "Active Directory groups and permissions update - Pausing 10 Seconds for all AD changes"
	write-host "`r"
	Start-Sleep -s 10

	
	
	#SETUP USER DIRECTORIES
	#----------------------
		
	# User profile
	$win7Profile = $ProfilePath +".V2"
	IF (!(TEST-PATH $win7Profile)) {NEW-ITEM $win7Profile -type Directory} 

	#Build Access Control Entry List
	$colRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
	$InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
	$PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None
	$objType =[System.Security.AccessControl.AccessControlType]::Allow 
	$objUser = New-Object System.Security.Principal.NTAccount("MyDomain\"+$username) 
	$objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 

	#Apply Permissions
	$objACL = Get-ACL $win7Profile 
	$objACL.AddAccessRule($objACE) 
	Set-ACL $win7Profile $objACL

	#Apply Ownership
	$objACL = Get-ACL $win7Profile 
	$objACL.SetOwner($objUser)
	Set-Acl -aclobject $objACL -path $win7Profile -passthru	
	
	# User Home Drive
	IF (!(TEST-PATH $HomePath)) {NEW-ITEM $HomePath -type Directory} 
	
	write-host -foregroundcolor Green "User Directories Created"
	write-host "`r"
	
	#SETUP LYNC
	#----------

	#enable for lync and configure settings
	Get-mailbox $username -DomainController $DC | Enable-csuser -DomainController $DC -registrarpool $lyncserver -sipaddresstype EmailAddress -sipdomain $sipdomain

	#pause for Lync changes
	write-host -foregroundcolor Green "User setup for Lync - Pausing 10 Seconds for Lync Changes"
	write-host "`r"
	Start-Sleep -s 10

	#Enable For Enterprise Voice
	if ($EnableVoice -eq $true) {Get-mailbox $username -DomainController $DC | Set-CSUser -DomainController $DC -enterprisevoiceenabled $True -lineuri $teluri}

		#OPTIONAL/FUTURE (if not <Default>): 
		#Get-mailbox $username | Grant-CSVoicePolicy -policyname $voicepolicy
		#Get-mailbox $username | Grant-CSDialPlan -policyname $dialplan
		#Get-mailbox $username | Grant-CSLocationPolicy -policyname $locationpolicy
		#Get-mailbox $username | Grant-CSExternalAccessPolicy -policyname $externalaccesspolicy


		
	#USER SUMMARY REPORT
	#-------------------
	cls
	write-host -foregroundcolor Green "Applying final user settings and generating report..."
	write-host "`r"
	Start-Sleep -s 10

	#Active Directory
	write-host -foregroundcolor Green "Active Directory Details"
	Get-ADUser $username -Server $DC
	write-host "`r"
	write-host "New user has been added to the following groups:" $RPTGroups
	write-host "`r"
	
	#Folders
	write-host -foregroundcolor Green "User Directories"
	write-host "`r"
	write-host "Profile Path is " $win7Profile
	write-host "Home Path is " $HomePath
	write-host "`r"
		
	#Exchange
	write-host -foregroundcolor Green "Exchange Details"
	Get-Mailbox $username -DomainController $DC
	
	if ($EnableVoice -eq $true) 
		{
		write-host -foregroundcolor Green "Unified Messaging Details"
		get-ummailbox $username -DomainController $DC 
		}
	
	#Lync
	write-host -foregroundcolor Green "Lync Details"
	Get-CSUser $username -DomainController $DC 

	
	#FINISH
	#------
	Stop-Transcript
	
	#Email IT report
	$MailSubject = "[AUTO] New account setup for " +$firstname+" "+$lastname
	$MailBody = "A new user account has been setup for " +$firstname+" "+$lastname+", by "+$ScriptUserName+" on "+(get-date)+". `r

"+$firstname+" was setup with the following information:
  Office: "+$office+"
  Username: "+$username+"
  First Name: "+$firstname+"
  Last Name: "+$lastname+"
  Initials: "+$initial+"
  Departsment: "+$department+"
  Role: "+$Role+"
  Title: "+$title+"
  Setup Same As: "+$SetupSameAs+"
  Manager: "+$manager+"
  Extension: "+$Extension+"
  Mobile: "+$Mobile+"
`r	
Please confirm these groups in AD (they have been copied from "+$SameAsDisplayName+", but may need to be changed) `n
"+$RPTGroups+"	
`r
IT to complete the following items:
 - Manual Steps 1 `r
 - Manual Steps 2 `r
 - Manual Steps 3 `r
`r
Regards
Admin Scripts"
	Send-MailMessage -To $ScriptUserMail -cc $AdminEmail,$EmailCC1,$EmailCC2  -From "admin.scripts@domain.com.au" -Subject $MailSubject -SmtpServer $exchangeserver -body $MailBody -attachment $LogFile



	#Pause for review, then load next line
	write-host "`r"
	write-host -foregroundcolor Green "User setup finished. Press any key to continue"
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	}

exit
