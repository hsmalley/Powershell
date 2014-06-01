<#
This script compares DLs to users. If the user is in the wrong DL for their assigned office an email is generated so it can be corrected.
This script needs a CSV files formated with the office names and DLs. Please update paths if moving files.
#>
$BadDNList = @() #Create Blank Array
$DLs = Import-Csv Office_DLs.csv #Import list of Office DL's
$DLs | Foreach { #Process the DLs.
	#Set meaningful variables
	$line = $_
	$Office = $Line.Office
	$AD_DL = $Line.DL
	#Now let's get to work.
	$UsersInOffice_DN = Get-ADUser -SearchBase "OU=Users - US,DC=DOMAIN,DC=LOCAL" -LDAPFilter "(physicalDeliveryOfficeName=$Office)" #Get the users
		$UsersInOffice_DN | ForEach { #Checks each user
			$User = $_
		    $objEntry = [adsi]("LDAP://"+$User)
    		$Member = $objEntry.memberOf | where { $_ -match $AD_DL}
    		IF ((-not($Member)) {$BadDNList+=$User} #Now we got'em!
    		}
	}
$BadUsers = $BadDNList | Select -Property Name | Out-String
$SMTP = New-Object Net.Mail.SMTPClient
$SMTP.Host = "mailserver.domain.local"
$SMTP.TargetName = "mailserver.domain.local"
$Subject = "User's in the wrong DL"
$Body = "The Script - Compare Office to DL - has detected the following user with the wrong Office DL:

$BadUsers

This script is running on: $Env:COMPUTERNAME"
$SMTP.Send("ITTechs@DOMAIN.LOCAL", "ITTechs@DOMAIN.LOCAL", $Subject, $Body)
