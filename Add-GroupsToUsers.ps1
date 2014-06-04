Function Generate-Password {
	$Password = $null
	$Random = $null
	#Set up random number generator
	$Random = New-Object System.Random
	#Generate a new 10 character password
	1..10 | ForEach { $Password = $Password + [char]$Random.next(33,127) }
	$Password
}
$Users = Import-Csv C:\Users\hsmalley\Desktop\accounts.csv
$Users | Foreach {
	$User = Get-QADUser -Identity $_.IDS
	$Account = Get-QADGroup -Identity $_.Accounts
	$UserNames = $User.Name -split ', '
	$Description = $_.Description
	$Domain_Users = Get-QADGroup -Identity "Domain Users"
	$External_Users = Get-QADGroup -Identity "External_Users"
	$Extranet_Users = Get-QADGroup -Identity "Extranet_Users"
	$Subscriber_Users = Get-QADGroup -Identity "Subscriber Users"
	Add-QADMemberOf -Identity $User -Group $Account
	Add-QADMemberOf -Identity $User -Group $External_Users
	Add-QADMemberOf -Identity $User -Group $Extranet_Users
	Add-QADMemberOf -Identity $User -Group $Subscriber_Users
	Set-QADUser -Identity $User -Description $Description -DisplayName $User.Name -FirstName $UserNames[1] -LastName $UserNames[0] -objectAttributes @{PrimaryGroupID='8311'} -Confirm:$false -UserPassword (Generate-Password)
	Remove-QADGroupMember -Identity $Domain_Users -Member $User
	Enable-ADAccount -Identity $_.IDS -Confirm:$false
}
