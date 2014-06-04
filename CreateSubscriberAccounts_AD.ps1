Clear-Host

# Get User/Pass
$Cred = Get-Credential

# Add Quest CMDLETS
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Connect to DC
Connect-QADService -Service 'DC.DOMAIN.LOCAL' -Credential $Cred

# Logging - Create XML Document
$XMLDocument = New-Object System.XML.xmlDataDocument
# Logging - Make XML Root Tag
$XMLRoot = $XMLDocument.CreateElement('UserList')
# Logging - Get Rid of the console output
[Void]$XMLDocument.AppendChild($XMLRoot)

# Set OU
$OU = 'OU=Users,OU=Subscribers,OU=External - Users and Groups,DC=DOMAIN,DC=LOCAL'
# Get User data
$UserList = "C:\Users\hsmalley\Desktop\UserList.csv"

#Import User Data
Import-Csv $UserList | Where-Object {

	# Clear vars on each pass
	Clear-Variable Email
	Clear-Variable sAMAccountName
	Clear-Variable FirstName
	Clear-Variable LastName
	Clear-Variable DisplayName
	Clear-Variable Description
	Clear-Variable Account
	Clear-Variable Account2

	# Setup vars and error control
	IF ($_.sAMAccountName -ne $null) {$sAMAccountName = $_.sAMAccountName}
	IF ($_.sAMAccountName -ne $null) {$DisplayName = $sAMAccountName}
	IF ($_.Email -ne $null) {$Email = $_.Email}
	IF ($_.First -ne $null) {$FirstName = $_.First}
	IF ($_.Last -ne $null) {$LastName = $_.Last}
	IF (($_.Last -ne $null) -and ($_.First -ne $null)) {$DisplayName = "$LastName, $FirstName"}
	IF ($_.Account -eq $null) {$Account = $null}
	IF ($_.Account2 -eq $null) {$Account = $null}
	IF ($_.Account -ne $null) {$Account = $_.Account
		$Description =  "Subscriber - Account" + " "
		$Description = $Description + $Account
		}
	IF ($_.Account2 -ne $null) {$Account2 = $_.Account2
		$Description = $Description + "," + " "
		$Description = $Description + $Account2
		}
	# Create account if needed, disable account, and copy groups from template.
	IF ((Get-QADUser $_.sAMAccountName) -eq $null) {
		New-QADUser -ParentContainer $OU -Name $DisplayName -sAMAccountName $sAMAccountName -UserPassword "password" | Disable-QADUser
		(Get-QADUser Subscriber).MemberOf | Add-QADGroupMember -Member $sAMAccountName
        }

	# Correct account details if needed.
	IF (((Get-QADUser $sAMAccountName) -ne $null) -and ((Get-QADUser $sAMAccountName).DisplayName -ne $null)) {Set-QADUser -Identity $sAMAccountName -FirstName $FirstName -LastName $LastName -DisplayName $DisplayName}
	IF (((Get-QADUser $sAMAccountName) -ne $null) -and ((Get-QADUser $sAMAccountName).Email -eq $null)) {Set-QADUser -Identity $sAMAccountName -Email $Email}
	IF (((Get-QADUser $sAMAccountName) -ne $null) -and ((Get-QADUser $sAMAccountName).Description -eq $null)) {Set-QADUser -Identity $sAMAccountName -Description $Description}
}

# Run report on users and save to XML file
$Users = Get-QADUser -SearchRoot $OU
$Users | ForEach-Object {
	Clear-Variable User
	$User = $_
	# Setup XML Fields
    $XMLLog = $XMLDocument.CreateElement('Users')
	$XMLLog.SetAttribute('AccountName',$XMLLog.AccountName)
    $XMLLog.AccountName = $User.SamAccountName
	$XMLLog.SetAttribute('Email',$XMLLog.Email)
    $XMLLog.Email = $User.Email
    $XMLLog.SetAttribute('FirstName',$XMLLog.FirstName)
    $XMLLog.FirstName = $User.FirstName
    $XMLLog.SetAttribute('LastName',$XMLLog.LastName)
    $XMLLog.LastName = $User.LastName
    $XMLLog.SetAttribute('Description',$XMLLog.Description)
    $XMLLog.Description = $User.Description
	[void]$XMLRoot.AppendChild($XMLLog)
}
#SAVE LOG
$Log = "C:\users\hsmalley\desktop\UserList.XML"
$XMLDocument.Save($Log)
