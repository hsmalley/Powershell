Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Setup Email fields
$SMTP = New-Object Net.Mail.SMTPClient
$SMTP.Host = "MAILSERVER.DOMAIN.LOCAL"
$SMTP.TargetName = "MAILSERVER.DOMAIN.LOCAL"
$EmailFrom = "ITTechs@DOMAIN.LOCAL"
$EmailTO = "HR@DOMAIN.LOCAL"
$Subject = "AD Title and Manager Report - $Env:COMPUTERNAME"
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
$Body = @()
Get-QADUser -SizeLimit 100000 | ForEach-Object {
	$User = $_
	IF ($User.ParentContainer -eq "DOMAIN.LOCAL/Users"){
		$UserName = $User.Name
		$Title = $User.Title
		IF ($User.Manager -ieq $null) {$Manager = "No Manager"}
		ELSE {$Manager = (Get-QADUser -Identity $User.Manager).Name}
		$Description = $User.Description
		$Log =  New-Object PSObject
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "User Name" -Value $UserName -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Manager" -Value $Manager -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Title" -Value $Title -Force
		$Body += $Log
		}
	}
# Convert Content to HTML and email report.
$message.Body = $Body | ConvertTo-Html -Head $Style
$SMTP.Send($Message)
