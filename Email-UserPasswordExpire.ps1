# Setup Email fields
$SMTP = New-Object Net.Mail.SMTPClient
$SMTP.Host = "MAILSERVER.DOMAIN.LOCAL"
$SMTP.TargetName = "MAILSERVER.DOMAIN.LOCAL"

# CSS Style for Email.
$Style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$Style = $Style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$Style = $Style + "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
$Style = $Style + "TD{border: 1px solid black; padding: 5px;}"
$Style = $Style + "</style>"

# All Users
$UserList = Get-QADUser -SizeLimit "100000" -Disabled:$false -SearchRoot "OU=Users - US,DC=DOMAIN,DC=LOCAL" -PasswordNotChangedFor "80"
# DL - Laptop Mobile Users
$UserList | ForEach-Object {
	$User = Get-QADUser $_
	$PWExpire = $User.PasswordLastSet + ((Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge)
	$First = $User.FirstName
	$Body = "<hr> Hi $First <br><br>"
	$Body += "Your password is going to expire on $PWExpire <br><br>"
	$Body += "Please change your password soon to avoid problems accessing Network Resources <br><br>"
	$Body += "If you require assistance in changing your password please contact the Help Desk.<br><br>Thank You,<br><br>"
	$Body += "Helpdesk <br>"
	$Body += "Phone: x999 <br>"
	$Body += "Email: <a herf=mailto:helpdesk@DOMAIN.LOCAL>helpdesk@DOMAIN.LOCAL</a> <br>"
	$Body += "Web: <a href=http://helpdesk/>http://HelpDesk/</a> <hr>"
	# Start crafting the message.
	$Message = New-Object System.Net.Mail.MailMessage $User.Email, $User.Email
	$Message.From = "ITTechs@DOMAIN.LOCAL"
	$Message.ReplyTo = "helpdesk@DOMAIN.LOCAL"
	$Message.Subject = "Password Expires in 10 days"
	$Message.IsBodyHTML = $true
	$Message.Priority = "Low"
	$Message.Body = ConvertTo-Html -Head $Style -Body $Body
	# Send the Message
	$SMTP.Send($Message)
}
