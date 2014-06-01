# Setup Email fields
$SMTP = New-Object Net.Mail.SMTPClient
$SMTP.Host = "MAILSERVER.DOMAIN.LOCAL"
$SMTP.TargetName = "MAILSERVER.DOMAIN.LOCAL"
$EmailFrom = "ITTechs@DOMAIN.LOCAL"
$Subject = "Litigation Hold Report - Running From $Env:COMPUTERNAME"

# CSS Style for Email.
$Style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$Style = $Style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$Style = $Style + "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
$Style = $Style + "TD{border: 1px solid black; padding: 5px;}"
$Style = $Style + "</style>"

# Exchange
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://bocaexch10/PowerShell/"
Import-PSSession $Session -AllowClobber

# The Work
$BodyLog = @()
Get-Mailbox | ForEach-Object {
	IF ($_.LitigationHoldEnabled -eq $true) {
		#$Body += $_ | Select-Object "Name", "RetentionComment"
		$Name = $_.Name
		$RetentionComment = $_.RetentionComment
		$Log =  New-Object PSObject
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Name" -Value $Name -Force
		Add-Member -MemberType NoteProperty -InputObject $Log -Name "Retention Comment" -Value $RetentionComment -Force
		$BodyLog += $Log
	}
}

# The Mail
$Body = $BodyLog | ConvertTo-Html

# Send Message to Helpdesk
$EmailTO = "helpdesk@DOMAIN.LOCAL"
$Message = New-Object System.Net.Mail.MailMessage $EmailFrom, $EmailTO
$Message.Subject = $Subject
$Message.IsBodyHTML = $true
$Message.Body = ConvertTo-Html -Head $Style -Body $Body -Title "Litigation Hold Report"
$SMTP.Send($Message)
