#Based off Of
#http://mikefrobbins.com/2011/10/06/find-ad-user-account-lockout-events-with-powershell/
$logName = "security"
$pcName = "DC", "DC2"
#$eventID = "4740"
$eventID = "644"
Get-EventLog -LogName $logName -ComputerName $pcName | Where {$_.eventID -eq $eventID} | Format-List -Property timegenerated, replacementstrings, message