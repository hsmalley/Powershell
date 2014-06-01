<#
Based off Of:
http://mikefrobbins.com/2011/10/06/find-ad-user-account-lockout-events-with-powershell/

EventID 4740 is for 2008+
EventID 644 is for 2003

Change the Domain Controlers if you don't get a result.
You need to run this with Domain Admin rights.
#>

$logName = "security"
$pcName = "BDC", "DC00002"
#$eventID = "4740"
$eventID = "644"
Get-EventLog -LogName $logName -ComputerName $pcName | Where {$_.eventID -eq $eventID} | Format-List -Property timegenerated, replacementstrings, message
