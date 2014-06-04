$logFileName = "Application" # Add Name of the Logfile (System, Application, etc)
$path = "C:\Users\hsmalley\Desktop\" # Add Path, needs to end with a backsplash

# do not edit
$exportFileName = "APPLICATION_EVENT.EVT"
$logFile = Get-WmiObject -computername ws01335 Win32_NTEventlogFile | Where-Object {$_.logfilename -eq $logFileName}
$logFile.backupeventlog($path + $exportFileName)
