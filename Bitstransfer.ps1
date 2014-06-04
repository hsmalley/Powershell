Import-Module bitstransfer
$cred = Get-Credential()
$sourcePath = \\server\example\file.txt
$destPath = C:\Local\Destination\
Start-BitsTransfer -Source $sourcePath -Destination $destPath -Credential $cred
