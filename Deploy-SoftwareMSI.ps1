$Cred = Get-Credential
<# ::: Script Block to run on remote system ::: #>
$Script = {
    # Get MSI Paths
    $MSI = "-i "\\FILESERVER\DEPLOY$\APPFOLDER\APP.msi" /qn /norestart"
    # Install MSI
	[diagnostics.process]::start("msiexec.exe", $MSI).WaitForExit()
}

<# ::: Pull Computers from AD ::: Change OU if needed. ::: #>
$Computers = Get-Content C:\Users\hsmalley\Desktop\computerlist.txt
$Skipped = @()
foreach($Computer in $Computers)
{

<# ::: If the computer is not responding, record that we skipped it and continue. We can review this collection after the script completes. ::: #>
	if (-not (Test-Connection -Quiet $Computer -Count 1))
	{
		$Skipped += $Computer
	}
	Write-Host $Computer
	Invoke-Command -ComputerName $Computer -ScriptBlock $Script -Credential $Cred
}
<# ::: Show Systems Skipped ::: #>
Write-Host $Skipped
$Skipped | Out-File -FilePath C:\Users\HS1\Desktop\NotOnLine.txt
