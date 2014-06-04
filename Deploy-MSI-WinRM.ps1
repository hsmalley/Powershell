$script = {
    #do preinstall stuff
    $args = "-i c:\path\to\msi\file.msi /qn /norestart"
    [diagnostics.process]::start("msiexec.exe", $args).WaitForExit()
    #do follow up stuff
}
invoke-command -computername (gc computerlist.txt) -scriptblock $script

