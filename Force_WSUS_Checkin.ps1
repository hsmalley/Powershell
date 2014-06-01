# Powershell Script to force clients check into WSUS server

# Import Active Directory PS Modules CMDLETS
Import-Module ActiveDirectory

$comps = Get-ADComputer -Filter {operatingsystem -like "*server*"}

$cred = Get-Credential

Foreach ($comp in $comps) {

Invoke-Command -computername $comp.Name -credential $cred { wuauclt.exe /detectnow }
Write-Host Forced WSUS Check-In on $comp.Name

}
