function Get-LoggedOnUser {param([string[]]$Computer)
    $_ = Get-WmiObject Win32_ComputerSystem -Comp $Computer
    "Host Name: " + $_.Name
    "User: " + $_.UserName
}
