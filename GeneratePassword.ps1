<# Generates a random password #>
$list = [Char[]]'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_-+='
$password = -join (1..10 | Foreach-Object { Get-Random $list -count 1 })
$password

