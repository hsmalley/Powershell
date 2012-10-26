Clear-Host
Import-Module NTFSSecurity
$TV = Get-ChildItem -Recurse "C:\Users\Hugh\Downloads\Torrents\Done"
$TV | ForEach-Object {
	$_ | Write-Host
	$_ | Enable-Inheritance
}
