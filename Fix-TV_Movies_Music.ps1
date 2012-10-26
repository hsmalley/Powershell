Clear-Host
Import-Module NTFSSecurity
$TV = Get-ChildItem -Recurse "C:\Users\Public\Videos\TV"
$Movies = Get-ChildItem -Recurse "C:\Users\Public\Videos\Movies"
$Music = Get-ChildItem -Recurse  "C:\Users\Public\Music"
$TV,$Movies,$Music | ForEach-Object {
	$_ | Write-Host
	$_ | Enable-Inheritance
}
