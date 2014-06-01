$Style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$Style = $Style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$Style = $Style + "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
$Style = $Style + "TD{border: 1px solid black; padding: 5px;}"
$Style = $Style + "</style>"

# Export AD Groups
$Data += $nul
$Groups = Get-QADGroup -SizeLimit "100000"
$Groups | ForEach-Object {
	<# START CSV LOG #>
	$ReportData =  New-Object PSObject
	Add-Member -MemberType NoteProperty -InputObject $ReportData -Name "Group" -Value $_.Name -Force
	Add-Member -MemberType NoteProperty -InputObject $ReportData -Name "Notes" -Value $_.Notes -Force
	$Report = @()
	$Report += $ReportData
	$ReportPath = "C:\Users\hsmalley\Desktop\ADGroups.csv"
	$Report | Export-Csv -Append -Force -NoTypeInformation $ReportPath
	<# END CSV LOG #>
}

# Process Groups
$Note = $null
$Data = Import-Csv $ReportPath
$Data | ForEach-Object {
	$Group = $_.Group
	$Notes = $_.Notes
	$FindNoteUSERID = $null
	$FindNoteUSERID = Select-String -Pattern "USERID" -InputObject $Notes
	IF ($FindNoteUSERID -ne $null) {
		$Note += "$Group"
		$Note += "<br>"
		$Note += "$FindNoteUSERID
		$Note += "<hr>"
	}
}

#Export to HTML
ConvertTo-Html -Head $Style -Body $Note | Out-File "C:\Users\hsmalley\Desktop\ADGroups.html"
