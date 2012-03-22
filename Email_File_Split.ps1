# Load the Winforms assembly
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")

# Create the form
$form = New-Object Windows.Forms.Form

#Set the dialog title
$form.text = "EMAIL FILE SPLITTER"

# Create the label control and set text, size and location
$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point 50,20
$label.Size = New-Object Drawing.Point 200,200
$label.text =  "Do you want to Split or Combine a File?"

# Create Button and set text and location
$button1 = New-Object Windows.Forms.Button
$button1.text = "Split"
$button1.Location = New-Object Drawing.Point 50,90

# Create Button and set text and location
$button2 = New-Object Windows.Forms.Button
$button2.text = "Combine"
$button2.Location = New-Object Drawing.Point 150,90

# Create Button and set text and location
$button3 = New-Object Windows.Forms.Button
$button3.text = "Done"
$button3.Location = New-Object Drawing.Point 100,150

# Define Button actions
$button1.add_click({

$bigFile = Open-FileDialog -Title "Select a large file to split" -Directory "$env:USERPROFILE"
$smallFile = Save-FileDialog -Title "Select Save Location and Prefix for file" -Directory "$env:USERPROFILE"
$fileSize = 10MB
split $bigFile $smallFile $fileSize

})

$button2.add_click({

$combineFile = Open-FileDialog -Title "Select files" -Directory "$env:USERPROFILE"
$saveFile = Save-FileDialog -Title "Select File Name and Location for the combined file" -Directory "$env:USERPROFILE"

Get-Content C:\Test\Email* -Encoding Byte -Read 512 | Set-Content C:\Test\EMail.zip -Encoding Byte

})

$button3.add_click({

exit

})

# Add the controls to the Form
$form.controls.add($button1)
$form.controls.add($button2)
$form.controls.add($button3)
$form.controls.add($label)

#Create the File Split
function split($inFile,  $outPrefix, [Int32] $bufSize){

  $stream = [System.IO.File]::OpenRead($inFile)
  $chunkNum = 1
  [Int64]$curOffset = 0
  $barr = New-Object byte[] $bufSize

  while( $bytesRead = $stream.Read($barr,0,$bufsize)){
    $outFile = "$outPrefix$chunkNum"
    $ostream = [System.IO.File]::OpenWrite($outFile)
    $ostream.Write($barr,0,$bytesRead);
    $ostream.close();
    echo "wrote $outFile"
    $chunkNum += 1
    $curOffset += $bytesRead
    $stream.seek($curOffset,0);
  }
}

#Create the Open File Dialog
function Open-FileDialog
{
	param([string]$Title,[string]$Directory,[string]$Filter="All Files (*.*)|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
    $objForm.Multiselect = $True
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Write-Error "Operation cancelled by user."
	}
}

#Create the Save Dialog Box
function Save-FileDialog
{
	param([string]$Title,[string]$Directory,[string]$Filter="All Files (*.*)|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.SaveFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Write-Error "Operation cancelled by user."
	}
}

# Display the dialog
$form.ShowDialog()
