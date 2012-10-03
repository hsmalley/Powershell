<#
Copies files to computers and logs it in Excel. Backout function is there if you need to backout of your copies. The logging is the weak link in this script as it runs as the user who is running the script.
#>

# Get User/Pass
$Cred = Get-Credential
# Add Quest CMDLETS
Add-PSSnapin Quest.ActiveRoles.ADManagement -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
# Get Computer's from AD
$Computers = Get-QADComputer -SearchRoot "OU=Workstations,DC=Domain,DC=local" -Credential $Cred
# Import BITS for the file transfers
Import-Module BitsTransfer

# Let's Copy the files!
Function Deploy {
	$Computers | ForEach-Object {
	# Format Computer Name
	$Computer = $_.name
	$LicensingBusinessObjects = "C:\Users\Public\AX Patch\Content Management\ApplicationXtender.Infrastructure.Licensing.LicensingBusinessObjects.dll"
	$CmConfigCtrls = "C:\Users\Public\AX Patch\Content Management\XtenderSolutions.Configuration.UI.CmConfigCtrls.dll"
	$CMXSLicenseManager = "C:\Users\Public\AX Patch\Content Management\XtenderSolutions.Utility.Licensing.CMXSLicenseManager.dll"
	$LicensingClientInterop = "C:\Users\Public\AX Patch\Content Management\XtenderSolutions.Utility.LicensingClientInterop.dll"
	$LsClient = "C:\Users\Public\AX Patch\Bin\LsClient.dll"
	# Check to see if the computer is online
	$Online = Test-Connection -Quiet -ComputerName $Computer
	IF ($Online -eq $true) {
		# If Online discover if system is 32 or 64 bit
		$WMI = Get-WmiObject -Credential $Cred -Class Win32_OperatingSystem -ComputerName $Computer
		$OSArch = $WMI.OSArchitecture
		IF ($OSArch -eq '64-bit') {
      		# 64 bit systems
			Start-BitsTransfer -Source $LicensingBusinessObjects -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files (x86)\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $CmConfigCtrls -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files (x86)\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $CMXSLicenseManager -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files (x86)\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $LicensingClientInterop -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files (x86)\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $LsClient -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files (x86)\Common Files\XtenderSolutions\bin\LsClient.dll"
			}
		ELSE {
			# 32 bit systems
			Start-BitsTransfer -Source $LicensingBusinessObjects -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $CmConfigCtrls -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $CMXSLicenseManager -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $LicensingClientInterop -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files\XtenderSolutions\Content Management"
			Start-BitsTransfer -Source $LsClient -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files\Common Files\XtenderSolutions\bin\LsClient.dll"
   			}
		}
	}
}

# For finding MD5 hash of files
function Get-MD5 ([System.IO.FileInfo] $file = $(throw 'Usage: Get-MD5 [System.IO.FileInfo]')) { 
	$stream = $null; 
	$cryptoServiceProvider = [System.Security.Cryptography.MD5CryptoServiceProvider]; 
	$hashAlgorithm = new-object $cryptoServiceProvider 
	$stream = $file.OpenRead(); 
	$hashByteArray = $hashAlgorithm.ComputeHash($stream); 
	$stream.Close(); 
	# We have to be sure that we close the file stream if any exceptions are thrown. 
	trap {if ($stream -ne $null) {$stream.Close();} 
    break;} 
	return [string]$hashByteArray; 
} 

# Let's check the versions!
# This is called in the logging function
Function GetFileInfo {
	# Check's to see if the system is online
	$Online = Test-Connection -Quiet -ComputerName $Computer
	IF ($Online -eq $true) {
		# If Online discover if system is 32 or 64 bit
		$WMI = Get-WmiObject -Credential $Cred -Class Win32_OperatingSystem -ComputerName $Computer
		$OSArch = $WMI.OSArchitecture
		IF ($OSArch -eq '64-bit') {
      		# 64 bit systems
			$File = "\\$Computer\C$\Program Files (x86)\Common Files\XtenderSolutions\bin\LsClient.dll"
			IF ((Test-Path $file) -eq $true) {
				$MD5 = Get-MD5 $File
				IF ($MD5 -eq "71 138 251 146 147 253 178 99 136 67 73 107 206 247 60 235"){
					$c.Cells.Item($intRow,2)  = "PATCHED"}
				ELSE {$c.Cells.Item($intRow,2)  = "NOT PATCHED"}
				}
			ELSE {$c.Cells.Item($intRow,2)  = "FILE MISSING OR ACCESS DENIED"}
			}
		ELSE {
			# 32 bit systems
			$File = "\\$Computer\C$\Program Files\Common Files\XtenderSolutions\bin\LsClient.dll"
			IF ((Test-Path $file) -eq $true) {
				$MD5 = Get-MD5 $File
				IF ($MD5 -eq "71 138 251 146 147 253 178 99 136 67 73 107 206 247 60 235"){
					$c.Cells.Item($intRow,2)  = "PATCHED"}
				ELSE {$c.Cells.Item($intRow,2)  = "NOT PATCHED"}
				}
			ELSE {$c.Cells.Item($intRow,2)  = "FILE MISSING OR ACCESS DENIED"}
			}	
		}
	# If Offline mark as such
	ELSE {$c.Cells.Item($intRow,2)  = "OFFLINE"}
}

# Let's log this with... EXCEL!
Function LogIT {
	# Generate Excel Document
	$a = New-Object -comobject Excel.Application
	# Let's the document be seen otherwise it would run the the background
	$a.visible = $True
	# Adds the workbook
	$b = $a.Workbooks.Add()
	# Sets up the Sheet
	$c = $b.Worksheets.Item(1)
	# Make the Headers
	$c.Cells.Item(1,1) = "Machine Name"
	$c.Cells.Item(1,2) = "LsClient Version"
	$c.Cells.Item(1,3) = "Report Time Stamp"
	# Set Font & Color
	$d = $c.UsedRange
	$d.Interior.ColorIndex = 19
	$d.Font.ColorIndex = 11
	# I like my headers to be bold. 
	$d.Font.Bold = $True
	# Starts writing in the next row
	$intRow = 2
	# Get the data
	$Computers | ForEach-Object {
		# Format the computer Name
		$Computer = $_.name
		# Input the computer name
		$c.Cells.Item($intRow,1)  = $Computer
 		# Get the file info
		GetFileInfo
		# Put's in the date
		$c.Cells.Item($intRow,3) = Get-date
		# Got to the next row
		$intRow = $intRow + 1
		}
	# Formats the data
	$d.EntireColumn.AutoFit()
}

# Oh Crap, bad things are happening! Need to backout of this NOW!
Function DangerWillRobinson {
	$Computers | ForEach-Object {
	# Format Computer Name
	$Computer = $_.name
	$LsClient = "C:\Users\Public\AX Patch\org BIN\LsClient.dll"
	# Check to see if the computer is online
	$Online = Test-Connection -Quiet -ComputerName $Computer
	IF ($Online -eq $true) {
		# If Online discover if system is 32 or 64 bit
		$WMI = Get-WmiObject -Credential $Cred -Class Win32_OperatingSystem -ComputerName $Computer
		$OSArch = $WMI.OSArchitecture
		IF ($OSArch -eq '64-bit') {
      		# 64 bit systems
			$LicensingBusinessObjects = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files (x86)\\XtenderSolutions\Content Management\\ApplicationXtender.Infrastructure.Licensing.LicensingBusinessObjects.dll'"
			$CmConfigCtrls = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files (x86)\\XtenderSolutions\Content Management\\XtenderSolutions.Configuration.UI.CmConfigCtrls.dll'"
			$CMXSLicenseManager = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files (x86)\\XtenderSolutions\Content Management\\XtenderSolutions.Utility.Licensing.CMXSLicenseManager.dll'"
			$LicensingClientInterop = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files (x86)\\XtenderSolutions\Content Management\\XtenderSolutions.Utility.LicensingClientInterop.dll'"
			Start-BitsTransfer -Source $LsClient -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files (x86)\Common Files\XtenderSolutions\bin\LsClient.dll"
			$LicensingBusinessObjects.Delete()
			$CmConfigCtrls.Delete()
			$CMXSLicenseManager.Delete()
			$LicensingClientInterop.Delete()
			}
		ELSE {
			# 32 bit systems
			$LicensingBusinessObjects = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files\\XtenderSolutions\Content Management\\ApplicationXtender.Infrastructure.Licensing.LicensingBusinessObjects.dll'"
			$CmConfigCtrls = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files\\XtenderSolutions\Content Management\\XtenderSolutions.Configuration.UI.CmConfigCtrls.dll'"
			$CMXSLicenseManager = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files\\XtenderSolutions\Content Management\\XtenderSolutions.Utility.Licensing.CMXSLicenseManager.dll'"
			$LicensingClientInterop = Get-WMIObject -computer $Computer -Credential $Cred -query "Select * From CIM_DataFile Where Name ='C:\\Program Files\\XtenderSolutions\Content Management\\XtenderSolutions.Utility.LicensingClientInterop.dll'"
			Start-BitsTransfer -Source $LsClient -Credential $cred -Description "$Computer - AX Patch" -Destination "\\$Computer\C$\Program Files\Common Files\XtenderSolutions\bin\LsClient.dll"
			$LicensingBusinessObjects.Delete()
			$CmConfigCtrls.Delete()
			$CMXSLicenseManager.Delete()
			$LicensingClientInterop.Delete()
			}
		}
	}
}
