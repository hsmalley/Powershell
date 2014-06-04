# Get-InstalledApp.ps1
# Written by Bill Stewart (bstewart@iname.com)
#
# Outputs installed applications on one or more computers that match one or
# more criteria.

param([String[]] $ComputerName,
      [String] $AppID,
      [String] $AppName,
      [String] $Publisher,
      [String] $Version,
      [Switch] $MatchAll,
      [Switch] $Help
     )

$HKLM = [UInt32] "0x80000002"
$UNINSTALL_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

# Outputs a usage message and ends the script.
function usage {
  $scriptname = $SCRIPT:MYINVOCATION.MyCommand.Name

  "NAME"
  "    $scriptname"
  ""
  "SYNOPSIS"
  "    Outputs installed applications on one or more computers that match one or"
  "    more criteria."
  ""
  "SYNTAX"
  "    $scriptname [-computername <String[]>] [-appID <String>]"
  "    [-appname <String>] [-publisher <String>] [-version <String>] [-matchall]"
  ""
  "PARAMETERS"
  "    -computername <String[]>"
  "        Outputs applications on the named computer(s). If you omit this"
  "        parameter, the local computer is assumed."
  ""
  "    -appID <String>"
  "        Select applications with the specified application ID. An application's"
  "        appID is equivalent to its registry subkey in the location"
  "        HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall. For Windows"
  "        Installer-based applications, this is the application's product code"
  "        GUID (e.g. {3248F0A8-6813-11D6-A77B-00B0D0160060})."
  ""
  "    -appname <String>"
  "        Select applications with the specified application name. The appname is"
  "        the application's name as it appears in the Add/Remove Programs list."
  ""
  "    -publisher <String>"
  "        Select applications with the specified publisher name."
  ""
  "    -version <String>"
  "        Select applications with the specified version."
  ""
  "    -matchall"
  "        Output all matching applications instead of stopping after the first"
  "        match."
  ""
  "NOTES"
  "    All installed applications are output if you omit -appID, -appname,"
  "    -publisher, and -version. Also, the -appID, -appname, -publisher, and"
  "    -version parameters all accept wildcards (e.g., -version 5.2.*)."

  exit
}

function main {
  # If -help is present, output the usage message.
  if ($Help) {
    usage
  }

  # Create a hash table containing the requested application properties.
  #CALLOUT A
  $propertyList = @{}
  if ($AppID -ne "")     { $propertyList.AppID = $AppID }
  if ($AppName -ne "")   { $propertyList.AppName = $AppName }
  if ($Publisher -ne "") { $propertyList.Publisher = $Publisher }
  if ($Version -ne "")   { $propertyList.Version = $Version }
  #END CALLOUT A

  # Use the local computer's name if no computer name(s) specified.
  if ($ComputerName -eq $NULL) {
    $ComputerName = $ENV:COMPUTERNAME
  }

  # Iterate the computer name(s).
  foreach ($machine in $ComputerName) {
    $err = $NULL

    # If WMI throws a RuntimeException exception,
    # save the error and continue to the next statement.
    #CALLOUT B
    trap [System.Management.Automation.RuntimeException] {
      set-variable err $ERROR[0] -scope 1
      continue
    }
    #END CALLOUT B

    # Connect to the StdRegProv class on the computer.
    #CALLOUT C
    $regProv = [WMIClass] "\\$machine\root\default:StdRegProv"

    # In case of an exception, write the error
    # record and continue to the next computer.
    if ($err) {
      write-error -errorrecord $err
      continue
    }
    #END CALLOUT C

    # Enumerate the Uninstall subkey.
    $subkeys = $regProv.EnumKey($HKLM, $UNINSTALL_KEY).sNames
    foreach ($subkey in $subkeys) {
      # Get the application's display name.
      $name = $regProv.GetStringValue($HKLM,
        (join-path $UNINSTALL_KEY $subkey), "DisplayName").sValue
      # Only continue of the application's display name isn't empty.
      if ($name -ne $NULL) {
        # Create an object representing the installed application.
        $output = new-object PSObject
        $output | add-member NoteProperty ComputerName -value $machine
        $output | add-member NoteProperty AppID -value $subkey
        $output | add-member NoteProperty AppName -value $name
        $output | add-member NoteProperty Publisher -value `
          $regProv.GetStringValue($HKLM,
          (join-path $UNINSTALL_KEY $subkey), "Publisher").sValue
        $output | add-member NoteProperty Version -value `
          $regProv.GetStringValue($HKLM,
          (join-path $UNINSTALL_KEY $subkey), "DisplayVersion").sValue
        # If the property list is empty, output the object;
        # otherwise, try to match all named properties.
        if ($propertyList.Keys.Count -eq 0) {
          $output
        } else {
          #CALLOUT D
          $matches = 0
          foreach ($key in $propertyList.Keys) {
            if ($output.$key -like $propertyList.$key) {
              $matches += 1
            }
          }
          # If all properties matched, output the object.
          if ($matches -eq $propertyList.Keys.Count) {
            $output
            # If -matchall is missing, break out of the foreach loop.
            if (-not $MatchAll) {
              break
            }
          }
          #END CALLOUT D
        }
      }
    }
  }
}

main
