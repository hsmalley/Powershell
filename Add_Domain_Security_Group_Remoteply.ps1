$a = New-Object -comobject Excel.Application
$a.visible = $TRUE

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "Time Stamp Added"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $TRUE

$intRow = 2

$domain = "ins-lua.com"
$username = "Techs"

$strComputers = Get-content C:\Users\hsmalley\Desktop\Computers.txt
foreach ($strComputer in $strComputers){

$c.Cells.Item($intRow,1)  = $strComputer.ToUpper()

# Using .NET method to ping test the servers
$ping = new-object System.Net.NetworkInformation.Ping

$Reply = $ping.send($strComputer)

if($Reply.status -eq "success")
{
$c.Cells.Item($intRow,2)  = Get-Date

$computer = [ADSI]("WinNT://" + $strComputer + ",computer")
$computer.name

$Group = $computer.psbase.children.find("administrators")
$Group.name

# This will list what’s currently in Administrator Group so you can verify the result

function ListAdministrators

{$members= $Group.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", ‘GetProperty’, $null, $_, $null)}
$members}
ListAdministrators

# Even though we are adding the AD account but we add it to local computer and so we will need to use WinNT: provider

$Group.Add("WinNT://" + $domain + "/" + $username)

ListAdministrators

#$Group.Remove("WinNT://" + $domain + "/" + $username)

#ListAdministrators
}
else
{$c.Cells.Item($intRow,2)  = "Not Responding"}
$Reply = ""
$intRow = $intRow + 1

}
