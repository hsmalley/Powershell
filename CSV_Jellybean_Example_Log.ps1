# http://www.reddit.com/r/PowerShell/comments/116zee/creating_a_custom_log/c6jwr79
#I tend to abuse XML in PowerShell for any log. However, it sounds like a CSV file may be right up your alley. The important thing for using a CSV file is building an array of objects and then passing that off to Export-CSV. For example:

# Initialize the array.  I just have better luck when I do this.
$myOutArray = @()
# I'm assuming that you are using some sort of For Loop to iterate through your users/accounts/jellybeans
# $myBag should have been filled as an array elsewhere
ForEach($bean in $myBag)
{
 # Create a new PSObject to attach all of your properties to
 $objBean = New-Object PSObject
 # Add each of the properties you want to store as a NoteProperty to that PSObject
 Add-Member -MemberType NoteProperty -InputObject $objBean -Name "Color" -Value $bean.color
 Add-Member -MemberType NoteProperty -InputObject $objBean-Name "size" -Value $bean.size
 Add-Member -MemberType NoteProperty -InputObject $objBean-Name "flavor" -Value $bean.flavor
 # Attach your PSObject to your return array
 $myOutArray += $objBean
}
# Pass your array of PSObjects to Export-CSV and have it spit out a nice csv file for you.
$myOutArray | Export-CSV -NoTypeInformation -Path "C:\iso\myJellyBeans.csv"
To use that CSV file, just use a line like:
$myBagOBeans = Import-CSV -Path "C:\iso\MyJellyBeans.csv"
ForEach($bean in $myBagOBeans)
{
 write-host "Color: " + $bean.Color
 write-host "Size: " + $bean.size
 write-host "Flavor: " + $bean.flavor
}

<#XML wouldn't be all that much harder; you would just have to instantiate an xmldocument, built the xmlelements and then join them all together. If you want to see an example, I can whip something up.
EDIT: Mixed up mu comment characters, too much SQL lately.#>
