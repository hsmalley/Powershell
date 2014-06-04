#http://www.reddit.com/r/PowerShell/comments/116zee/creating_a_custom_log/c6jyc8p
#Ok, let the XML abuse begin:

# Assuming $myBag is some array of objects which we want to pass out
# Initalize the XML data document
$myXMLDoc = New-Object System.XML.xmlDataDocument
# XML documents must have a single root tag.
$xeRoot = $myXMLDoc.CreateElement('bag')
# the AppendChild method tends to create console output, explicitly delare as [void] to get rid of this output
[void]$myXMLDoc.AppendChild($xeRoot)
ForEach($bean in $myBag)
{
 # Create the XML Element for the bean. Note that the CreateElement Method is called from the document
 $xeBean = $myXMLDoc.CreateElement('bean')
 # Now we add our attributes.  In the XML document these will show up as part of the XML tag
 # e.g. <bean color='red' size='large' flavor='cherry' />
 $xeBean.SetAttribute('color',$bean.color)
 $xeBean.SetAttribute('size',$bean.size)
 $xeBean.SetAttribute('flavor',$bean.flavor)
 # If you want to create children, this is a good place to do it.  
 ForEach($ingredient in $bean.ingredients)
 {
  # Again, the XML Element is created using the xmldatadocument you want to use it in.
  $xeIngredient = $myXMLDoc.CreateElememt('ingredient')
  $xeIngredient.SetAttribute('name',$ingredient.name)
  $xeIngredient.SetAttribute('quantity',$ingredient.quantity)
  # Add our ingredient to our bean.  Since we're in a ForEach loop, this can happen n times 
  # giving several ingredient tags with two attributes each.
  [void]$xeBean.AppendChild($xeIngredient)
 }
 #finally, add the bean to the bag
 [void]$xeRoot.AppendChild($xeBean)
}
# And, lastly save the XML document out to disk
$myXMLDoc.Save('C:\iso\myJellyBeans.xml')


#To read it back, it is even easier:

# Create our XML Data Document object
$myBagOBeans = New-Object System.XML.xmlDataDocument
# Load our xml document from disk
$myBagOBeans.Load('C:\iso\myJellyBeans.xml')
# You can now navigate the XML tree as if it were just objects
ForEach($bean in $myBagOBeans.bag.bean)
{
 write-host "Color: " + $bean.color
 write-host "Size: " + $bean.size
 write-host "Flavor: " + $bean.flavor
 ForEach($ingredient in $bean.ingredient)
 {
   write-host "ingredient name: " + $ingredient.name
   write-host "ingredient quantity" + $ingredient.quantity
 }
}


# Merge XML Files
# http://stackoverflow.com/questions/2972264/merge-multiple-xml-files-into-one-using-powershell-2-0
$finalXml = "<root>"
foreach ($file in $files) {
    [xml]$xml = Get-Content $file    
    $finalXml += $xml.InnerXml
}
$finalXml += "</root>"
([xml]$finalXml).Save("$pwd\final.xml")
