Set-executionpolicy remotesigned
Set-ExecutionPolicy unrestricted 


# See best option below
# Option 1: Simple, but breaks with "xml" element
# based on an answer posted on stackoverflow by user "Start-Automating"
# see http://stackoverflow.com/questions/2972264/merge-multiple-xml-files-into-one-using-powershell-2-0
$files = get-childitem "C:\Users\eramos\Desktop\All Manifest Data\" 
$finalXml = "<xml>"
foreach ($file in $files) {
    [xml]$xml = Get-Content $file.fullname    
    $finalXml += $xml.xml.InnerXml
}
$finalXml += "</xml>"
$([xml]$finalXml).Save("C:\Users\eramos\Desktop\New folder\")


# Best option. Pretty prints XML 
$xmldoc = new-object xml
$rootnode = $xmldoc.createelement("stuff")
$xmldoc.appendchild($rootnode)
$finalxml = $null
$files = gci "C:\Users\eramos\Desktop\All Manifest Data\" 

foreach ($file in $files) {
    [xml]$xmlstuff = gc $file.fullname
    $innerel = $xmlstuff.selectnodes("/*/*")

    foreach ($inone in $innerel) {
        $inone = $xmldoc.importnode($inone, $true)
        $rootnode.appendchild($inone)
    }
}
# get rid of multiple spaces. might want to add regex to replace line breaks etc.
foreach ($t34 in $rootnode.selectnodes("//*/text()")) {
    $t34.innertext = [regex]::replace($t34.innertext,"\s+"," ")
}

# create and set xmlwritersettings
$xws = new-object system.xml.XmlWriterSettings
$xws.Indent = $true
$xws.indentchars = "`t"
$xtw = [system.xml.XmlWriter]::create("C:\Users\eramos\Desktop\New folder\", $xws)
$xmldoc.WriteContentTo($xtw)
$xtw.flush()
$xtw.dispose()