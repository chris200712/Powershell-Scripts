########################################################################
# Date: 5/19/2010
# Author: Rich Prescott
# Blog: blog.richprescott.com
# Twitter: @Rich_Prescott
########################################################################

$ClassList = gwmi -list win32_*
foreach ($class in $ClassList)
{
$classname = $class.name
write-host "`$ComputerName = `".`"`r`n"
write-host "`$colItems = Get-WMIObject -class `"$classname`" -namespace `"root\CIMV2`" -computername `$Computername `r`n"
write-host "ForEach(`$objItem in `$colItems){"
foreach ($property in $class.properties){"`tWrite-Host `"$($property.Name): `" `$objitem.$($property.name)"}
write-host "`tWrite-Host"
write-host "}"
write-host "-----------------------------------------------------------------------"
} #End foreach class