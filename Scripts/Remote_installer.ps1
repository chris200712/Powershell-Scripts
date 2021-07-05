#Variables
$computername = Get-Content C:\PSdeploy\list.txt
$sourcefile = "\\Servershare\7u25\jre-7u25-windows-x64.exe"
#This section will install the software 
foreach ($computer in $computername) 
{
	$destinationFolder = "\\$computer\C$\Temp"
	#This section will copy the $sourcefile to the $destinationfolder. If the Folder does not exist it will create it.
	if (!(Test-Path -path $destinationFolder))
	{
		New-Item $destinationFolder -Type Directory
	}
	Copy-Item -Path $sourcefile -Destination $destinationFolder
	Invoke-Command -ComputerName $computer -ScriptBlock {Start-Process 'c:\temp\jre-7u25-windows-x64.exe'}