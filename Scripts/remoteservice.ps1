# this script permits to get service info for a set of remote server 

# syntax
#get-content .\MyServerList.txt | get-serviceonRemoteServer.ps1

#command line to export-csv
#get-content .\MyServerList.txt | get-serviceonRemoteServer.ps1 | export-csv .\MyServiceResult.csv -noTypeInformation

#Comand line to generate an HTML file
#get-content .\MyServerList.txt | get-serviceonRemoteServer.ps1 | convertToHtml | out-file .\MyServiceResult.Html



process{
$servername = $_
$serviceArray = Get-WmiObject -class Win32_Service -computerName $servername
$temp = @()  

$serviceArray | foreach {
	$serviceinfo = "" | select servername, name, startmode, state, status

	$serviceinfo.servername = $servername
	$serviceinfo.name = $_.name
	$serviceinfo.startmode = $_.startmode
	$serviceinfo.state = $_.state
	$serviceinfo.status = $_.status
	
	$temp +=$serviceinfo
	}
$temp 
}