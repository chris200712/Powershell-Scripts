# Insert timestamps into PowerShell outputs

“$(Get-Date -format g) Start logging”
“$(Get-Date -format F) Start logging”
“$(Get-Date -format o) Start logging”

# Report all of the USB devices installed

gwmi Win32_USBControllerDevice -Computername BLACKWIDOW | fl Antecedent,Dependent

# Enumerate USB devices

Get-USB -Computername BLACKWIDOW ! fl

# Processor Info

Get-Processor

# PowerShell script to enumerate the event logs.

Get-Eventlog -list

# PowerShell script to search Error messages in the System eventlog.

Clear-Host
Get-Eventlog system -newest 2000 | where {$_.entryType -match "Error"}

# Cmdlet to find latest 2000 errors in the System eventlog

Clear-Host
$SysEvent = Get-Eventlog -logname system -newest 2000
$SysError = $SysEvent |where {$_.entryType -match "Error"}
$SysError | sort eventid | `
Format-Table EventID, Source, TimeWritten, Message -auto

# PowerShell script to list the event logs.
Get-WmiObject -class Win32_NTLogEvent

# WMI Win32_NTLogEvent PowerShell script
Clear-Host
$Logs = Get-WmiObject -class Win32_NTLogEvent `
-filter "(logfile='Application') AND (type='error')" 
$Logs | Format-Table EventCode, EventType, Message -auto

# Win32_NTLogEvent with Select *

Clear-Host
$Logs = Get-WmiObject -query `
"SELECT * FROM Win32_NTLogEvent WHERE (logfile='Application') AND (type='error')"
$Logs | Format-Table EventCode, EventType, Message -auto

# WMI Win32_NTLogEvent Properties

Clear-Host
Get-WmiObject Win32_NTLogEvent | Get-Member -memberType Properties

# WMI Win32_NTLogEvent filter example

Clear-Host
$Logs = Get-WmiObject -query `
"SELECT * FROM Win32_NTLogEvent WHERE (logfile='Application') AND (SourceName='Application Hang')"
$Logs | Format-Table EventCode, SourceName, Message -auto

# PowerShell Get-WinEvent script to list the event logs.

Get-WinEvent -Listlog *

# PowerShell Get-WinEvent script to list classic event logs.
Clear-Host
Get-WinEvent -listlog * | Where {$_.IsClassicLog -eq 'True'}

# PowerShell Get-WinEvent script find errors in Application log
Clear-Host
Get-WinEvent -logName Application -maxEvents 500 | `
Where-Object {$_.DisplayName -eq 'Error'} | `
Format-Table DisplayName, id, ProviderName -auto

# PowerShell example which groups event then sorts in descending order.
Clear-Host
Get-WinEvent -logName System -maxEvents 2000 | `
Group-Object ProviderName | Sort-Object Count -descending | `
Format-Table Count, Name -auto

# Investigate PowerShell Get-WinEvent -parameters
Clear-Host
Get-Help Get-WinEvent -full

# Investigate PowerShell Get-WinEvent Properties
Clear-Host
Get-WinEvent system -maxEvents 1 | Get-Member -MemberType property

# PowerShell Get-WinEvent script find errors in Application log
Get-WinEvent -logName Application -maxEvents 500 | `
Where-Object {$_.DisplayName -eq 'Error'} | `
Select DisplayName, id, ProviderName

# PowerShell script to list the dll files under C:\Windows\System32
$i =0
$Files = Gci "C:\Windows\" -recurse | ? {$_.extension -eq ".dll"}
Foreach ($Dll in $Files) {
"{0,-28} {1,-20} {2,12}" -f `
$Dll.name, $DLL.CreationTime, $Dll.Length
$i++
}
Write-Host The total number of dlls is: $i



Get-WinEvent -ListLog * -EA silentlycontinue |
where-object { $_.recordcount -AND $_.lastwritetime -gt [datetime]::today} |
foreach-object { get-winevent -LogName $_.logname -MaxEvents 1 } |
Format-Table TimeCreated, ID, ProviderName, Message -AutoSize –Wrap

#Eventlogs between two dates

$May1 = get-date 01/05/16
$July1 = get-date 01/07/16

$birthday = (get-date 26/07/88).dayofweek

get-eventlog -log "Application" -entrytype Error -after $may1 -before $july1

#Get a list of installed software on a computer and output it to a csv file

Get-WmiObject win32_Product  | Select Name,Version,PackageName,Installdate,Vendor | Sort InstallDate -Descending| Export-Csv c:\report.csv

