$a = New-Object -comobject Excel.Application
$a.visible = $True 

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Server Name"
$c.Cells.Item(1,2) = "Operating System"
$c.Cells.Item(1,3) = "OS Version"
$c.Cells.Item(1,4) = "IP Address"
$c.Cells.Item(1,5) = "System Type"
$c.Cells.Item(1,6) = "Install Date"
$c.Cells.Item(1,7) = "Manufacturer"
$c.Cells.Item(1,8) = "Model"
$c.Cells.Item(1,9) = "Service Tag"
$c.Cells.Item(1,10) = "Serial Number"
$c.Cells.Item(1,11) = "SKU Number"
$c.Cells.Item(1,12) = "Number of Processors"
$c.Cells.Item(1,13) = "Total Phsyical Memory (GB)"
$c.Cells.Item(1,14) = "Last Reboot Time"
$c.Cells.Item(1,15) = "Report Time Stamp"
$c.Cells.Item(1,16) = "Service Packs"
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$colComputers = Get-Content C:\MachineList.txt


foreach ($strComputer in $colComputers)

{
$OS = Get-WmiObject Win32_OperatingSystem -computername $strComputer
$IP = Get-WmiObject Win32_NetworkAdapterConfiguration -computername $strComputer | ? { $_.IPAddress -ne $null }
$Computer = Get-WmiObject Win32_ComputerSystem -computername $strComputer
$Bios = Get-WmiObject Win32_Bios -computername $strComputer


$c.Cells.Item($intRow,1) = $strComputer.Toupper()
$c.Cells.Item($intRow,2) = $OS.Caption
$c.Cells.Item($intRow,3) = $OS.Version
$c.Cells.Item($intRow,4) = $IP.IPAddress
$c.Cells.Item($intRow,5) = $Computer.SystemType
$c.Cells.Item($intRow,6) = [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.InstallDate)
$c.Cells.Item($intRow,7) = $Computer.Manufacturer
$c.Cells.Item($intRow,8) = $Computer.Model
$c.Cells.Item($intRow,9) = $Bios.SerialNumber
$c.Cells.Item($intRow,10) = $OS.SerialNumber
$c.Cells.Item($intRow,11) = $Computer.SystemSKUNumber
$c.Cells.Item($intRow,12) = $Computer.NumberOfProcessors
$c.Cells.Item($intRow,13) = "{0:N0}" -f ($computer.TotalPhysicalMemory/1GB)
$c.Cells.Item($intRow,14) = [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.LastBootUpTime)
$c.Cells.Item($intRow,15) = Get-Date -format g
$c.Cells.Item($intRow,16) = $OS.CSDVersion
$intRow = $intRow + 1
}
$d.EntireColumn.AutoFit()
cls