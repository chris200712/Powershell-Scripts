$computer = "LocalHost" 
$namespace = "root\CIMV2" 
Get-WmiObject -class Win32_RegistryAction -computername $computer -namespace $namespace