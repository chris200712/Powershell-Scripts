Get-ADComputer -Filter * | % {$_.Name}
Get-ADComputer -Filter * | % {$_.DNSHostName}
Get-ADComputer -Filter * | Export-Csv computerList.csv -NoTypeInformation

Get-ADComputer -Filter * | % {$_.Name} | Out-File computers.txt 
$s = New-CimSession –ComputerName (Get-Content computers.txt)

Test-Connection (Get-Content computers.txt) -count 1 | select @{Name="Computername";Expression={$_.Address}},Ipv4Address
