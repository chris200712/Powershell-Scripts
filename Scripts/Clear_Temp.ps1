Clear-Host
$Target = "$env:windir\Temp\"
$Before = (Get-ChildItem $Target -Recurse | Measure-Object Length -Sum).Sum
$Aged = (Get-Date) - (New-TimeSpan -Days 0)
$List = Get-ChildItem $Target -Recurse | ` 
Where-Object { $_.Length -ne $Null } | ` 
Where-Object { $_.LastWriteTime -lt $Aged } | ` 
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
$After = (Get-ChildItem $Target -Recurse | Measure-Object Length -Sum).Sum
'You now have an extra {0:0.00} MB of disk space' -f (($Before-$After)/1MB)