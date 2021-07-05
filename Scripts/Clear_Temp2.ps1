Clear-Host
$Target = "$env:windir\Temp\"
$List = Get-ChildItem $Target -Recurse | ` 
Where-Object { $_.Length -ne $Null } | ` 
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue 
