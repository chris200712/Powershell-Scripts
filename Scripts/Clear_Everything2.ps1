$t_path_7 = "C:\Users\$env:username\AppData\Local\Microsoft\Windows\Temporary Internet Files"
$c_path_7 = "C:\Users\$env:username\AppData\Local\Microsoft\Windows\Caches"


$temporary_path =  Test-Path $t_path_7
$check_cache =    Test-Path $c_path_7


if($temporary_path -eq $True -And $check_cache -eq $True)
{
    echo "Clean history"
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1

    echo "Clean Temporary internet files"
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
    (Remove-Item $t_path_7\* -Force -Recurse) 2> $null
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2

    echo "Clean Cache"
    (Remove-Item $c_path_7\* -Force -Recurse) 2> $null


    echo "Done"
}