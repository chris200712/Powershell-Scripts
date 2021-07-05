#Possible script parameters
$dbServer = "(local)"
$dbDatabase = "myData"
$sqlServerBackFolder = "C:\SQLServer\Backups"
$passwordFile = "C:\Projects\PSScripts\backup.pas"

#set the location to the path of the script file 
#and load the extra functions required to run the script
Set-Location ([System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition))
. .\BackUp-Database.ps1
. .\Defrag-SQLServer-Indexes.ps1
. .\get-DBfileStats.ps1
. .\Update-SQLServer-Statistics.ps1
. .\ManagedPassword.ps1


#Script variables
$backupPassword = Get-ManagedPassword -fileLocation $passwordFile
$dateString = Get-Date -Format yyyyMMdd

#output log files
$DefragSQLServerIndexesFileName = "DefragSQLServerIndexes.csv"
$StatisticsFileName = "Statistics.csv"
$DeletedFileFilename = "fileDeleted.csv"
$DataFileGrowthName = "DataFileGrowth_$dateString.csv"

#Check which day of the week it is for now
if([DateTime]::Now.DayOfWeek -eq [DayOfWeek]::Sunday)
{
	#Remove any old backups which are 7 days old
	Get-ChildItem -Path $sqlServerBackFolder -Recurse | % { 	
		$dayDiff = [DateTime]::Now - $_.CreationTime 
		if($dayDiff.Days -gt 7)
		{
			#record the any file that has been deleted
			[string[]]$filesDeleted += $_.Name
			Remove-Item $_.FullName
		}
	}
	#If files have been delete export the list to CSV file
	if($filesDeleted  -ne $null)
	{
		$filesDeleted | Export-Csv $DeletedFileFilename -notype
	}
	else
	{
	#If backup files haven't been deleted and deleted file exist remove the file
		if(Test-Path $DeletedFileFilename)
		{
			Remove-Item $DeletedFileFilename
		}
	}
	#Do a full back up of the database with the a password
	BackUp-Database -server $dbServer -databaseName $dbDatabase -backFolderPath $sqlServerBackFolder -isIncremental $false -backupPassword $backupPassword


	#remove the previous result files
	if(Test-Path $DefragSQLServerIndexesFileName)
	{
		Remove-Item $DefragSQLServerIndexesFileName
	}
	if(Test-Path $StatisticsFileName)
	{
		Remove-Item $StatisticsFileName
	}
	
	#Defrag the database indexes as the will update the stats at the sametime
	Defrag-SQLServer-Indexes -databaseName $dbDatabase | Select-Object Database, Table, Index, AverageFragmentation, ActionTaken | Export-Csv $DefragSQLServerIndexesFileName -notype
	
	#Updated any Stats that are not linked to index or where indexes not updated
	Update-SQLServer-Statistics -databaseName $dbDatabase | Select-Object Database, Table, Stat, LastUpdated, NumberOfDaysOld, WasUpdated | Export-Csv $StatisticsFileName -notype
	get-dbFileStats -server $dbServer | Select-Object Database, Filegroup, LogicalFileName, PsychicalFileName, UsedSpace, Size, VolumeFreeSpace | Export-Csv $DataFileGrowthName -notype



}
else
{	
	#Do a differential backup of the database with a password
	BackUp-Database -server $dbServer -databaseName $dbDatabase -backFolderPath $sqlServerBackFolder -isIncremental $true -backupPassword $backupPassword
}
	                  