#Author: Steve Wright
#Date: 20/01/2012 
#Script: Update-SQLServer-Statistics

function Update-SQLServer-DBStatistics
{
<#
.SYNOPSIS
Updates the SQL Server Database Statistics

.DESCRIPTION
Updates the SQL Server Database Statistics of the request database on a SQL Server if the Statistics are over a day old. 
Using windows authentication for connecting to remote servers.
Also to be able to use the funcation the SQL Server SMO and powershell SQL Provider

.PARAMETER server 
The name of the SQL Server to connect to gather the file usages of the database. Default Value:(local)

.PARAMETER databaseName
The name of the database which the statistics need to be updated.

.PARAMETER scanType
Specify the ways in which statistical information is collected from tables or views during the creation or update of a statistic counter.

Default:(Default) Specifies that a percentage of the table or indexed view is used when collecting statistics. The actual percentage is calculated by the SQL Server engine automatically.  
Resample:Specifies that the percentage ratio of the table or indexed view used when collecting statistics is inherited from existing the statistics.  
FullScan:Specifies that all rows in the table or view are read when gathering statistics. This option must be used if a view is specified and it references more than one table.  
Percent:Specifies that a percentage of the table or indexed view is used when collecting statistics. This options cannot be used if a view is specified and it references more than one table. When specified, use the sampleValue argument to indicate number of rows.  
Rows:Specifies that a number of rows in the table or indexed view are used when collecting statistics. This option cannot be used if a view is specified and it references more than one table. When specified, use the sampleValue argument to indicate number of rows.  

.PARAMETER scanRows 
The Specifies the percentage of the table or indexed view, or the number of rows to sample when collecting statistics for larger tables or views. 
If 0 then not used. Default Value 0
 
.EXAMPLE
Connects to the local server database myDB using windows authentication
	
Update-SQLServer-DBStatistics -databaseName myDB

.EXAMPLE
Connects to the remote server database using windows authentication
	
Update-SQLServer-DBStatistics -server SQLServer01 -databaseName myDB

.EXAMPLE
Connects to the local server database using windows authentication with scanType of Resample
	
Update-SQLServer-DBStatistics -databaseName myDB -scanType [Microsoft.SqlServer.Management.Smo.StatisticsScanType]::Resample 

.EXAMPLE
Connects to the remote server database using windows authentication with scanType of FullScan
	
Update-SQLServer-DBStatistics -server SQLServer01 -databaseName myDB -scanType [Microsoft.SqlServer.Management.Smo.StatisticsScanType]::FullScan 

.EXAMPLE
Connects to the local server database using windows authentication with scanType of Percent and set the scan rows 4000
	
Update-SQLServer-DBStatistics -databaseName myDB -scanType [Microsoft.SqlServer.Management.Smo.StatisticsScanType]::Percent -scanRows 4000

.EXAMPLE
Connects to the remote server database using windows authentication with scanType of Rows and set the scan rows 4000
	
Update-SQLServer-DBStatistics -server SQLServer01 -databaseName myDB -scanType [Microsoft.SqlServer.Management.Smo.StatisticsScanType]::Rows -scanRows 4000
	
.INPUTS
None. You cannot pipe objects to Update-SQLServer-DBStatistics
 
.Outputs
Array of PSObject with the following properties: 
Database 
Table
Stat
IsAutoCreated 
IsFromIndexCreation 
LastUpdated
NumberOfDaysOld
WasUpdated

.COMPONENT
Microsoft® Windows PowerShell Extensions for SQL Server® 2008 R2.

.COMPONENT
Microsoft® SQL Server® 2008 R2 Shared Management Objects.

.LINK
Components Download http://www.microsoft.com/download/en/details.aspx?displaylang=en&id=16978 
#>
	[CmdletBinding()]
	param (
		[string]
		#The server that the Job should be run on
		$server = "(local)",
		[string]
		#The name of the database to Backup
		$databaseName, 
		[Microsoft.SqlServer.Management.Smo.StatisticsScanType]
		$scanType = [Microsoft.SqlServer.Management.Smo.StatisticsScanType]::Default,
		[int]
		$scanRows = 0 
	)
	
	$properties = @{Database = [string] "";
	                Table = [string] "";
					Stat = [string] "";
					IsAutoCreated = [bool] $false;
  					IsFromIndexCreation = [bool] $false;
					LastUpdated = [DateTime] [DateTime]::MinValue;
					NumberOfDaysOld = [int] 0;
					WasUpdated = [bool] $false;
					}

	$srv = New-Object -TypeName Microsoft.SqlServer.Management.SMO.Server -ArgumentList $server
	$db = $srv.Databases[$databaseName]
	$results = @()

	foreach($dbtable in $db.Tables) 
	{
		foreach($dbStatistic in $dbtable.Statistics) 
		{
			$dayDiff = [DateTime]::Now - $dbStatistic.LastUpdated 
			$methodTaken = New-Object PSObject -Property $properties
			$methodTaken.Database = $db.Name
			$methodTaken.Table = $dbtable
			$methodTaken.Stat = $dbStatistic.Name
			$methodTaken.IsAutoCreated = $dbStatistic.IsAutoCreated 
  			$methodTaken.IsFromIndexCreation = $dbStatistic.IsFromIndexCreation
			$methodTaken.LastUpdated = $dbStatistic.LastUpdated
			$methodTaken.NumberOfDaysOld = $dayDiff.Days
		
			if($dayDiff.Days -ge 1)
			{
				if($scanRows -eq 0)
				{
					$dbStatistic.Update($scanType)
				}
				else
				{
					$dbStatistic.Update($scanType, $scanRows)
				}
				
				$methodTaken.WasUpdated = $true
			}
					 
			$results += $methodTaken
		}	
	}
	return $results
}
