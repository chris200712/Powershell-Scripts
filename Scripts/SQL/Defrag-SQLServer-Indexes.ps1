#Author: Steve Wright
#Date: 20/01/2012 
#Script: Update-SQLServer-Statistics

function Defrag-SQLServer-Indexes
{
<#
.SYNOPSIS
Defrag the Indexes of the request SQL Server database

.DESCRIPTION
For the selected database on the SQL Server will loop through all the user tables indexes to see if they 
need to be Reorganize or Rebuild.
Using windows authentication for connecting to remote servers.
Also to be able to use the funcation the SQL Server SMO and powershell SQL Provider

.NOTES
The function will either Reorganize Index or Rebuild Index
If the index AverageFragmentation ranges in between 5% to 30% then it is better to perform Reorganize Index.
If the index AverageFragmentation is greater than 30% then the best strategy will be to use Rebuild Index.
Recommandations where found on many articles on the internet

.PARAMETER server 
The name of the SQL Server to connect to gather the file usages of the database. Default Value:(local)

.PARAMETER databaseName
The name of the database which needs to have the indexes Defrag

.PARAMETER fragmentationOption
The specify the levels of detail of collected fragmentation information
Fast:Calculates statistics based on parent level pages only. This option is available starting with SQL Server 2000.  
Sampled:Calculates statistics based on samples of data. This option is available starting with SQL Server 2005.  
Detailed:(Default) Calculates statistics based on 100% of the data. This option is available starting with SQL Server 2005.  

.EXAMPLE
Connects to the local server database myDB using windows authentication
	
Defrag-SQLServer-Indexes -databaseName myDB

.EXAMPLE
Connects to the remote server database using windows authentication
	
Defrag-SQLServer-Indexes -server SQLServer01 -databaseName myDB

.EXAMPLE
Connects to the local server database using windows authentication with fragmentationOption of Fast
	
Defrag-SQLServer-Indexes -databaseName myDB -fragmentationOption [Microsoft.SqlServer.Management.Smo.FragmentationOption]::Fast 

.EXAMPLE
Connects to the remote server database using windows authentication with fragmentationOption of Sampled
	
Defrag-SQLServer-Indexes -server SQLServer01 -databaseName myDB -fragmentationOption [Microsoft.SqlServer.Management.Smo.FragmentationOption]::Sampled 
	
.INPUTS
None. You cannot pipe objects to Defrag-SQLServer-Indexes
 
.Outputs
Array of PSObject with the following properties: 
Database
Table
Index
AverageFragmentation
ActionTaken

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
		# The server that the Job should be run on
		$server = "(local)",
		[string]
		# The name of the database to Backup
		$databaseName,
		#The level of the Fragmentation Scan
		[Microsoft.SqlServer.Management.Smo.FragmentationOption]
		$fragmentationOption = [Microsoft.SqlServer.Management.Smo.FragmentationOption]::Detailed
	)
	
	$properties = @{Database = [string] "";
	                Table = [string] "";
					Index = [string] "";
					AverageFragmentation = [float] 0.0;
					ActionTaken = [string] "";
					}

	$srv = New-Object -TypeName Microsoft.SqlServer.Management.SMO.Server -ArgumentList $server
	$db = $srv.Databases[$databaseName]
	$results = @()

	foreach($dbtable in $db.Tables) 
	{
		foreach($dbIndex in $dbtable.Indexes) 
		{
			$indexResults = $dbIndex.EnumFragmentation($fragmentationOption)
			$methodTaken = New-Object PSObject -Property $properties
			$methodTaken.Database = $db.Name
			$methodTaken.Table = $dbtable
			$methodTaken.Index = $dbIndex.Name
			$methodTaken.AverageFragmentation = $indexResults.Rows[0]["AverageFragmentation"]
			  
			 if($methodTaken.AverageFragmentation -ge 30 )
			 {
			 	$methodTaken.ActionTaken = "Rebuild"
			 	$dbIndex.Rebuild()
			 }
			 elseif($methodTaken.AverageFragmentation -ge 5) 
			 {
			 	$methodTaken.ActionTaken = "Reorganize"
			 	$dbIndex.Reorganize()
			 }
			 else
			 {
			 	$methodTaken.ActionTaken = "None"
			 }
			 
			 $results += $methodTaken
		}	
	}
	return $results
}

