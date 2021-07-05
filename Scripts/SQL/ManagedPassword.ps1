#Author: Steve Wright
#Date: 20/01/2012 
#Script: The functions to set and get the password

function Get-ManagedPassword
{
<#
.SYNOPSIS
Get the password for the file and return a SecureString

.DESCRIPTION
Reads the contents of the file and converts to a secure string

.PARAMETER fileLocation 
The location of the file contencts which needs to be converted to a SecureString

.EXAMPLE
Get the secure string from a SQL.bin file in password folder 
Get-ManagedPassword -fileLocation c:\password\SQL.bin

.INPUTS
None. You cannot pipe objects to Get-ManagedPassword
 
.Outputs
a SecureString of the contents of the file. 
#>
	param
	(
		[string] $fileLocation 
	)
	if(Test-Path $fileLocation)
	{
		return ConvertTo-SecureString (gc $fileLocation)
    }
	else 
	{
		return $null
	}
}

function Set-ManagedPassword
{
<#
.SYNOPSIS
Store the password to the file location.  If the file location already exists then compare the old
password before setting the password

.DESCRIPTION
Store the password to the file location as a SecureString.  If the file location already exists then compare the old
password before setting the new password.

.PARAMETER fileLocation 
The full path to the file to hold the password

.PARAMETER oldPassword
The old password which currently stored witin the file, or blank if this is a new file.

.PARAMETER passwordString
The password to store within the file

.EXAMPLE
Createing a new password file
	
Set-ManagedPassword -fileLocation c:\passwords\sql.pass -oldPassword "" -passwordString Pa55w0rd

.EXAMPLE
Changing the password stored in an old file

Set-ManagedPassword -fileLocation c:\passwords\sql.pass -oldPassword Pa55w0rd -passwordString N3wPa55W0rd
	
.INPUTS
None. You cannot pipe objects to Set-ManagedPassword
 
.Outputs
None.

#>
	param
	(
		[string] $fileLocation, 
		[string] $oldPassword,
		[string] $passwordString
	)
	
	$passwordStringss = ConvertTo-SecureString -string ($passwordString) -AsPlainText -Force
	$oldPasswordss = ConvertTo-SecureString -string ($oldPassword) -AsPlainText -Force
		
	if(Test-Path $fileLocation)
	{
		$filePassword = Get-ManagedPassword -fileLocation $fileLocation
		if($oldPasswordss -ne $filePassword)
		{
			throw "Password not set"
		}
		
	}
	else
	{
		if($oldPassword -ne "")
		{
			throw "Password not set"
		}
	}
	
	convertfrom-securestring $passwordStringss | out-file  $fileLocation -encoding ascii	
}