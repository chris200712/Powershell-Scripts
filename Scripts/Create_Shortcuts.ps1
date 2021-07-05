#--------------------------------------------------------------------------------- 
#The script is not supported under any Microsoft standard support 
#program or service. The script is provided AS IS without warranty  
#of any kind. Microsoft further disclaims all implied warranties including,  
#without limitation, any implied warranties of merchantability or of fitness for 
#a particular purpose. The entire risk arising out of the use or performance of  
#the sample scripts and documentation remains with you. In no event shall 
#Microsoft, its authors, or anyone else involved in the creation, production, or 
#delivery of the scripts be liable for any damages whatsoever (including, 
#without limitation, damages for loss of business profits, business interruption, 
#loss of business information, or other pecuniary loss) arising out of the use 
#of or inability to use the sample scripts or documentation, even if Microsoft 
#has been advised of the possibility of such damages 
#--------------------------------------------------------------------------------- 

#requires -Version 3.0

Function New-SaveLocation
{
<#
 	.SYNOPSIS
        New-Storage is an advanced function which can be used to create shortcuts for Google Drive and Dropbox in Office 2013.
    .DESCRIPTION
    .PARAMETER  <Gdrive>
        Creates a shortcut to Google Drive for Office 2013
    .PARAMETER <Dropbox>
        Creates a shortcut to Dropbox for Office 2013
    
    .EXAMPLE
        C:\PS> New-SaveLocation
		
		This command creates a shortcut to both Google Drive and Dropbox for Office 2013
    .EXAMPLE
        C:\PS> New-SaveLocation -Gdrive
		
		This command creates a shortcut to Google Drive for Office 2013
    .EXAMPLE
        C:\PS> New-SaveLocation -Dropbox

        This command creates a shortcut to Dropbox for Office 2013

#>

    [CmdletBinding(SupportsShouldProcess,DefaultParameterSetName="__AllParameterSets")]
    Param
    (
        [Parameter(Position=0,Mandatory,ParameterSetName="GoogleDrive")]
        [Alias('googledrive')][Switch]$GdriveShortcut,
        [Parameter(Position=0,Mandatory,ParameterSetName="Dropbox")]
        [Alias('dropbox')][Switch]$DropboxShortcut
    )

Begin
    {
        $Shell = New-Object -ComObject Shell.Application
	    $Desktop = $Shell.NameSpace(0X0)
        $WshShell = New-Object -comObject WScript.Shell
    }

    Process
    {
        If($GdriveShortcut)
        {
            CreateGdrive
        }
        ElseIf($DropboxShortcut)
        {
            CreateDropbox
        }
        Else
        {
            CreateGdrive
            CreateDropbox
        }
    }
}

Function CreateGdrive
{
    #Get UserName and check default location of folder
    [string]$username = [System.Environment]::UserName
    [string]$gdrivedir = "C:\Users\" + $($username) + "\Google Drive"

    #Ask for input if folder can't be found in default location
    If(!(Test-Path -Path "$($gdrivedir)"))
    {
        [string]$gdrivedir = Read-Host 'Could not find Google Drive. Enter path to folder manually: '
    }
    
    #Add registry values
    New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66' -Name DisplayName -PropertyType String -Value 'Google Drive'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66' -Name Description -PropertyType String -Value 'Google Drive is a free service that lets you bring all your photos, docs, and videos anywhere.'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66' -Name Url48x48 -PropertyType String -Value http://k37.kn3.net/7DDDA8544.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66' -Name LearnMoreURL -PropertyType String -Value https://drive.google.com/
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66' -Name ManageURL -PropertyType String -Value https://drive.google.com/
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66' -Name LocalFolderRoot -PropertyType String -Value $gdrivedir
    
    New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails' -Name Url48x48 -PropertyType String -Value http://k37.kn3.net/7DDDA8544.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails' -Name Url40x40 -PropertyType String -Value http://k35.kn3.net/AA5EBCDA4.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails' -Name Url32x32 -PropertyType String -Value http://k31.kn3.net/022B096E1.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails' -Name Url24x24 -PropertyType String -Value http://k35.kn3.net/397FB33CC.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails' -Name Url20x20 -PropertyType String -Value http://k35.kn3.net/397FB33CC.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\ecc30fd0-4546-11e2-bcfd-0800200c9a66\Thumbnails' -Name Url16x16 -PropertyType String -Value http://k43.kn3.net/E5162FC3D.png
}

Function CreateDropbox
{
    #Get UserName and check default location of folder
    [string]$username = [System.Environment]::UserName
    [string]$dropboxdir = "C:\Users\" + $($username) + "\Dropbox"

    #Ask for input if folder can't be found in default location
    If(!(Test-Path -Path "$($dropboxdir)"))
    {
        [string]$dropboxdir = Read-Host 'Could not find Dropbox. Enter path to folder manually: '
    }
    
    #Add registry values
    New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad' -Name DisplayName -PropertyType String -Value 'Dropbox'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad' -Name Description -PropertyType String -Value 'Dropbox is a free service that lets you bring all your photos, docs, and videos anywhere.'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad' -Name Url48x48 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_48x48.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad' -Name LearnMoreURL -PropertyType String -Value https://www.dropbox.com/
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad' -Name ManageURL -PropertyType String -Value https://www.dropbox.com/account
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad' -Name LocalFolderRoot -PropertyType String -Value $dropboxdir

    New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails'
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails' -Name Url48x48 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_48x48.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails' -Name Url40x40 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_40x40.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails' -Name Url32x32 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_32x32.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails' -Name Url24x24 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_24x24.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails' -Name Url20x20 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_24x24.png
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\2c0ed794-6d21-4c07-9fdb-f076662715ad\Thumbnails' -Name Url16x16 -PropertyType String -Value http://dl.dropbox.com/u/46565/metro/Dropbox_16x16.png
}