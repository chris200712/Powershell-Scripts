[CmdletBinding()]
[OutputType([System.Management.Automation.PSModuleInfo])]
param(
    # The name of the module to install from GitHub.
    [Parameter(Position=0, Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [System.String[]]
    $ModuleName,

    # The scope from which the module should be discoverable.
    [Parameter(Position=1)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('CurrentUser','AllUsers')]
    [System.String]
    $Scope = 'AllUsers',

    # The GitHub user name where the module was released.
    [Parameter(Position=2)]
    [ValidateNotNullOrEmpty()]
    [System.String]
    $GitHubUserName = 'KirkMunro',

    # The branch from which you want to download source code.
    [Parameter(Position=2)]
    [ValidateNotNullOrEmpty()]
    [System.String]
    $Branch = 'release'
)
try {
    foreach ($item in $ModuleName) {
        #region Extract the GUID from the hosted module manifest.

        # Create a temporary file name
        $manifestPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$','.psd1'
        try {
            Write-Progress -Activity "Installing ${item}" -Status "Downloading the manifest for module ${item}."
            # Download the module manifest
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/${GitHubUserName}/${item}/${Branch}/${item}.psd1" -OutFile $manifestPath
            # Unblock the downloaded manifest
            Unblock-File -LiteralPath $manifestPath
            # Identify the module GUID from the manifest
            $manifestContent = Get-Content -LiteralPath $manifestPath -Raw
            $manifestScriptBlock = [System.Management.Automation.ScriptBlock]::Create($manifestContent)
            $manifestHashtableAst = $manifestScriptBlock.Ast.Find({$args[0] -is [System.Management.Automation.Language.HashtableAst]},$false)
            if (-not $manifestHashtableAst) {
                $message = "The manifest for module '${item}' does not appear to be a manifest at all. No hashtable was found in the '${item}.psd1' file."
                $itemNotFoundException = New-Object -TypeName System.Management.Automation.ItemNotFoundException -ArgumentList $message
                $errorRecord = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $itemNotFoundException,$itemNotFoundException.GetType().Name,'ObjectNotFound',$item
                throw $errorRecord
            }
            if ($PSVersionTable.PSVersion -ge [System.Version]'5.0.10514.6') {
                $guidEntryTuple = $manifestHashtableAst.KeyValuePairs | Where-Object {$_.Item1.SafeGetValue() -eq 'Guid'}
                $moduleGuid = $guidEntryTuple.Item2.SafeGetValue() -as [System.Guid]
            } else {
                $moduleGuid = $null
                $guidEntryTuple = $manifestHashtableAst.KeyValuePairs | Where-Object {($_.Item1 -is [System.Management.Automation.Language.StringConstantExpressionAst]) -and ($_.Item1.Value -eq 'Guid')}
                if ($guidAst = $guidEntryTuple.Item2.Find({$args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst]},$false)) {
                    $moduleGuid = $guidAst.Value -as [System.Guid]
                }
            }
        } catch [System.Net.WebException] {
            if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) {
                $message = "The manifest for module '${item}' was not found in the GitHub repository for user '${GitHubUserName}' on the  '${Branch}' branch. Please verify your module name is correct and then try again."
                $itemNotFoundException = New-Object -TypeName System.Management.Automation.ItemNotFoundException -ArgumentList $message
                $errorRecord = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $itemNotFoundException,$itemNotFoundException.GetType().Name,'ObjectNotFound',$item
                throw $errorRecord
            }
        } finally {
            # Remove the manifest that was downloaded
            if (Test-Path -LiteralPath $manifestPath) {
                Remove-Item -LiteralPath $manifestPath
            }
        }

        # Raise an error if the GUID was not found in the manifest
        if ($moduleGuid -eq $null) {
            [System.String]$message = "Failed to identify the GUID for hosted module ${item}."
            [System.Management.Automation.ItemNotFoundException]$exception = New-Object -TypeName System.Management.Automation.ItemNotFoundException -ArgumentList $message
            [System.Management.Automation.ErrorRecord]$errorRecord = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $exception,'ItemNotFoundException',([System.Management.Automation.ErrorCategory]::ObjectNotFound),$manifestPath
            throw $errorRecord
        }

        #endregion

        #region Make sure that multiple installs or loaded assemblies won't prevent the installation from succeeding.

        # If there are multiple instances of the module we want to install/upgrade installed already, raise an error
        Write-Progress -Activity "Installing ${item}" -Status "Looking for an installed ${item} module."
        $module = Get-Module -ListAvailable | Where-Object {$_.Guid -eq $moduleGuid}
        if ($module -is [System.Array]) {
            [System.String]$message = "More than one version of ${item} is installed on this system. Manually remove the old versions and then try again."
            [System.Management.Automation.SessionStateException]$exception = New-Object -TypeName System.Management.Automation.SessionStateException -ArgumentList $message
            [System.Management.Automation.ErrorRecord]$errorRecord = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $exception,'SessionStateException',([System.Management.Automation.ErrorCategory]::InvalidOperation),$module
            throw $errorRecord
        }

        # Check to see if there are any assemblies loaded in an existing module folder before continuing
        if ($module -and
            ($assemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {$_.Location -and $_.Location.StartsWith($module.ModuleBase)})) {
            [System.String]$message = "It appears that the ${item} module was loaded at least once during this session and that an assembly it contains is still loaded. You must open a new PowerShell session where ${item} has not been loaded in order to upgrade the existing module."
            [System.Management.Automation.SessionStateException]$exception = New-Object -TypeName System.Management.Automation.SessionStateException -ArgumentList $message
            [System.Management.Automation.ErrorRecord]$errorRecord = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $exception,'SessionStateException',([System.Management.Automation.ErrorCategory]::InvalidOperation),$assemblies
            throw $errorRecord
        }

        #endregion

        #region Identify the parent folder where the module will be installed.

        Write-Progress -Activity "Installing ${item}" -Status 'Identifying the target modules folder.'
        if ($Scope -eq 'AllUsers') {
            $modulesFolder = Join-Path -Path ([System.Environment]::GetFolderPath('ProgramFiles')) -ChildPath WindowsPowerShell\Modules
            # If we're on PowerShell 3.0, make sure that the All Users modules folder is
            # in PSModulePath in the right spot if it is not already present.
            if (($PSVersionTable.PSVersion -lt [System.Version]'4.0') -and
                (-not (@($env:PSModulePath -split ';') -match "^$([System.Text.RegularExpressions.Regex]::Escape($modulesFolder))\\?$"))) {
                Write-Progress -Activity "Installing ${item}" -Status 'Adding the All Users modules folder to the PSModulePath environment variable.'
                $environmentVariableTarget = [System.EnvironmentVariableTarget]::Machine
                if ($systemPSModulePath = [System.Environment]::GetEnvironmentVariable('PSModulePath',$environmentVariableTarget) -as [System.String]) {
                    $systemPSModuleList = [System.Collections.ArrayList]@($systemPSModulePath -split ';')
                    if ($systemPSModuleList.Count -gt 1) {
                        $systemPSModuleList.Insert(0,$modulesFolder)
                    } else {
                        $systemPSModuleList.Add($modulesFolder)
                    }
                    $systemPSModulePath = $systemPSModuleList -join ';'
                    [System.Environment]::SetEnvironmentVariable('PSModulePath',$systemPSModulePath,$environmentVariableTarget)
                }
                $psModuleList = [System.Collections.ArrayList]@($env:PSModulePath -split ';')
                if ($psModuleList.Count -gt 2) {
                    $psModuleList.Insert(1,$modulesFolder)
                } else {
                    $psModuleList.Add($modulesFolder)
                }
                $env:PSModulePath = $psModuleList -join ';'
            }
        } else {
            $modulesFolder = Join-Path -Path ([System.Environment]::GetFolderPath('MyDocuments'))  -ChildPath WindowsPowerShell\Modules
        }

        # Create the modules folder if it does not exist
        if (-not (Test-Path -LiteralPath $modulesFolder)) {
            Write-Progress -Activity "Installing ${item}" -Status 'Creating modules folder.'
            New-Item -Path $modulesFolder -ItemType Directory -ErrorAction Stop > $null
        }

        #endregion

        #region Download the module and extract it into the modules folder.

        # Remove any previously extracted zip file contents from the modules folder (these
        # may have accidentally been left behind before, so we need to clean them up first)
        Join-Path -Path $modulesFolder -ChildPath "${GitHubUserName}-${item}-*" | Remove-Item -Recurse -ErrorAction Stop
        try {
            # Download and unblock the latest release from GitHub
            Write-Progress -Activity "Installing ${item}" -Status "Downloading the latest version of ${item}."
            $zipFilePath = Join-Path -Path $modulesFolder -ChildPath "${item}.zip"
            $response = Invoke-WebRequest -Uri "https://github.com/${GitHubUserName}/${item}/zipball/${Branch}" -ErrorAction Stop
            [System.IO.File]::WriteAllBytes($zipFilePath, $response.Content)
            Unblock-File -LiteralPath $zipFilePath -ErrorAction Stop
            # Extract the contents of the downloaded zip file into the modules folder
            Write-Progress -Activity "Installing ${item}" -Status "Extracting the ${item} zip file contents."
            # Check to see if we have the System.IO.Compression.FileSystem assembly installed.
            # This comes as part of .NET 4.5 and later.
            try {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
            } catch {
            }
            if ('System.IO.Compression.ZipFile' -as [System.Type]) {
                # If we have .NET 4.5 installed, use the ExtractToDirectory static method
                [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFilePath, $modulesFolder)
            } else {
                # Otherwise, use the CopyHere COM method (this is significantly slower)
                $shell = New-Object -ComObject Shell.Application
                $zip = $shell.NameSpace($zipFilePath)
                foreach($item in $zip.items()) {
                    $shell.Namespace($modulesFolder).CopyHere($item)
                }
            }
        } catch [System.Net.WebException] {
            if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) {
                $message = "The '${item}' module was not found in the GitHub repository for user '${GitHubUserName}' on the  '${Branch}' branch. Please verify your module name is correct and then try again."
                $itemNotFoundException = New-Object -TypeName System.Management.Automation.ItemNotFoundException -ArgumentList $message
                $errorRecord = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $itemNotFoundException,$itemNotFoundException.GetType().Name,'ObjectNotFound',$item
                throw $errorRecord
            }
        } finally {
            # Remove the downloaded zip file
            Write-Progress -Activity "Installing ${item}" -Status "Removing the ${item} zip file."
            Remove-Item -LiteralPath $zipFilePath
        }

        #endregion

        #region Remove the old module version if one exists.

        if ($module) {
            Write-Progress -Activity "Installing ${item}" -Status "Unloading and removing the installed ${item} module."
            # Unload the module if it is currently loaded.
            if ($loadedModule = Get-Module | Where-Object {$_.Guid -eq $module.Guid}) {
                $loadedModule | Remove-Module -ErrorAction Stop
            }
            # Remove the currently installed module.
            Remove-Item -LiteralPath $module.ModuleBase -Recurse -Force -ErrorAction Stop
        }

        #endregion

        #region Rename the extracted zip file contents folder as the module name and return the module to the caller.

        # Rename the extracted zip file contents folder as the module name
        Write-Progress -Activity "Installing ${item}" -Status "Installing the new ${item} module."
        Join-Path -Path $modulesFolder -ChildPath "${GitHubUserName}-${item}-*" `
            | Get-Item `
            | Sort-Object -Property LastWriteTime -Descending `
            | Select-Object -First 1 `
            | Rename-Item -NewName $item
        # Now return the updated module to the caller
        Get-Module -ListAvailable -Name $item

        #endregion
    }
} catch {
    $PSCmdlet.ThrowTerminatingError($_)
}