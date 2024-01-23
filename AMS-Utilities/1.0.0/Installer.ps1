<#
    .SYNOPSIS
        Installs the AMS-Utilities module in the user's scope.

    .DESCRIPTION
        This script installs the AMS-Utilities module in the user's scope by copying the module folder to the user's scope modules path and
        cleaning up any unused files or folders in the destination.

    .PARAMETER None
        This script does not take any parameters.

    .EXAMPLE
        PS C:\> .\AMS-Utilities\1.0.0\Installer.ps1
        Installs the AMS-Utilities module in the user's scope.

    .NOTES
        This script requires the Microsoft.VisualBasic namespace to clean up unused files and folders.
#>

# Requirement to clean up unused files and folders
Using Namespace Microsoft.VisualBasic

# Main script block
Try {
    # Get user's scope modules path
    $UserModuleModulePath = $env:PSModulePath.Split(';') | Where-Object -FilterScript {
        $_ -eq [Environment]::GetFolderPath('MyDocuments') + '\PowerShell\Modules'
    }

    # Thow error if module path is not found
    If ($null -eq $UserModuleModulePath) {
        Throw "User's scope modules path not found."
    }

    # Copy module to user's scope modules path
    $ModuleFolder = $PSScriptRoot
    $ModuleRootFolder = Split-Path -Path $PSScriptRoot -Parent
    Copy-Item -Path $ModuleFolder -Destination "$($UserModuleModulePath)\AMS-Utilities" -Recurse -Force

    # Check for unused files or folders comparing source and destination
    $CopiedModule = $($UserModuleModulePath + '\AMS-Utilities')
    $SourceItems = Get-ChildItem -Path $ModuleRootFolder -Recurse | Select-Object -Property FullName, @{L = 'RelativePath'; E = { $_.FullName.Replace($ModuleRootFolder, '') } }
    $DestinationItems = Get-ChildItem -Path $CopiedModule -Recurse | Select-Object -Property FullName, PSIsContainer, @{L = 'RelativePath'; E = { $_.FullName.Replace($CopiedModule, '') } }
    $ItemsRelativePathsToDelete = Compare-Object -ReferenceObject $SourceItems.RelativePath -DifferenceObject $DestinationItems.RelativePath -PassThru | Where-Object -FilterScript { $_.SideIndicator -eq '=>' }

    # Filter all items to be deleted
    [Array]$ItemsToDelete = $DestinationItems | Where-Object -FilterScript { $_.RelativePath -in $ItemsRelativePathsToDelete }

    # Filter folders to be deleted
    [Array]$FoldersToDelete = $ItemsToDelete | Where-Object -FilterScript { $_.PSIsContainer } | Sort-Object -Property { $_.FullName.Length }

    # Delete parent folders first
    ForEach ($FolderToDelete in $FoldersToDelete) {
        # Check if the folder still exists before deleting
        If (Test-Path -Path $FolderToDelete.FullName) {
            # Delete the current folder and its subfolders
            [FileIO.FileSystem]::DeleteDirectory($FolderToDelete.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin')
        }
    }

    # Filter files to be deleted by checking if they were inside already deleted folders
    [Array]$FilesToDelete = $ItemsToDelete | Where-Object -FilterScript {
        $_.PSIsContainer -eq $false -and
        $FoldersToDelete.FullName -notcontains (Split-Path -Path $_.FullName -Parent)
    }

    # Delete remaining files
    ForEach ($FileToDelete in $FilesToDelete) {
        # Delete the file
        [FileIO.FileSystem]::DeleteFile($FileToDelete.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin')
    }
}
Catch {
    Write-Host ($_ | Out-String) -ForegroundColor Red
}