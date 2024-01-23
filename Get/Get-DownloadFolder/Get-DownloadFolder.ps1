
# Function that gets the default download folder for the current user
function Get-DownloadFolder
{
    <#
    .SYNOPSIS
        Gets the default download folder for the current user.

    .DESCRIPTION
        This script queries the Windows Registry to find the path of the default download directory for the current user.

    .EXAMPLE
        .\GetDownloadFolder.ps1
    #>

    try
    {
        # Registry key path for the Shell Folders
        $KeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'

        # Get the value of the Downloads folder
        $DownloadFolder = (Get-ItemProperty -Path $KeyPath).'{374DE290-123F-4565-9164-39C4925E467B}'

        # Expand any environment variables in the path
        $ResolvedPath = [Environment]::ExpandEnvironmentVariables($DownloadFolder)

        return $ResolvedPath
    }
    catch
    {
        Throw
    }
}

# Call the function
$DownloadFolder = Get-DownloadFolder
$DownloadFolder