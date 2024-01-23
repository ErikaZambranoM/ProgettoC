#Requires -Version 7 -Modules PnP.PowerShell

<#
    ToDo:
        - Parameters fo single site or csv sites
        - Progress bar for csv
#>

Param (
    [Parameter(Mandatory = $true, HelpMessage = 'Path to the CSV with the list of sites to export permissions from')]
    [String]
    $CsvPath
)

# Import required modules
Import-Module PnP.PowerShell

# Function to export SharePoint Online site permissions to CSV
Function Export-SPOSitePermissionsToCSV {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = 'SharePoint Site URL')]
        [string] $SiteUrl,

        [Parameter(Mandatory = $true, HelpMessage = 'Path where to save CSV file')]
        [string] $CsvPath
    )

    Begin {
        Write-Host 'Initializing SharePoint connection...'
        Try {
            Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue
        }
        Catch {
            Throw
        }
    }

    Process {
        Write-Host "Fetching permissions for site '$($SiteName)'..." -ForegroundColor Cyan
        $permissionEntries = @()

        Try {
            $context = Get-PnPContext
            $web = Get-PnPWeb -Includes RoleAssignments
            $context.Load($web.RoleAssignments)
            $context.ExecuteQuery()

            $index = 0
            $totalItems = $web.RoleAssignments.Count

            foreach ($roleAssignment in $web.RoleAssignments) {
                $index++
                Write-Progress -PercentComplete (($index / $totalItems) * 100) -Status 'Processing' -Activity 'Getting site permissions...' -CurrentOperation "$index of $totalItems"

                $context.Load($roleAssignment.Member)
                $context.Load($roleAssignment.RoleDefinitionBindings)
                $context.ExecuteQuery()

                $user = $roleAssignment.Member.Title
                $role = ($roleAssignment.RoleDefinitionBindings | Where-Object -FilterScript { $_.RoleTypeKind -ne 'Guest' }).Name #.Remove("Limited Access")

                If ($role) {
                    $permissionEntries += [PSCustomObject]@{
                        Group   = $user
                        Level   = $role -join ';'
                        SiteURL = $SiteUrl
                    }
                }
            }
        }
        Catch {
            Throw 'An error occurred while fetching permissions.'
        }
    }

    End {
        Write-Host 'Exporting to CSV...'
        Try {
            Write-Progress -Completed -Activity 'Getting site permissions...'
            $permissionEntries | Export-Csv -Path $CsvPath -NoTypeInformation -Delimiter ';'
        }
        Catch {
            Throw
        }

        Write-Host "Permissions exported successfully.`n" -ForegroundColor Green
    }
}
Try {
    $CsvPath = $CsvPath -replace '"', '' -replace "'", ''
    $Sites = Import-Csv -Path $CsvPath -Delimiter ';'
    $CSVFolderPath = Split-Path -Path $CsvPath -Parent
    $DateTime = Get-Date -Format 'MM-dd-yyyy_HH-mm-ss'
    $CSVExportFolderPath = (New-Item -Path $CSVFolderPath -Name "SitesPermissions_$($DateTime)" -ItemType 'Directory').FullName

    ForEach ($Site in $Sites) {
        $SiteUrl = $Site.Url
        $SiteName = $SiteUrl.Split('/')[-1]
        $CsvExportPath = Join-Path -Path $CSVExportFolderPath -ChildPath "$($SiteName).csv"
        Export-SPOSitePermissionsToCSV -SiteUrl $SiteUrl -CsvPath $CsvExportPath
    }
}
Catch {
    Throw
}

