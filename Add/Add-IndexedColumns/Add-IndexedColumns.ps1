#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2.0" }

Function Add-IndexedColumns {
    <#
    .SYNOPSIS
    Add indexed columns to SharePoint Online document libraries and lists.

    .DESCRIPTION
    Add indexed columns to an array og SharePoint Online document libraries and lists.

    .PARAMETER SiteUrl
    SharePoint site URL.

    .PARAMETER ListsToUpdate
    Array of document libraries and/or lists to update.

    .PARAMETER NewColumnsToIndex
    Array of columns to add to indexed columns.

    .EXAMPLE
    Add-IndexedColumns -SiteUrl "https://contoso.sharepoint.com/sites/contoso" -ListsToUpdate "Documents", "Documents/Shared Documents" -NewColumnsToIndex "Title", "Author"
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, HelpMessage = 'SharePoint site URL')]
        [String]$SiteUrl,

        [Parameter(Mandatory = $true, HelpMessage = 'Array of document libraries and lists to update')]
        [String[]] $ListsToUpdate,

        [Parameter(Mandatory = $true, HelpMessage = 'Array of columns to add to indexed columns')]
        [String[]] $NewColumnsToIndex
    )

    Try {
        # Connect to SharePoint Online
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

        $csvReport = @()
        $SiteUrl = $SiteUrl.TrimEnd('/')


        # Process each library/list
        ForEach ($ListToUpdate in $ListsToUpdate) {
            Write-Host "Processing list: $ListToUpdate"
            $resultReason = @()
            $AddedColumns = @()

            # Check if it's a list or library and fetch indexed columns
            $List = Get-PnPList -Identity $ListToUpdate -ErrorAction SilentlyContinue

            If ($List) {
                $fields = Get-PnPField -List $ListToUpdate
                $indexedColumns = $fields | Where-Object { $_.Indexed -eq $true }

                # Check indexed column limit
                if ($indexedColumns.Count -lt 20) {
                    # Process each column
                    ForEach ($Column in $NewColumnsToIndex) {
                        If ($indexedColumns.InternalName -notcontains $Column) {
                            Try {
                                Set-PnPField -List $ListToUpdate -Identity $column -Values @{Indexed = $true }
                                $Result = 'Success'
                                $AddedColumns += $Column
                                $resultReason += "Column '$Column' indexed successfully"
                                Write-Host $resultReason -ForegroundColor Green
                            }
                            Catch {
                                $Result = 'Failed'
                                $resultReason = "Failed to index column '$Column'"
                                Write-Host $resultReason -ForegroundColor Red
                            }
                        }
                        Else {
                            $resultReason = "Column '$Column' is already indexed"
                            $Result = 'Skipped'
                            Write-Host $resultReason
                        }
                    }
                }
                Else {
                    $resultReason = "Maximum indexed columns reached for list '$ListToUpdate'"
                    $Result = 'Failed'
                    Write-Host $resultReason -ForegroundColor Red
                }
            }
            Else {
                $Result = 'Failed'
                $resultReason = "List or library not found: $ListToUpdate"
                Write-Host $resultReason -ForegroundColor Red
            }

            # Update CSV report
            $csvRow = [PSCustomObject]@{
                'SiteUrl'              = $SiteUrl
                'DocumentLibrary/List' = $ListToUpdate
                'AddedIndexedColumns'  = $($AddedColumns -join ', ')
                'Result'               = $Result
                'ResultReason'         = $($resultReason -join ', ')
            }
            $csvReport += $csvRow
        }

        Write-Host 'Document Libraries/Lists indexed columns update completed.' -ForegroundColor Green
    }
    Catch
    { Throw }
    Finally {
        If ($csvReport) {
            # Export CSV report
            $csvReport | Export-Csv -Path "$($PSScriptRoot)\Logs\$((Get-Date).ToString('dd_MM_yyyy-HH_mm_ss'))_IndexedColumnReport.csv" -NoTypeInformation -Delimiter ';'
        }
        Else {
            Write-Host 'No data to be exported.' -ForegroundColor Yellow
        }
    }
}