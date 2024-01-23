#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2.0" }

<#
    ToDo:
        - #! Replace with read-host parameter on single site or CSV import on multiple sites
            - #! Filter only TCM sites
#>

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
        [String]
        $SiteUrl,

        [Parameter(Mandatory = $true, HelpMessage = 'Array of document libraries and lists to update')]
        [String[]] $ListsToUpdate,

        [Parameter(Mandatory = $true, HelpMessage = 'Array of columns to add to indexed columns')]
        [String[]] $NewColumnsToIndex
    )

    Try {
        # Connect to SharePoint Online
        #! Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
        $Connection = Connect-SPOSite -SiteUrl $SiteUrl

        $csvReport = @()
        $SiteUrl = $SiteUrl.TrimEnd('/')


        # Process each library/list
        ForEach ($ListToUpdate in $ListsToUpdate) {
            Write-Host "Processing list: $ListToUpdate"
            $resultReason = @()
            $AddedColumns = @()

            # Check if it's a list or library and fetch indexed columns
            $List = Get-PnPList -Identity $ListToUpdate -Connection $Connection -ErrorAction SilentlyContinue

            if ($List) {
                $fields = Get-PnPField -List $ListToUpdate -Connection $Connection
                $indexedColumns = $fields | Where-Object { $_.Indexed -eq $true }

                # Check indexed column limit
                if ($indexedColumns.Count -lt 20) {
                    # Process each column
                    ForEach ($Column in $NewColumnsToIndex) {

                        if ($indexedColumns.InternalName -notcontains $Column) {
                            Try {
                                Set-PnPField -List $ListToUpdate -Identity $column -Values @{Indexed = $true } -Connection $Connection
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
                else {
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
            $csvReport | Export-Csv -Path "$($PSScriptRoot)\$((Get-Date).ToString('dd_MM_yyyy-HH_mm_ss'))_IndexedColumnReport.csv" -NoTypeInformation -Delimiter ';'
        }
        Else {
            Write-Host 'No data to be exported.' -ForegroundColor Yellow
        }
    }
}


Function Connect-SPOSite {
    <#
    .SYNOPSIS
        Connects to a SharePoint Online Site or Sub Site.

    .DESCRIPTION
        This function connects to a SharePoint Online Site or Sub Site and returns the connection object.
        If a connection to the specified Site already exists, the function returns the existing connection object.

    .PARAMETER SiteUrl
        Mandatory parameter. Specifies the URL of the SharePoint Online site or subsite.

    .EXAMPLE
        PS C:\> Connect-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/contoso"
        This example connects to the "https://contoso.sharepoint.com/sites/contoso" site.

    .OUTPUTS
        The function returns an object with the following properties:
            - SiteUrl: The URL of the SharePoint Online site or subsite.
            - Connection: The connection object to the SharePoint Online site or subsite as returned by the Connect-PnPOnline cmdlet.
#>
    Param(
        # SharePoint Online Site URL
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                # Match a SharePoint Main Site or Sub Site URL
                If ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$') {
                    $True
                }
                Else {
                    Throw "`n'$($_)' is not a valid SharePoint Online site or subsite URL."
                }
            })]
        [String]
        $SiteUrl
    )

    Try {
        # Initialize Global:SPOConnections array if not already initialized
        If (-not $Global:SPOConnections) {
            $Global:SPOConnections = @()
        }
        Else {
            # Check if SPOConnection to specified Site already exists
            $SPOConnection = ($Global:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $SiteUrl }).Connection
        }

        # Create SPOConnection to specified Site if not already established
        If (-not $SPOConnection) {
            # Create SPOConnection to SiteURL
            $SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

            # Add SPOConnection to the list of connections
            $Global:SPOConnections += [PSCustomObject]@{
                SiteUrl    = $SiteUrl
                Connection = $SPOConnection
            }
        }

        Return $SPOConnection
    }
    Catch {
        Throw
    }
}

Try {
    $SitesList = Import-Csv -Path 'C:\Users\ST-442\Downloads\Sites.csv'

    ForEach ($Site in $SitesList) {
        Write-Host "Processing site: $($Site.SiteUrl + 'C')" -ForegroundColor Cyan
        $Connection = Connect-SPOSite -SiteUrl $Site.SiteUrl

        # Get all ClientDisciplines libraries from list 'ClientDepartmentCodeMapping' on TCM site
        $ClientDisciplinesLibraries = (Get-PnPListItem -List 'ClientDepartmentCodeMapping' -Fields ListPath -PageSize 5000 -Connection $Connection).FieldValues.ListPath | Select-Object -Unique | Sort-Object

        # If no libraries found, search for libraries on Client site by alphabet letters
        If ($null -eq $ClientDisciplinesLibraries) {
            Write-Host "'ClientDepartmentCodeMapping' empty on '$($Site.SiteUrl + 'C')'" -ForegroundColor Yellow
            Write-Host 'Searching libraries on Client site by alphabet letters...' -ForegroundColor Cyan

            # Filter out only document libraries
            $Connection = Connect-SPOSite -SiteUrl $($Site.SiteUrl + 'C')
            $ClientDisciplinesLibraries = (Get-PnPList -Connection $Connection | Where-Object {
                    $_.BaseType -eq 'DocumentLibrary' -and
                    (
                        (
                            $_.EntityTypeName -match '^[a-zA-Z]$' -or
                            $_.Title -match '^[a-zA-Z] - '
                        ) -or
                        (
                            $_.EntityTypeName -match '^[a-zA-Z]{2}$' -and
                            $_.Title -match '^[a-zA-Z]{2} - '
                        )
                    )
                }).EntityTypeName | Sort-Object
        }

        Add-IndexedColumns -SiteUrl $($Site.SiteUrl + 'C') -ListsToUpdate $ClientDisciplinesLibraries -NewColumnsToIndex 'IDDocumentList', 'IDClientDocumentList'
        Write-Host "Site '$($Site.SiteUrl + 'C')' completed.`n" -ForegroundColor Green
    }
}
Catch {
    Throw
}