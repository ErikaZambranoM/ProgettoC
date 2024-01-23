#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2.0" }

# Script parameters
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory = $True, HelpMessage = 'Path of the permission matrix Excel file to import')]
    [ValidateScript({
            $TrimmedPath = $_.Trim('"').Trim("'")
            if (-Not ($TrimmedPath | Test-Path -PathType Leaf))
            {
                Throw "Invalid file path:$([Environment]::NewLine)'$TrimmedPath'"
            }
            if ($TrimmedPath -notmatch '\.xlsx?$')
            {
                Throw "Not an Excel file:$([Environment]::NewLine)'$TrimmedPath'"
            }
            Return $True
        }
    )]
    [String]
    $PermissionsMappingExcelPath,

    [Parameter(Mandatory = $False, HelpMessage = 'Before applying every permission table, shows the table in Grid-View and prompt the user for confirmation.')]
    [Switch]$ConfirmTable = $False,

    [Parameter(Mandatory = $False, HelpMessage = 'Start applying permissions from the specified table name.')]
    [ValidateSet('Site', 'DocumentLibraries', 'Lists', 'VendorSite')]
    [String]$StartFromTable,

    [Parameter(Mandatory = $False, HelpMessage = 'In case of VDM site, when processing Vendors, start applying permissions from the specified Vendor name (processed in alphabetic order).')]
    [ValidateNotNullOrEmpty()]
    [String]$StartFromVendor
)

#Region Variables

# Array of Lists and Document Libraries that must not interrupt the script if not found
$SkippableLists = @(
    'ContractorDisciplines',
    'ETR External Partners',
    'ETR Internal Distribution',
    'ETR Messages',
    'ETR Registry',
    'ETR Settings',
    'ETR Sharing areas',
    'ETR Transmittal Purpose'
    'MD Basket Areas',
    'MD Download Sessions',
    'MD Messages',
    'MD Settings',
    'TransmittalExtraColumns'
    'I4M Flow History',
    'I4M Messages',
    'I4M Settings'
)

# Array of Lists and Document Libraries that must not interrupt the script if not found
$SkippableDocumentLibraries = @(
    'ETR Temporary Sharing Area',
    'Temporary Documents Downloads',
    'Reserved',
    'User Manual'
)

# Array of default project site group names whose actual group names need to be mapped
$Script:SiteProjectGroups = @(
    'Special Senders',
    'Project Admins',
    'Project Readers',
    'Clients',
    'VDL Editors',
    'Owners',
    'Members',
    'Visitors',
    'Vendor Documents Clients',
    'Vendor Documents Project Admins ',
    'Vendor Documents Project Readers',
    'Vendor Documents VDL Editors',
    'Vendor Documents AMS Support'
)

# Standard Digital Documents template site url
$DD_STD_TemplateSite = 'https://tecnimont.sharepoint.com/sites/templateSTD_DigitalDocuments'

# KT Digital Documents template site url
$DD_KT_TemplateSite = 'https://tecnimont.sharepoint.com/sites/templateKT_DigitalDocuments'

#EndRegion Variables

#Region Functions

function Import-PermissionsMappingExcel
{
    <#
    .SYNOPSIS
    Read data from all sheets in the Excel file with permission mapping and return the data as an array of PSObjects

    .DESCRIPTION
    This function imports data from all sheets in the specified Excel file.
    It returns an array of ordered PSObjects named after the Sheets Name which contain the mapped data.

    .PARAMETER ExcelPath
    The path to the Excel file to import.

    .EXAMPLE
    Import-ExcelPermissionMapping -ExcelPath "C:\path\to\file.xlsx"

    .NOTES
    ! Sheets or Tables that are not in the $PermissionsDataMapping array are skipped.
    ! Every column whose header is not in the NonGroupHeaders array is considered to be a group name.
    If a required Sheet or Table mapped in the $PermissionsDataMapping array is not found, an error is thrown.
    Excel must be installed on the machine running the script.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $True, HelpMessage = 'Path of the Excel file to import')]
        [ValidateScript({
                $TrimmedPath = $_.Trim('"').Trim("'")
                if (-Not ($TrimmedPath | Test-Path -PathType Leaf))
                {
                    Throw "Invalid file path:$([Environment]::NewLine)'$TrimmedPath'"
                }
                if ($TrimmedPath -notmatch '\.xlsx?$')
                {
                    Throw "Not an Excel file:$([Environment]::NewLine)'$TrimmedPath'"
                }
                Return $True
            })]
        [String]
        $ExcelPath
    )

    # Initialize Excel application and variables
    Begin
    {
        Try
        {
            # Sheet to table name mapping
            $PermissionsDataMapping = @(
                [PSCustomObject]@{SheetName = 'Site'; TableName = 'Site' },
                [PSCustomObject]@{SheetName = 'Document Libraries'; TableName = 'DocumentLibraries' },
                [PSCustomObject]@{SheetName = 'Lists'; TableName = 'Lists' },
                [PSCustomObject]@{SheetName = 'Permission Levels'; TableName = 'PermissionLevels' }
            )

            $ExcelPath = $ExcelPath.Trim('"').Trim("'")
            $FileName = [System.IO.Path]::GetFileName($ExcelPath)

            if ($FileName.ToLower().Contains('vdm'))
            {
                $PermissionsDataMapping += @(
                    [PSCustomObject]@{SheetName = 'Vendor Site'; TableName = 'VendorSite' },
                    [PSCustomObject]@{SheetName = 'Vendor Site Document Libraries'; TableName = 'VendorSiteDocumentLibraries' }
                    [PSCustomObject]@{SheetName = 'Vendor Site Lists'; TableName = 'VendorSiteLists' }
                )
            }

            # Array of headers that are not group names
            $NonGroupHeaders = (
                'DL',
                'Status',
                'Inherits',
                'Contribute Required',
                'Notes'
            )

            # Create Excel COM object and open the Excel file
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $False
            $Workbook = $Excel.Workbooks.Open($ExcelPath)

            # Check that required sheets are present in the Excel file
            Compare-Object -ReferenceObject $PermissionsDataMapping.SheetName -DifferenceObject $Workbook.Sheets.GetEnumerator().Name -PassThru | ForEach-Object {
                if ($_.SideIndicator -eq '<=')
                {
                    Throw "Sheet '$($_)' not found in Excel file '$ExcelPath'."
                }
            }
        }
        Catch
        {
            # Collect the error and return it in the End block
            $ErrorCaught = $_
            $IsErrorCaught = $true
            Return
        }
    }

    # Process the Excel file
    Process
    {
        Try
        {
            # if an error was caught, skip the Process block
            if ($IsErrorCaught) { Return }

            $SheetIndex = 0
            $TemplateTablesData = [System.Collections.Generic.List[PSObject]]::New()
            $TotalSheets = $Workbook.Sheets.Count

            # Process each Worksheet
            foreach ($Worksheet in $Workbook.Sheets.GetEnumerator())
            {
                # Process only the sheets mapped in the PermissionsDataMapping array
                if ($PermissionsDataMapping.SheetName -Contains $Worksheet.Name)
                {
                    # Update the progress bar
                    $SheetIndex++
                    $WSPercentComplete = [Math]::Round(($SheetIndex / $TotalSheets) * 100)
                    $SheetsProgress = @{
                        Activity        = "Importing Excel File ($($WSPercentComplete)%)"
                        Status          = "Processing Worksheet '$($Worksheet.Name)' ($SheetIndex of $TotalSheets)"
                        PercentComplete = $WSPercentComplete
                        Id              = 0
                    }
                    Write-Progress @SheetsProgress

                    # Get the Table
                    $CurrentSheetTable = $PermissionsDataMapping[$PermissionsDataMapping.SheetName.IndexOf($Worksheet.Name)].TableName
                    Try
                    {
                        $Table = $Worksheet.ListObjects.Item($CurrentSheetTable)
                    }
                    Catch
                    {
                        Throw "Table '$CurrentSheetTable' not found in Worksheet '$($Worksheet.Name)'."
                    }

                    # Ensure that all required columns are present in the Table
                    if ($Table.Name -ne 'PermissionLevels')
                    {
                        $MissingHeaders = $NonGroupHeaders | Where-Object -FilterScript { $_ -notin ($Table.ListColumns | Select-Object Name).Name }
                        if ($MissingHeaders.Count -ne 0)
                        {
                            Throw "Missing required columns: $($MissingHeaders -join ', ')"
                        }
                    }

                    # Get the data from the Table
                    $TableData = [System.Collections.Generic.List[PSObject]]::New()
                    $RowCount = $Table.ListRows.Count
                    $ColumnsCount = $Table.ListColumns.Count

                    # Get the data rows
                    $RowIndex = 0
                    Try
                    {
                        For ($Row = 1; $Row -le $RowCount; $Row++)
                        {
                            # Update the progress bar
                            $RowIndex++
                            $RowPercentComplete = [Math]::Round(($RowIndex / $RowCount) * 100)
                            $RowProgress = @{
                                Activity        = "Importing Excel File ($($RowPercentComplete)%)"
                                Status          = "Processing Row $RowIndex of $RowCount"
                                PercentComplete = $RowPercentComplete
                                Id              = 1
                                ParentId        = 0
                            }
                            Write-Progress @RowProgress

                            # Get row data for each column
                            $GroupData = [System.Collections.Generic.List[PSObject]]::New()
                            $RowData = New-Object -TypeName PSObject
                            For ($Column = 1; $Column -le $ColumnsCount; $Column++)
                            {
                                Try
                                {
                                    # Get the header name
                                    $Header = $Table.HeaderRowRange.Cells.Item(1, $Column).Value()?.Trim()

                                    # Get the cell value and trim it to remove any leading or trailing spaces
                                    $CellValue = $Table.DataBodyRange.Cells.Item($Row, $Column).Value()?.Trim()

                                    # if the cell value is empty, then set it to N/A
                                    if (-not $CellValue)
                                    {
                                        $CellValue = 'N/A'
                                    }

                                    # if processing a mapping sheet and the header is not in the NonGroupHeaders array, then is intended to be a group name
                                    if ($NonGroupHeaders -notcontains $Header -and $Worksheet.Name -ne 'Permission Levels')
                                    {
                                        # Create a PSObject for the group and add it to the $GroupData array
                                        $GroupObject = New-Object -TypeName PSObject -Property ([Ordered]@{
                                                'GroupName'       = $Header
                                                'PermissionLevel' = $CellValue
                                            })
                                        $GroupData.Add($GroupObject)
                                    }
                                    else
                                    {
                                        # Since the header is in the NonGroupHeaders array, simply add the cell value to the $RowData object
                                        $RowData | Add-Member -MemberType NoteProperty -Name $Header -Value $CellValue
                                    }
                                }
                                Catch
                                {
                                    Throw "Error processing row $Row, column $Column in Worksheet '$($Worksheet.Name)': $($_.Exception.Message, $_.ScriptStackTrace)"
                                }
                            }

                            # When processing permission mapping sheets, only add the row data if the status is New
                            if ($Worksheet.Name -ne 'Permission Levels' -and $RowData.Status -eq 'New')
                            {
                                # if the permissions for the row are not inherited, then add the groups to update
                                if ($RowData.Inherits -eq 'No')
                                {
                                    $RowData | Add-Member -MemberType NoteProperty -Name 'GroupsToUpdate' -Value $GroupData
                                }
                                # if the permissions for the row are inherited, there is no need to store the groups to update
                                else
                                {
                                    $RowData | Add-Member -MemberType NoteProperty -Name 'GroupsToUpdate' -Value 'N/A'
                                }
                                $TableData.Add($RowData)
                            }
                            # When processing permission levels sheet, always add the row data
                            elseif ($Worksheet.Name -eq 'Permission Levels')
                            {
                                $TableData.Add($RowData)
                            }
                        }
                    }
                    Catch
                    {
                        Throw "Error processing row $Row in Worksheet '$($Worksheet.Name)': $($_.Exception.Message, $_.ScriptStackTrace)"
                    }
                    #| ? {$_.$($_.psobject.properties.name)}
                    # Combine the sheet name and data into a PSObject and add it to the $TemplateTablesData array
                    $OrderedSheetData = New-Object -TypeName PSObject -Property @{
                        ($Worksheet.Name.Replace(' ', '')) = $TableData
                    }
                    $TemplateTablesData.Add($OrderedSheetData)
                }
                else
                {
                    Write-Warning "Skipping Worksheet '$($Worksheet.Name)' as it is not in the PermissionsSheets array."
                }
            }
        }
        Catch
        {
            # Collect the error and return it in the End block
            $ErrorCaught = $_
            $IsErrorCaught = $true
            Return
        }
    }

    # Close Excel application and release COM objects
    End
    {
        Try
        {
            # Close Excel application
            $Workbook.Close($false)
            $Excel.Quit()

            # Release COM objects
            if ($Worksheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) | Out-Null }
            if ($Workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null }
            if ($Excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null }
            Remove-Variable -Name Worksheet -ErrorAction SilentlyContinue
            Remove-Variable -Name Workbook -ErrorAction SilentlyContinue
            Remove-Variable -Name Excel -ErrorAction SilentlyContinue

            # Complete the progress bar
            Write-Progress -Activity "Importing Excel File ($($RowPercentComplete)%)" -Status 'Import Complete' -Id 0 -Completed
            Write-Progress -Activity "Importing Excel File ($($WSPercentComplete)%)" -Status 'Import Complete' -Id 1 -Completed

            # Throw the error if one was caught
            if ($IsErrorCaught)
            {
                Throw $ErrorCaught
            }
        }
        Catch
        {
            Throw
        }
        Return $TemplateTablesData
    }
}

function Get-ExcelFileTitle
{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (-not (Test-Path -Path $_ -PathType Leaf))
                {
                    throw "File not found: $_"
                }
                if (-not ($_ -match '\.xlsx$'))
                {
                    throw "File is not a .xlsx Excel file: $_"
                }
                return $true
            })]
        [string]$ExcelFilePath
    )

    try
    {
        # Create a Shell.Application object
        $Shell = New-Object -ComObject Shell.Application

        # Get the folder
        $Folder = $Shell.Namespace((Get-Item $ExcelFilePath).DirectoryName)

        # Get the file
        $File = $Folder.ParseName((Get-Item $ExcelFilePath).Name)

        # Loop through properties to find the Title
        for ($i = 0; $i -lt 266; $i++)
        {
            $PropertyName = $Folder.GetDetailsOf($Folder.Items, $i)
            if ($PropertyName -eq 'Title')
            {
                $Title = $Folder.GetDetailsOf($File, $i)
                break
            }
        }
        if (-not $Title)
        {
            throw "Title metadata not found for file: $ExcelFilePath"
        }
    }
    catch
    {
        throw
    }
    finally
    {
        # Clear COM object
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null
        Remove-Variable Shell
    }

    return $Title
}

function Copy-PermissionsMappingObject
{
    <#
    .SYNOPSIS
    Creates a deep copy of the permissions mapping object.

    .DESCRIPTION
    This function creates a deep copy of the permissions mapping object generated by Import-PermissionsMappingExcel.

    .PARAMETER PermissionsMappingObject
    The permissions mapping object to be deeply copied.

    .EXAMPLE
    $deepCopiedObject = Copy-PermissionsMappingObject -PermissionsMappingObject $originalObject
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [System.Collections.Generic.List[PSObject]]
        $PermissionsMappingObject
    )

    process
    {
        $DeepCopy = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($Sheet in $PermissionsMappingObject)
        {
            $SheetProperties = $Sheet.PSObject.Properties
            $SheetCopy = New-Object PSObject

            foreach ($Property in $SheetProperties)
            {
                if ($Property.Value -is [System.Collections.Generic.List[PSObject]])
                {
                    $SubListCopy = [System.Collections.Generic.List[PSObject]]::new()
                    foreach ($Item in $Property.Value)
                    {
                        $ItemCopy = New-Object PSObject
                        foreach ($SubProp in $Item.PSObject.Properties)
                        {
                            $ItemCopy | Add-Member -MemberType NoteProperty -Name $SubProp.Name -Value $SubProp.Value
                        }
                        $SubListCopy.Add($ItemCopy)
                    }
                    $SheetCopy | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $SubListCopy
                }
                else
                {
                    $SheetCopy | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value
                }
            }
            $DeepCopy.Add($SheetCopy)
        }

        return $DeepCopy
    }
}

function Get-SiteType
{
    param(
        [Parameter(Mandatory = $True, HelpMessage = 'URL of the site to be processed')]
        [String]
        $SiteURL
    )
    try
    {
        # Vendor Documents site
        if ($SiteURL.ToLower().Contains('vdm'))
        {
            $SiteType = 'VDM'
        }
        # Digital Documents site
        else
        {
            # Compare first Discipline group found in the site with the first Discipline group in the template sites to check if the site is STD or KT
            $FirstDisciplineGroup = ($Script:SiteGroups.Title | Where-Object -FilterScript { $_ -match '[A-Z] - ' })[0]
            Switch ($FirstDisciplineGroup)
            {
                # Digital Documents Standard
                { $_ -eq $Script:DD_STD_TemplateSite_Disciplines[0] }
                {
                    $SiteTypeSuffix = 'STD'
                    break
                }
                # Digital Documents KT
                { $_ -eq $Script:DD_KT_TemplateSite_Disciplines[0] }
                {
                    $SiteTypeSuffix = 'KT'
                    break
                }
                Default { Throw "Discipline groups matching error for site '$($SiteURL)'" }
            }

            If (
                # Processed site is a Digital Documents TCM site
                (
                    $SiteURL.ToLower().Replace('_test', '').Contains('digitaldocuments') -and
                    !($SiteURL.ToLower().EndsWith('c'))
                ) -or
                $SiteURL.ToLower().EndsWith('ddwave2')
            )
            {
                $SiteType = 'DD' + $SiteTypeSuffix
            }
            elseif (
                # Processed site is a Digital Documents Client site
                (
                    $SiteURL.ToLower().Replace('_test', '').Contains('digitaldocuments') -and
                    $SiteURL.ToLower().EndsWith('c')
                ) -or
                $SiteURL.ToLower().EndsWith('ddwave2c')
            )
            {
                $SiteType = 'DDC' + $SiteTypeSuffix
            }
        }
        Return $SiteType
    }
    catch
    {
        Throw
    }
}

function Convert-PermissionsMappingPlaceholders
{
    param(
        [Parameter(Mandatory = $True, HelpMessage = 'Object containing the permission mapping data')]
        [Object[]]$PermissionsMappingObject,

        [Parameter(Mandatory = $False, HelpMessage = 'If specified, the function will process the Vendor SubSites objects only')]
        [Microsoft.SharePoint.Client.Web]$SubSite
    )
    Try
    {
        # Check if the global variables exist
        if ($Script:SiteGroups -isnot [Object[]])
        {
            Throw "Variable '`$Script:SiteGroups' is missing. Please run: `$Script:SiteGroups = Get-PnPGroup"
        }

        if ($Script:SiteType -isnot [String] )
        {
            Throw "Variable '`$Script:SiteType' is missing. Please run: `$Script:SiteType = Get-SiteType -SiteURL `$SiteURL"
        }

        if ($Script:SiteProjectGroups -isnot [Object[]])
        {
            Throw "Variable '`$Script:SiteProjectGroups' is missing. Please define your array of default or custom site groups to be matched with the site groups: `$Script:SiteProjectGroups = @('Owners', 'Visitors', 'Members')"
        }

        if ($Script:DD_STD_TemplateSite_Disciplines -isnot [Object[]])
        {
            Throw "Variable '`$Script:DD_STD_TemplateSite_Disciplines' is missing."
        }
        if ($Script:DD_KT_TemplateSite_Disciplines -isnot [Object[]])
        {
            Throw "Variable '`$Script:DD_KT_TemplateSite_Disciplines' is missing."
        }
        $Script:POLibraries = $Null
        $Script:DisciplinesDocumentLibraries = $Null

        # If $SubSite switch is specified, process SubSite objects only
        if ($SubSite)
        {
            if ($Script:SiteType -ne 'VDM')
            {
                Throw "SubSite switch is supported only on VDM site type. Type '$($Script:SiteType)' not supported"
            }

            # Find the 'DocumentLibraries' object in the $PermissionsMappingObject
            $VendorSiteDocumentLibrariesObject = $PermissionsMappingObject | Where-Object {
                $_.PSObject.Properties.Name -eq 'VendorSiteDocumentLibraries'
            }

            # Get $POLibraries
            [Array]$Script:POLibraries = @(Get-PnPList | Where-Object -FilterScript { $_.BaseTemplate -eq 101 -and $_.Title -match $Script:PONumber_Regex }).Title

            # Also search for PO libraries based on Vendor name
            $VendorNameBasedPOs = @(Get-PnPList | Where-Object -FilterScript { $_.BaseTemplate -eq 101 -and $_.Title -match $($SubSite.ServerRelativeUrl.Split('/')[-1]) }).Title
            if ($VendorNameBasedPOs.Count -gt 0)
            {
                $Script:POLibraries += $VendorNameBasedPOs
            }

            # Create a new PSObject to hold the permissions mapping for each PO library and replace the existing placeholder one
            $POLibraryTemplateRow = $VendorSiteDocumentLibrariesObject.VendorSiteDocumentLibraries | Where-Object { $_.DL -eq 'PO Lib' }
            if ($POLibraryTemplateRow)
            {
                $Script:POLibraries | ForEach-Object {
                    $Library = [PSCustomObject]@{
                        DL                 = $_
                        Status             = $POLibraryTemplateRow.Status
                        Inherits           = $POLibraryTemplateRow.Inherits
                        ContributeRequired = $POLibraryTemplateRow.'Contribute Required'
                        Notes              = $POLibraryTemplateRow.Notes
                        GroupsToUpdate     = $POLibraryTemplateRow.GroupsToUpdate
                    }
                    $VendorSiteDocumentLibrariesObject.VendorSiteDocumentLibraries.Add($Library)
                }
                $VendorSiteDocumentLibrariesObject.VendorSiteDocumentLibraries.Remove($POLibraryTemplateRow) | Out-Null
            }
            Return $PermissionsMappingObject
        }

        # Get specific data (Disciplines and/or Vendors) to be replaced based on the site type
        switch ($Script:SiteType)
        {
            # Vendor Documents Management site
            'VDM'
            {
                # Get the Discipline groups
                [Array]$DisciplinesGroups = @(Get-PnPListItem -List 'Disciplines' -PageSize 5000 | ForEach-Object { $_.FieldValues.'VD_GroupName' } ) | Sort-Object
                if ($DisciplinesGroups.Count -eq 0)
                {
                    Throw "No Discipline groups found in the list 'Disciplines'"
                }
                # Get the Vendors groups and their associated SubSites
                [Array]$Script:Vendors = @(Get-PnPListItem -List 'Vendors' -PageSize 5000 | ForEach-Object {
                        [PSCustomObject]@{
                            VendorName      = $($_.FieldValues.Title).Trim()
                            SubSiteURL      = $($_.FieldValues.'VD_SiteUrl').TrimEnd('/')
                            VendorGroupName = $($_.FieldValues.'VD_GroupName').Trim()
                        }
                    }) | Sort-Object -Property VendorGroupName #| Where-Object { $_.VendorGroupName -eq 'VD Termomeccanica Pompe' } #! DEBUG: Filter for a specific Vendor (add parameter to handle this)

                # If -StartFromVendor parameter is specified, ensure that the specified Vendor exists and set the $Script:Vendors array to start from its index
                if ($StartFromVendor)
                {
                    $VendorName = ($Script:Vendors | Where-Object -FilterScript { $_.VendorName -eq $StartFromVendor }).VendorName.ToLower()
                    if ($VendorName.Count -gt 1)
                    {
                        Throw "$($VendorName.Count) Vendors found with name '$StartFromVendor'"
                    }

                    $StartFromVendorIndex = $Script:Vendors.VendorName.ToLower().IndexOf($VendorName)
                    if ($StartFromVendorIndex -eq -1)
                    {
                        Throw "Vendor '$StartFromVendor' not found in the list 'Vendors'"
                    }
                    $Script:Vendors = $Script:Vendors[$StartFromVendorIndex..($Script:Vendors.Count - 1)]
                }

                [Array]$Script:VendorGroups = $Script:SiteGroups | Where-Object -FilterScript { $_.Title -match ($Script:Vendors.VendorGroupName -join '|') }
                if ($Script:Vendors.Count -eq 0)
                {
                    Throw "No Vendors found in the list 'Vendors'"
                }
                $DisciplinesPlaceholder = 'DS Disciplines'
                break
            }

            # Digital Documents Standard
            { $_ -in ('DDSTD', 'DDCSTD') }
            {
                $DisciplinesGroups = $Script:DD_STD_TemplateSite_Disciplines
                $DisciplinesPlaceholder = 'Discipline TCM'
                break
            }

            # Digital Documents KT
            { $_ -in ('DDKT', 'DDCKT') }
            {
                $DisciplinesGroups = $Script:DD_KT_TemplateSite_Disciplines
                $DisciplinesPlaceholder = 'Discipline TCM'
                break
            }

            Default { Throw "Site type '$($Script:SiteType)' not supported" }
        }

        # Replace Disciplines group placeholders in the $PermissionsMappingObject with the actual data
        foreach ($Table in $PermissionsMappingObject)
        {
            foreach ($Row in $Table.PSObject.Properties.Value)
            {
                # Find the placeholder for Disciplines group and its permission level
                $Placeholder = $Row.GroupsToUpdate | Where-Object { $_.GroupName -eq $DisciplinesPlaceholder }
                if ($Placeholder)
                {
                    # Replace the placeholder with actual group names
                    foreach ($GroupName in $DisciplinesGroups)
                    {
                        $NewGroup = New-Object -TypeName PSObject -Property @{
                            GroupName       = $GroupName
                            PermissionLevel = $Placeholder.PermissionLevel
                        }
                        $Row.GroupsToUpdate.Add($NewGroup)
                    }

                    # Remove the placeholder
                    $Row.GroupsToUpdate.Remove($Placeholder) | Out-Null
                }
            }
        }

        # Replace Vendors group placeholders in the $PermissionsMappingObject with the actual data
        if ($Script:SiteType -eq 'VDM')
        {
            foreach ($Table in $PermissionsMappingObject)
            {
                foreach ($Row in $Table.PSObject.Properties.Value)
                {
                    # Find the placeholder for Vendors group and its permission level
                    $Placeholder = $Row.GroupsToUpdate | Where-Object { $_.GroupName -eq 'VD Vendors' }
                    if ($Placeholder)
                    {
                        # Replace the placeholder with actual group names
                        foreach ($VendorGroup in $Script:VendorGroups)
                        {
                            $NewGroup = New-Object -TypeName PSObject -Property @{
                                GroupName       = $VendorGroup.Title
                                PermissionLevel = $Placeholder.PermissionLevel
                            }
                            $Row.GroupsToUpdate.Add($NewGroup)
                        }

                        # Remove the placeholder
                        $Row.GroupsToUpdate.Remove($Placeholder) | Out-Null
                    }
                }
            }
        }

        # Replace rows placeholders in the $PermissionsMappingObject with the actual data
        switch ($Script:SiteType)
        {
            # Vendor Documents Management site
            'VDM'
            {
                # Find the 'VendorSite' object in the $PermissionsMappingObject
                $VendorSitesObject = $PermissionsMappingObject | Where-Object {
                    $_.PSObject.Properties.Name -eq 'VendorSite'
                }

                # Create a new PSObject to hold the permissions mapping for each Vendor Site and replace the existing placeholder one
                $VendorSiteTemplateRow = $VendorSitesObject.VendorSite | Where-Object { $_.DL -eq 'VendorSite' }
                if ($VendorSiteTemplateRow)
                {
                    $Script:Vendors.SubSiteURL | ForEach-Object {
                        $VendorSite = [PSCustomObject]@{
                            DL                 = $_
                            Status             = $VendorSiteTemplateRow.Status
                            Inherits           = $VendorSiteTemplateRow.Inherits
                            ContributeRequired = $VendorSiteTemplateRow.'Contribute Required'
                            Notes              = $VendorSiteTemplateRow.Notes
                            GroupsToUpdate     = $VendorSiteTemplateRow.GroupsToUpdate
                        }
                        $VendorSitesObject.VendorSite.Add($VendorSite)
                    }
                    $VendorSitesObject.VendorSite.Remove($VendorSiteTemplateRow) | Out-Null
                }
                break
            }

            # Digital Documents Client
            { $_.Contains('DDC') }
            {
                # Find the 'DocumentLibraries' object in the $PermissionsMappingObject
                $DocumentLibrariesObject = $PermissionsMappingObject | Where-Object {
                    $_.PSObject.Properties.Name -eq 'DocumentLibraries'
                }

                # Create a new PSObject to hold the permissions mapping for each Discipline library and replace the existing placeholder one
                $AssociatedTCMSiteUrl = @(Get-PnPListItem -List 'Settings' -PageSize 5000) | ForEach-Object { ($_.FieldValues.Title -eq 'AssociatedTCMSiteUrl') ? $_.FieldValues.Value : $null } | Where-Object -FilterScript { $_ -ne $null }
                $TCMSiteURL = $SiteURL -replace '(/sites).*', $($AssociatedTCMSiteUrl)
                $TCMSiteConnection = Connect-PnPOnline -Url $TCMSiteURL -UseWebLogin -ReturnConnection -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
                $DisciplinesDocumentLibrariesPaths = @(Get-PnPListItem -List 'ClientDepartmentCodeMapping' -PageSize 5000 -Connection $TCMSiteConnection | ForEach-Object { $_.FieldValues.ListPath }) | Sort-Object -Unique
                $Script:DisciplinesDocumentLibraries = (Get-PnPList | Where-Object -FilterScript { $_.BaseTemplate -eq 101 -and $_.EntityTypeName -in $DisciplinesDocumentLibrariesPaths }).Title
                $DisciplineLibraryTemplateRow = $DocumentLibrariesObject.DocumentLibraries | Where-Object { $_.DL -eq 'Discipline Libraries' }
                if ($DisciplineLibraryTemplateRow)
                {
                    $Script:DisciplinesDocumentLibraries | ForEach-Object {
                        $Library = [PSCustomObject]@{
                            DL                 = $_
                            Status             = $DisciplineLibraryTemplateRow.Status
                            Inherits           = $DisciplineLibraryTemplateRow.Inherits
                            ContributeRequired = $DisciplineLibraryTemplateRow.'Contribute Required'
                            Notes              = $DisciplineLibraryTemplateRow.Notes
                            GroupsToUpdate     = $DisciplineLibraryTemplateRow.GroupsToUpdate
                        }
                        $DocumentLibrariesObject.DocumentLibraries.Add($Library)
                    }
                    $DocumentLibrariesObject.DocumentLibraries.Remove($DisciplineLibraryTemplateRow) | Out-Null
                }
                break
            }

            # Digital Documents TCM
            { $_.Contains('DD') }
            {
                # Find the 'DocumentLibraries' object in the $PermissionsMappingObject
                $DocumentLibrariesObject = $PermissionsMappingObject | Where-Object {
                    $_.PSObject.Properties.Name -eq 'DocumentLibraries'
                }

                # Create a new PSObject to hold the permissions mapping for each Discipline library and replace the existing placeholder one
                $Script:DisciplinesDocumentLibraries = (Get-PnPList | Where-Object -FilterScript { $_.BaseTemplate -eq 101 -and $_.EntityTypeName -match '^[A-Z]$' }).Title
                $DisciplineLibraryTemplateRow = $DocumentLibrariesObject.DocumentLibraries | Where-Object { $_.DL -eq 'Discipline Libraries' }
                if ($DisciplineLibraryTemplateRow)
                {
                    $Script:DisciplinesDocumentLibraries | ForEach-Object {
                        $Library = [PSCustomObject]@{
                            DL                 = $_
                            Status             = $DisciplineLibraryTemplateRow.Status
                            Inherits           = $DisciplineLibraryTemplateRow.Inherits
                            ContributeRequired = $DisciplineLibraryTemplateRow.'Contribute Required'
                            Notes              = $DisciplineLibraryTemplateRow.Notes
                            GroupsToUpdate     = $DisciplineLibraryTemplateRow.GroupsToUpdate
                        }
                        $DocumentLibrariesObject.DocumentLibraries.Add($Library)
                    }
                    $DocumentLibrariesObject.DocumentLibraries.Remove($DisciplineLibraryTemplateRow) | Out-Null
                }
                break
            }

            Default { Throw "Site type '$($Script:SiteType)' not supported" }
        }

        # Get the default site groups and store them in an object
        $SiteOwnersGroup = (Get-PnPGroup -AssociatedOwnerGroup).Title
        $SiteVisitorsGroup = (Get-PnPGroup -AssociatedVisitorGroup).Title
        $SiteMembersGroup = (Get-PnPGroup -AssociatedMemberGroup).Title
        $DefaultSiteGroups = [PSCustomObject]@{
            'SiteOwners'   = $SiteOwnersGroup
            'SiteVisitors' = $SiteVisitorsGroup
            'SiteMembers'  = $SiteMembersGroup
        }

        # Replace default site groups placeholders in the $PermissionsMappingObject with the actual data
        foreach ($Table in $PermissionsMappingObject)
        {
            foreach ($Row in $Table.PSObject.Properties.Value)
            {
                $GroupsToAdd = [System.Collections.Generic.List[PSObject]]::New()
                $PlaceholdersToDelete = [System.Collections.Generic.List[PSObject]]::New()
                foreach ($RowGroup in $Row.GroupsToUpdate)
                {
                    :MatchFinder foreach ($SiteProjectGroupName in $Script:SiteProjectGroups)
                    {
                        # Find the placeholders for default site groups and their permission level
                        $MatchedGroupName = $Null
                        if ($RowGroup.GroupName -match $SiteProjectGroupName)
                        {
                            $MatchingProperty = $DefaultSiteGroups | Get-Member -MemberType NoteProperty | Where-Object -FilterScript { $_.Name -match $SiteProjectGroupName } | ForEach-Object { $_.Name }
                            # Set the actual group name from $DefaultSiteGroups variable
                            if ($MatchingProperty)
                            {
                                $MatchedGroupName = $DefaultSiteGroups.$MatchingProperty
                                break MatchFinder
                            }
                            else
                            {
                                foreach ($SiteGroup in $Script:SiteGroups)
                                {
                                    # Set the actual group name by matching it with site groups
                                    if ($SiteGroup.Title -match $SiteProjectGroupName)
                                    {
                                        $MatchedGroupName = $SiteGroup.Title
                                        break MatchFinder
                                    }
                                }
                            }
                        }
                    }

                    # Replace the Placeholders with actual group names
                    if ($MatchedGroupName)
                    {
                        $SiteProjectGroup = New-Object -TypeName PSObject -Property @{
                            GroupName       = $MatchedGroupName
                            PermissionLevel = $RowGroup.PermissionLevel
                        }

                        # Store the Placeholder group names to be added and be removed
                        $GroupsToAdd.Add($SiteProjectGroup)
                        $PlaceholdersToDelete.Add($RowGroup)
                    }
                }
                # Remove the Placeholders and add the actual group names
                foreach ($Group in $GroupsToAdd)
                {
                    $Row.GroupsToUpdate.Add($Group)
                }
                foreach ($Placeholder in $PlaceholdersToDelete)
                {
                    $Row.GroupsToUpdate.Remove($Placeholder) | Out-Null
                }
            }
        }
        Return $PermissionsMappingObject
    }
    Catch { Throw }
}

function Confirm-PermissionsGridView
{
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [object]$PermissionsMappingObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InheritanceParent
    )

    $PermissionsTableGridView = $PermissionsMappingObject.$TableName | ForEach-Object {
        if ($_)
        {
            $ClonedObject = ($_).PSObject.Copy()
            foreach ($GroupColumn in $_.GroupsToUpdate )
            {
                if ($GroupColumn -eq 'N/A')
                {
                    ($PermissionsMappingObject.$InheritanceParent | Where-Object -FilterScript { $_.DL -eq $InheritanceParent -or $_.DL -match $Script:SubSite_Regex }).GroupsToUpdate | ForEach-Object {
                        $ClonedObject | Add-Member -MemberType NoteProperty -Name $_.GroupName -Value $_.PermissionLevel
                    }
                }
                else
                {
                    $ClonedObject | Add-Member -MemberType NoteProperty -Name $GroupColumn.GroupName -Value $GroupColumn.PermissionLevel
                }
            }
            $ClonedObject | Select-Object -ExcludeProperty GroupsToUpdate
        }
    }
    Write-Host "`nShowing the Permissions Mapping table for '$($TableName)' in another window.`nClose the window when ready to decide if going ahead.`n" -ForegroundColor Yellow
    $PermissionsTableGridView | Out-GridView -Title "Permissions Mapping for $($TableName)" -Wait

    # Prompt the user to confirm if he wants to continue
    $Title = 'Do you confirm?'
    $Info = ('Do you want to go ahead applying shown permission table?{0}{0}' -f
        "`n"
    )

    $YesChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', (
        'Yes{0}Continue the script.{0}{0}' -f
        "`n"
    )

    $NoChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&No', (
        'No{0}Terminate the process here.{0}{0}' -f
        "`n"
    )

    $Options = [System.Management.Automation.Host.ChoiceDescription[]] @($YesChoice, $NoChoice)
    [int]$DefaultChoice = 0
    $ChoicePrompt = $host.UI.PromptForChoice($Title, $Info, $Options, $DefaultChoice)

    Switch ($ChoicePrompt)
    {
        # Simply continue to run the script as intended
        0
        {
            Write-Host "User choose to continue the script.`n" -ForegroundColor Green
            Break
        }

        # Terminate the script
        1
        {
            Throw 'User terminated the script.'
        }

        Default { Throw 'Invalid choice' }
    }
}

function Set-SPOObjectPermission
{
    [CmdletBinding(DefaultParameterSetName = 'None', SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'SetListPermission')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ListInheritance')]
        [ValidateNotNullOrEmpty()]
        [object]$List,

        [Parameter(Mandatory = $true, ParameterSetName = 'SetSitePermission')]
        [Parameter(Mandatory = $true, ParameterSetName = 'SiteInheritance')]
        [switch]$Site,

        [Parameter(Mandatory = $true, ParameterSetName = 'SetListPermission')]
        [Parameter(Mandatory = $true, ParameterSetName = 'SetSitePermission')]
        [ValidateNotNullOrEmpty()]
        [object]$Group,

        [Parameter(Mandatory = $true, ParameterSetName = 'SetListPermission')]
        [Parameter(Mandatory = $true, ParameterSetName = 'SetSitePermission')]
        [ValidateNotNullOrEmpty()]
        [string]$PermissionLevel,

        [Parameter(Mandatory = $true, ParameterSetName = 'ListInheritance')]
        [Parameter(Mandatory = $true, ParameterSetName = 'SiteInheritance')]
        [switch]$ResetRoleInheritance,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.SharePoint.Client.SecurableObject]$SiteObject
    )

    try
    {
        # Check that at least one of the context parameters is supplied
        $ParameterSet = $PSCmdlet.ParameterSetName
        if ($ParameterSet -eq 'None')
        {
            Throw 'At least one target must be specified: -Site or -List.'
        }

        # If the Group parameter is a string, convert it to a Group object
        if ($Group -and $Group -isnot [Microsoft.SharePoint.Client.Principal])
        {
            $Group = Get-PnPGroup -Identity $Group
        }

        # If SiteObject parameter is not supplied, get the Site object (in loops, this will avoid to get the Site object multiple times)
        if (-not $SiteObject -or ($SiteObject.GetType().Name -ne 'Web'))
        {
            $SiteObject = Get-PnPWeb -Includes HasUniqueRoleAssignments, RoleAssignments
        }

        # Retrieve the Site and List object (if required) and set the target object
        If (
            $List -and
            (
                $List -isnot [Microsoft.SharePoint.Client.List] -or
                (
                    -not $List.RoleAssignments -or
                    $null -eq $List.HasUniqueRoleAssignments
                )
            )
        )
        {
            $ListObject = Get-PnPList -Identity $List -Includes HasUniqueRoleAssignments, RoleAssignments
            $TargetObject = $ListObject
        }
        elseif ($List)
        {
            $ListObject = $List
            if ($ListObject.HasUniqueRoleAssignments -eq $false -and -not $ResetRoleInheritance)
            {
                $ListObject = Get-PnPList -Identity $List -Includes HasUniqueRoleAssignments, RoleAssignments
            }
            $TargetObject = $ListObject
        }
        elseif ($Site)
        {
            $TargetObject = $SiteObject
        }
        else
        {
            Throw 'Exception not handled'
        }

        # Create an object to be returned for logging purposes
        $SplittedSiteUrl = $SiteObject.ServerRelativeUrl.Split('/')
        $Target = $ListObject.Title ?? ($SplittedSiteUrl.Count -eq 3 ? 'Site' : "SubSite $($SplittedSiteUrl[-1])")
        $OperationType = $WhatIfPreference ? 'Simulation' : 'Permission change'
        $Operation = $ResetRoleInheritance ? 'Enable inheritance' : ( $PermissionLevel -eq 'N/A' ? 'Remove permissions' : 'Assign permissions')
        $TargetGroup = $Group.Title ?? 'N/A'
        $TargetPermissionLevel = $PermissionLevel ? $PermissionLevel : 'N/A'
        $Timestamp = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
        $LogData = New-Object -TypeName PSObject -Property ([Ordered]@{
                'SiteURL'                 = $SiteObject.Url
                'DL'                      = $Target
                'Operation Type'          = $OperationType
                'Operation'               = $Operation
                'Target Group'            = $TargetGroup
                'Target Permission Level' = $TargetPermissionLevel
                'Operation Result'        = 'Unknown'
                'Timestamp'               = $Timestamp
            })

        #* Enable inheritance
        if ($ResetRoleInheritance)
        {
            try
            {
                if ($PSCmdlet.ShouldProcess($Target, 'Enable inheritance'))
                {
                    If ($TargetObject.HasUniqueRoleAssignments -eq $true)
                    {
                        $TargetObject.ResetRoleInheritance()
                        Invoke-PnPQuery
                        $Message = 'Inheritance enabled'
                    }
                    else
                    {
                        $Message = 'Inheritance already enabled'
                    }
                }
                $LogData.'Operation Result' = 'Succeeded'
            }
            catch
            {
                $LogData.'Operation Result' = 'Failed'
                throw
            }
        }
        #* Assign or remove permissions
        elseif ($Group -and $PermissionLevel)
        {
            try
            {
                # Disable inheritance, if required
                If ($TargetObject.HasUniqueRoleAssignments -ne $true)
                {
                    if ($PSCmdlet.ShouldProcess($Target, 'Disable inheritance'))
                    {
                        $TargetObject.BreakRoleInheritance($true, $false)
                        Invoke-PnPQuery
                        Write-Host "[$($Timestamp)] Inheritance disabled" -ForegroundColor Magenta
                    }
                }

                #* Remove permissions
                If ($PermissionLevel -eq 'N/A')
                {
                    if ($PSCmdlet.ShouldProcess($Target, "Remove group '$($Group.Title)'"))
                    {
                        $GroupRoleAssignment = $TargetObject.RoleAssignments | Where-Object { $_.PrincipalId -eq $Group.ID }
                        if ($GroupRoleAssignment)
                        {
                            $TargetObject.RoleAssignments.GetByPrincipal($Group).DeleteObject()
                            Invoke-PnPQuery
                            $Message = "Group '$($Group.Title)' removed"
                        }
                        else
                        {
                            $Message = "Group '$($Group.Title)' already not present"
                        }
                    }
                }
                #* Assign permissions
                else
                {
                    if ($PSCmdlet.ShouldProcess($Target, "Set '$($PermissionLevel)' permission for group '$($Group.Title)'"))
                    {
                        # Get current permission levels for the group
                        If ($Site)
                        {
                            $CurrentPermissionLevels = (Get-PnPGroupPermissions -Identity $Group.ID -ErrorAction SilentlyContinue) | Where-Object -FilterScript { $_.Hidden -ne $True }
                            $TargetListSplatting = @{}
                        }
                        elseif ($List)
                        {
                            $CurrentPermissionLevels = (Get-PnPListPermissions -PrincipalId $Group.ID -Identity $ListObject.Title -ErrorAction SilentlyContinue ) | Where-Object -FilterScript { $_.Hidden -ne $True }
                            $TargetListSplatting = @{
                                List = $Target
                            }
                        }

                        # Set specified permission level if not already set
                        if ( $Permissionlevel -notin $CurrentPermissionLevels.Name )
                        {
                            Set-PnPGroupPermissions -Identity $Group.ID -AddRole $PermissionLevel @TargetListSplatting
                        }

                        # Remove any other existing permissions, if any
                        $PermissionLevelsToRemove = $CurrentPermissionLevels | Where-Object -FilterScript { $_.Name -ne $PermissionLevel }
                        foreach ($PermissionLevelToRemove in $PermissionLevelsToRemove)
                        {
                            Set-PnPGroupPermissions -Identity $Group.ID -RemoveRole $($PermissionLevelToRemove.Name) @TargetListSplatting
                        }
                        $Message = "Permission for group '$($Group.Title)' set to '$PermissionLevel'"
                    }
                }
                $LogData.'Operation Result' = 'Succeeded'
            }
            catch
            {
                $LogData.'Operation Result' = 'Failed'
                throw
            }
        }
        else
        {
            Throw 'Exception not handled'
        }

        # Print the message, if any and return the log data
        if ($Message) { Write-Host "[$($Timestamp)] $Message" -ForegroundColor Magenta }
        Return $LogData
    }
    catch
    {
        # Check if the function is assigned to a variable and, if so, set the variable value
        $AssignedVariableName = if ($($MyInvocation.Line) -match '^\s*\$([a-zA-Z0-9_]+)\s*=') { $Matches[1] }
        If ($AssignedVariableName)
        {
            Set-Variable -Name $AssignedVariableName -Value $LogData -Scope Script -WhatIf:$false
        }
        else
        {
            $LogData
        }
        throw
    }
}

#EndRegion Functions

# Transcript-friendly error handling
Trap
{
    $_
    Try { Stop-Transcript } Catch {}
    Continue
}

#Region Main script

Try
{
    # Initialize variables
    $CSVLogData = @()
    $StartFromVendor = $StartFromVendor.Trim().ToLower()
    $ErrorActionPreference = 'Stop'
    $PermissionsMappingExcelPath = $PermissionsMappingExcelPath.Trim('"').Trim("'")
    $Script:PONumber_Regex = '^\d{10}(-\d)?$'
    $Script:Discipline_Regex = '\b[A-Z]+ - [A-Z]+\b'
    $Script:MainSite_Regex = '(?i)^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[\w&-]+/?$'
    $Script:SubSite_Regex = '(?i)^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[\w&-]+/[\w&-]+/?$'
    if ($WhatIfPreference) { $WhatIfPrefix = 'WhatIf-' } else { $WhatIfPrefix = $null }

    #Region User Prompt

    # Prompt the user for the Site URL or CSV Path
    $SiteURLs = @((Read-Host -Prompt 'Site URL / CSV Path').Trim('"').Trim("'"))

    # if the CSV Path is provided, check if it is a valid file
    if ($SiteURLs.ToLower().Contains('.csv'))
    {
        if (Test-Path -Path $SiteURLs -PathType Leaf)
        {
            $SiteURLs = Import-Csv -Path $SiteURLs -Delimiter ';'
        }
        else
        {
            Throw "File '$($SiteURLs)' cannot be found."
        }
        $CSVColumns = ($SiteURLs | Get-Member -MemberType NoteProperty).Name
        if ( $CSVColumns -notcontains 'SiteURL')
        {
            Throw "Invalid file '$($SiteURLs)'. A mandatory column is missing: SiteURL"
        }
        if ( $SiteURLs.Count -lt 1)
        {
            Throw "No site urls found on file '$($SiteURLs)'."
        }
    }
    # if the Site URL is provided, check if it is a valid URL
    elseif ($SiteURLs -eq '')
    {
        Throw 'No Site URL or CSV Path provided.'
    }
    elseif ($SiteURLs -notmatch '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+/?$')
    {
        Throw "Invalid Site URL: $SiteURLs"
    }
    else
    {
        $SiteURLs = [PSCustomObject]@{'SiteURL' = $SiteURLs }
    }

    # if -Whatif parameter has not been specified, ask confirm to user before proceeding, also providing a Whatif option
    if (-not $WhatIfPreference)
    {
        $Title = 'Do you confirm?'
        $Info = ('This operation will impact the permissions set on the following {0} site{1}:{2}{5}{2}{2}The permission matrix Excel in use is:{2}{3}{4}{4}' -f
            $($SiteURLs.Count),
            $(($SiteURLs.Count -eq 1) ? '' : 's'),
            $([Environment]::NewLine),
            $PermissionsMappingExcelPath,
            "`n",
            $($SiteURLs.SiteURL -Join "`n")
        )

        $YesChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', (
            'Yes{0}Update the following sites with provided permission matrix Excel:{0}{1}{0}{0}' -f
            "`n",
            $($SiteURLs -Join "`n")
        )

        $NoChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&No', (
            'No{0}Terminate the process without doing anything.{0}{0}' -f
            "`n"
        )

        $WhatIfChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&WhatIf', (
            'Whatif{0}Only simulate the task.{0}{0}' -f
            "`n"
        )

        $Options = [System.Management.Automation.Host.ChoiceDescription[]] @($YesChoice, $NoChoice, $WhatifChoice)
        [int]$DefaultChoice = 2
        $ChoicePrompt = $host.UI.PromptForChoice($Title, $Info, $Options, $DefaultChoice)

        Switch ($ChoicePrompt)
        {
            # Simply continue to run the script as intended
            0
            {
                $WhatIfPreference = $false
                Break
            }

            # Terminate the script
            1 { Exit 0 }

            # Set the WhatifPreference to true to only simulate the task
            2
            {
                $WhatIfPreference = $true
                $WhatIfPrefix = 'WhatIf-'
                Break
            }

            Default { Throw 'Invalid choice' }
        }
    }
    if ($WhatIfPreference)
    {
        Write-Host "User choose the WhatIf mode. No changes will be made.`n" -ForegroundColor Yellow

    }
    else
    {
        Write-Warning "User choose to continue the script and apply.`n"
    }

    #EndRegion User Prompt

    #Region Log and Stopwatch

    # Create Log folder if it doesn't exist and start transcript
    Write-Host ''
    if (-not (Test-Path -Path "$PSScriptRoot\Logs" -PathType Container))
    {
        New-Item -Path "$PSScriptRoot\Logs" -ItemType Directory -WhatIf:$false | Out-Null
    }
    $ScriptName = (Get-Item -Path $MyInvocation.MyCommand.Path).BaseName

    # Start $ScriptStopwatch to script measure execution time
    $ScriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $ScriptExecutionStartDate = (Get-Date -Format 'dd/MM/yyyy - HH:mm:ss')
    Write-Host ("Script execution start date and time: $($ScriptExecutionStartDate)") -ForegroundColor Green

    #EndRegion Log and Stopwatch

    # Get DigitalDocuments' Disciplines groups from the template sites
    Connect-PnPOnline -Url $DD_STD_TemplateSite -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    $Script:DD_STD_TemplateSite_Disciplines = (Get-PnPGroup).Title | Where-Object -FilterScript { $_ -match $Script:Discipline_Regex }
    $Script:DD_STD_TemplateSite_DisciplinesLibraries = (Get-PnPList).Title | Where-Object -FilterScript { $_ -match $Script:Discipline_Regex }
    Connect-PnPOnline -Url $DD_KT_TemplateSite -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    $Script:DD_KT_TemplateSite_Disciplines = (Get-PnPGroup).Title | Where-Object -FilterScript { $_ -match $Script:Discipline_Regex }
    $Script:DD_KT_TemplateSite_DisciplinesLibraries = (Get-PnPList).Title | Where-Object -FilterScript { $_ -match $Script:Discipline_Regex }
    Disconnect-PnPOnline -ErrorAction Stop



    # Loop through each site in the CSV file or simply process the single Site URL provided
    $PermissionsMappingObject = $null
    $SitesCounter = 0
    foreach ($SiteURL in $SiteURLs)
    {
        #Region Begin

        # Initialize variables
        $CSVRowLogData = $null
        $WarningMessage = $null
        $ProgressBarsIdsCounter = 0
        $SiteURL = $SiteURL.SiteURL.Trim('/')
        $SiteName = $SiteURL.Split('/')[-1]

        # Start $SiteStopwatch to measure execution time for the site
        $SiteStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $SiteExecutionStartDate = (Get-Date -Format 'dd/MM/yyyy - HH:mm:ss')
        Write-Host ("`nStarting processing Site '$($SiteURL)' at: $($SiteExecutionStartDate)") -ForegroundColor Green

        # Log options
        $SiteRunDateTime = Get-Date -Format 'dd-MM-yyyy_HH-mm-ss'
        $CSVLogPath = "$PSScriptRoot\Logs\$($SiteName)-$($WhatIfPrefix)$($ScriptName)_$($SiteRunDateTime).csv"
        Start-Transcript -Path "$PSScriptRoot\Logs\$($SiteName)-$($WhatIfPrefix)$($ScriptName)_$($SiteRunDateTime).log" -IncludeInvocationHeader -WhatIf:$false -Confirm:$false
        Write-Host ''

        # Update the progress bar
        $SitesCounter++
        $SitesPercentComplete = [Math]::Round(($SitesCounter / $SiteURLs.Count) * 100)
        $SitesProgress = @{
            Activity        = "Looping through sites ($($SitesPercentComplete)%)"
            Status          = "Processing site '$($SiteURL.Split('/')[-1].ToUpper())' ($SitesCounter of $($SiteURLs.Count))"
            PercentComplete = $SitesPercentComplete
            Id              = 0
        }
        Write-Progress @SitesProgress

        # Connect to the site if the Site URL is valid
        if ($SiteURL -match $Script:MainSite_Regex)
        {
            Connect-PnPOnline -Url $SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
        }
        else
        {
            Throw "Invalid Site URL: $SiteURL"
        }
        $SiteObject = Get-PnPWeb -Includes HasUniqueRoleAssignments, RoleAssignments

        #EndRegion Begin

        #Region Validate Mapping Data

        # Retrieve and convert the actual data that the placeholders in the $PermissionsMappingObject need to be replaced with
        $Script:SiteGroups = Get-PnPGroup
        $Script:SiteType = Get-SiteType -SiteURL $SiteURL
        $ExcelFileTitle = Get-ExcelFileTitle -ExcelFilePath $PermissionsMappingExcelPath -WhatIf:$false

        # Check if the Excel file is valid for the site type
        if (
            ($Script:SiteType -eq 'VDM' -and $ExcelFileTitle -ne 'VDM Site Permission Mappings') -or
            ($Script:SiteType.StartsWith('DDC') -and $ExcelFileTitle -ne 'DD Client Site Permission Mappings') -or
            (!($Script:SiteType.StartsWith('DDC')) -and $Script:SiteType.StartsWith('DD') -and $ExcelFileTitle -ne 'DD Client Site Permission Mappings')
        )
        {
            throw "Excel file '$($PermissionsMappingExcelPath)' is not valid for site type '$($Script:SiteType)'"
        }

        # Import the Excel file with the deafult permission mapping template
        if (-not $PermissionsMappingObject)
        {
            $PermissionsMappingObject = Import-PermissionsMappingExcel -ExcelPath $PermissionsMappingExcelPath -WhatIf:$false
        }

        # Clone the $PermissionsMappingObject and replace the placeholders with the actual data
        $CurrentSitePermissionMatrix = Copy-PermissionsMappingObject -PermissionsMappingObject $PermissionsMappingObject
        $CurrentSitePermissionMatrix = Convert-PermissionsMappingPlaceholders -PermissionsMappingObject $CurrentSitePermissionMatrix

        # Check if all the permission levels in the Excel file exist in the site
        $SitePermissionLevels = (Get-PnPRoleDefinition | Where-Object -FilterScript { $_.Hidden -ne $true }).Name
        $MissingPermissionLevels = Compare-Object -ReferenceObject $PermissionsMappingObject.PermissionLevels.'Permission Level' -DifferenceObject $SitePermissionLevels -PassThru | ForEach-Object {
            if ($_.SideIndicator -eq '<=' -and $_ -ne 'N/A') { $_ }
        }
        if ($MissingPermissionLevels)
        {
            if ($PSCmdlet.ShouldProcess('Permission Levels', 'Stopping for missing mandatory resources'))
            {
                Throw "The following Permission Levels are not present in the site '$($SiteURL)': $($MissingPermissionLevels -join ', ')"
            }
            else
            {
                $WarningMessage += "`n`nThe following Permission Levels are not present in the site '$($SiteURL)':`n$($MissingPermissionLevels -join "`n")`n`nABOVE PERMISSION LEVELS WILL BE IGNORED DURING THE SIMULATION BUT WILL BE REQUIRED FOR THE ACTUAL PROCESS`n"
            }
        }

        # Check if all the groups in the Excel file exist in the site
        $UpdatedExcelPermissionsGroups = @($CurrentSitePermissionMatrix | ForEach-Object { $_.$($_.PSObject.Properties.Name).GroupsToUpdate.GroupName } | Sort-Object -Unique)
        $MissingPermissionGroups = Compare-Object -ReferenceObject $UpdatedExcelPermissionsGroups -DifferenceObject $Script:SiteGroups.Title -PassThru | ForEach-Object {
            if ($_.SideIndicator -eq '<=') { $_ }
        }
        if ($MissingPermissionGroups)
        {
            $WarningMessage += "`n`nThe following Groups are not present in the site '$($SiteURL)':`n$($MissingPermissionGroups -join "`n")`n`nABOVE GROUPS WILL BE SKIPPED`n"
        }

        # Get all the Lists and Document Libraries in the site
        $SiteListsAndDocumentLibraries = Get-PnPList -Includes HasUniqueRoleAssignments, RoleAssignments

        # Check if all the Lists in the Excel file exist in the site
        $SiteLists = ($SiteListsAndDocumentLibraries | Where-Object -FilterScript { $_.BaseTemplate -notin (101, 119) }).Title
        $MissingLists = Compare-Object -ReferenceObject $([String[]]$CurrentSitePermissionMatrix.Lists.'DL') -DifferenceObject $SiteLists -PassThru | ForEach-Object {
            if ($_.SideIndicator -eq '<=') { $_ }
        }
        # Remove the lists that are stored in the $SkippableLists variable
        if ($SkippableLists)
        {
            $MissingMandatoryLists = $MissingLists | Where-Object -FilterScript { $_ -notmatch ($SkippableLists -join '|') }
            $MissingSkippableLists = $MissingLists | Where-Object -FilterScript { $_ -match ($SkippableLists -join '|') }
        }
        else
        {
            $MissingMandatoryLists = $MissingLists
            $MissingSkippableLists = $null
        }
        if ($MissingMandatoryLists)
        {
            if ($PSCmdlet.ShouldProcess('Site Lists', 'Stopping for missing mandatory resources'))
            {

                Throw "The following Lists are not present in the site '$($SiteURL)': $($MissingMandatoryLists -join ', ')"
            }
            else
            {
                $WarningMessage += "`n`nThe following Lists are not present in the site '$($SiteURL)':`n$($MissingMandatoryLists -join "`n")`n`nABOVE LISTS WILL BE IGNORED DURING THE SIMULATION BUT WILL BE REQUIRED FOR THE ACTUAL PROCESS`n"
            }
        }

        # Check if all non-Discipline Document Libraries in the Excel file exist in the site
        $SiteDocumentLibrariesNames = ($SiteListsAndDocumentLibraries | Where-Object -FilterScript { $_.BaseTemplate -in (101, 119) -and $_.Hidden -ne $True }).Title
        $NonDisciplineDocumentLibraries = $CurrentSitePermissionMatrix.DocumentLibraries.'DL' | Where-Object -FilterScript { $_ -notin $Script:DisciplinesDocumentLibraries }
        if ($NonDisciplineDocumentLibraries)
        {
            $MissingDocumentLibraries = Compare-Object -ReferenceObject $NonDisciplineDocumentLibraries -DifferenceObject $SiteDocumentLibrariesNames -PassThru | ForEach-Object {
                if (
                    $_.SideIndicator -eq '<=' -and
                    $_ -ne 'Discipline Libraries'
                ) { $_ }
            }
            # Remove the Document Libraries that are stored in the $SkippableDocumentLibraries variable
            if ($SkippableDocumentLibraries)
            {
                $MissingMandatoryDocumentLibraries = $MissingDocumentLibraries | Where-Object -FilterScript { $_ -notmatch ($SkippableDocumentLibraries -join '|') }
                $MissingSkippableDocumentLibraries = $MissingDocumentLibraries | Where-Object -FilterScript { $_ -match ($SkippableDocumentLibraries -join '|') }
            }
            else
            {
                $MissingMandatoryDocumentLibraries = $MissingDocumentLibraries
                $MissingSkippableDocumentLibraries = $null
            }
            if ($MissingMandatoryDocumentLibraries)
            {
                if ($PSCmdlet.ShouldProcess('Site Document Libraries', 'Stopping for missing mandatory resources'))
                {
                    Throw "The following Document Libraries are not present in the site '$($SiteURL)': $($MissingMandatoryDocumentLibraries -join ', ')"
                }
                else
                {
                    $WarningMessage += "`n`nThe following Document Libraries are not present in the site '$($SiteURL)':`n$($MissingMandatoryDocumentLibraries -join "`n")`n`nABOVE DOCUMENT LIBRARIES WILL BE IGNORED DURING THE SIMULATION BUT WILL BE REQUIRED FOR THE ACTUAL PROCESS`n"
                }
            }
        }

        # Check if all Discipline  Document Libraries in the Excel file exist in the site
        if ($Script:SiteType.Contains('DD'))
        {
            $MissingDisciplinesDocumentLibraries = Compare-Object -ReferenceObject $Script:DisciplinesDocumentLibraries -DifferenceObject $SiteDocumentLibrariesNames -PassThru | ForEach-Object {
                if ($_.SideIndicator -eq '<=' -and $_ -ne 'Discipline Libraries') { $_ }
            }
            if ($MissingDisciplinesDocumentLibraries)
            {
                if ($PSCmdlet.ShouldProcess('Site Discipline Document Libraries', 'Stopping for missing mandatory resources'))
                {
                    Throw "The following Disciplines Document Libraries are not present in the site '$($SiteURL)': $($MissingDisciplinesDocumentLibraries -join ', ')"
                }
                else
                {
                    $WarningMessage += "`n`nThe following Disciplines Document Libraries are not present in the site '$($SiteURL)':`n$($MissingDisciplinesDocumentLibraries -join "`n")`n`nABOVE DISCIPLINES DOCUMENT LIBRARIES WILL BE IGNORED DURING THE SIMULATION BUT WILL BE REQUIRED FOR THE ACTUAL PROCESS"
                }
            }
        }

        # Return warning messages if any
        If ($WarningMessage)
        {
            Write-Host ''
            Write-Warning $WarningMessage
        }
        If ($MissingSkippableLists)
        {
            Write-Host ''
            Write-Warning "The following missing Lists will be skipped during the process since they are stored in the 'SkippableLists' variable:`n$($MissingSkippableLists -join "`n")"
        }
        If ($MissingSkippableDocumentLibraries)
        {
            Write-Host ''
            Write-Warning "The following missing Document Libraries will be skipped during the process since they are stored in the 'SkippableDocumentLibraries' variable:`n$($MissingSkippableDocumentLibraries -join "`n")"
        }
        If ($WarningMessage -or $MissingSkippableLists -or $MissingSkippableDocumentLibraries)
        {
            $Title = 'Do you confirm?'
            $Info = "Please check above warnings before going ahead`n`n"
            $YesChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes'
            $NoChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&No'
            $Options = [System.Management.Automation.Host.ChoiceDescription[]] @($YesChoice, $NoChoice)
            [int]$DefaultChoice = 1
            $ChoicePrompt = $host.UI.PromptForChoice($Title, $Info, $Options, $DefaultChoice)

            Switch ($ChoicePrompt)
            {
                # Simply continue to run the script as intended
                0
                {
                    Break
                }

                # Terminate the script
                1
                {
                    Write-Host ''
                    Throw 'User terminated the script.'
                }

                Default { Throw 'Invalid choice' }
            }
        }
        Write-Host ''

        #EndRegion Validate Mapping Data

        # Get Tables to apply for main site
        $TablesToApply = ($CurrentSitePermissionMatrix | Where-Object -FilterScript { $_.psobject.Properties.Name -notin ('VendorSiteDocumentLibraries', 'VendorSiteLists', 'PermissionLevels') })

        # If parameter -StartFromTable is specified, skip all the Tables before the specified one
        if ($StartFromTable)
        {
            $StartingTableObject = $TablesToApply | Where-Object -FilterScript { $_.psobject.Properties.Name -eq $StartFromTable }
            $StartFromTableIndex = $TablesToApply.IndexOf($StartingTableObject)
            if ($StartFromTableIndex -lt 0)
            {
                Throw "Table '$($StartFromTable)' not found in the permission mapping object."
            }
            $TablesToSkip = ($TablesToApply[0..($StartFromTableIndex - 1)] | ForEach-Object { $_.psobject.Properties.Name }) -Join "`n"
            $TablesToApply = $TablesToApply[$StartFromTableIndex..($TablesToApply.Length - 1)]
            Write-Warning "Starting from table '$($StartFromTable)'.`n`nSkipping tables:`n$($TablesToSkip)" -WarningAction Inquire
            Write-Host ''
        }

        # Loop through each table in the permission mapping object
        $PermissionSheetsCounter = 0
        foreach ($Table in  $TablesToApply)
        {
            $InheritanceParent = 'Site'
            $TableName = $Table.psobject.Properties.Name

            # Update the progress bar
            $ProgressBarsIdsCounter = ($ProgressBarsIdsCounter -ge 1) ? $ProgressBarsIdsCounter : 1
            $PermissionSheetsCounter++
            $PermissionSheetsPercentComplete = [Math]::Round(($PermissionSheetsCounter / $TablesToApply.Count) * 100)
            $PermissionSheetsProgress = @{
                Activity        = "Looping through tables ($($PermissionSheetsPercentComplete)%)"
                Status          = "Processing table '$TableName' ($PermissionSheetsCounter of $($TablesToApply.Count))"
                PercentComplete = $PermissionSheetsPercentComplete
                Id              = 1
                ParentId        = 0
            }
            Write-Progress @PermissionSheetsProgress
            Write-Host '-------------------------------------------------------------' -ForegroundColor Blue
            Write-Host "Processing matrix table '$($TableName)'" -ForegroundColor Blue
            Write-Host '-------------------------------------------------------------' -ForegroundColor Blue

            #* Set to $true to show the Permissions Mapping table in a GridView and prompt the user to confirm if he wants to continue (default: Yes)
            if ($ConfirmTable)
            {
                Confirm-PermissionsGridView -PermissionsMappingObject $CurrentSitePermissionMatrix -TableName $TableName -InheritanceParent $InheritanceParent
            }

            # Loop through each row in the table
            $RowCounter = 0

            if ($Table.$TableName.Count -gt 0)
            {
                :RowsLoop foreach ($Row in $Table.$TableName)
                {
                    #Region DD, DDC, VDM Main Site

                    if ($Row.DL -match $Script:SubSite_Regex)
                    {
                        $ProgressBarItem = "$($Row.DL.Split('/')[-2..-1] -join '/')"
                    }
                    else
                    {
                        $ProgressBarItem = $Row.DL
                    }


                    # Update the progress bar
                    $ProgressBarsIdsCounter = ($ProgressBarsIdsCounter -ge 2) ? $ProgressBarsIdsCounter : 2
                    $RowCounter++
                    $RowsPercentComplete = [Math]::Round(($RowCounter / $($Table.$TableName.Count)) * 100)
                    $RowsProgress = @{
                        Activity        = "Looping through rows ($($RowsPercentComplete)%)"
                        Status          = "Processing '$ProgressBarItem' ($RowCounter of $($Table.$TableName.Count))"
                        PercentComplete = $RowsPercentComplete
                        Id              = 2
                        ParentId        = 1
                    }
                    Write-Progress @RowsProgress
                    Write-Host "### Processing object '$($Row.DL)' ###" -ForegroundColor Cyan

                    # Check if the row is a skippable missing object
                    $SkipReason = $null
                    $SkipWarning = $false
                    if ($Row.DL -in $MissingSkippableLists)
                    {
                        $SkipReason = 'missing and on SkippableLists variable'
                    }
                    elseif ($Row.DL -in $MissingSkippableDocumentLibraries )
                    {
                        $SkipReason = 'missing and on SkippableDocumentLibraries variable'
                    }

                    # Check if the row is a mandatory missing object, if so, stop the script unless in WhatIf mode
                    if ($Row.DL -in $MissingMandatoryLists -or $Row.DL -in $MissingMandatoryDocumentLibraries)
                    {
                        if ($PSCmdlet.ShouldProcess($($Row.DL), 'Stopping for missing mandatory resources'))
                        {
                            Throw "Object '$($Row.DL)' missing in the site '$($SiteURL)'. Initial validation failed."
                        }
                        else
                        {
                            $SkipReason = 'missing but in WhatIf mode'
                            $SkipWarning = $True
                        }
                    }

                    # Skip the row if not mandatory and not present in the site or if in WhatIf mode
                    if ($SkipReason)
                    {
                        if ($Row.DL -match $Script:MainSite_Regex)
                        {
                            $SkipTarget = 'Site'
                        }
                        elseif ($Row.DL -match $Script:SubSite_Regex)
                        {
                            $SkipTarget = "SubSite $($Row.DL[-1])"
                        }
                        else
                        {
                            $SkipTarget = $Row.DL
                        }
                        $CSVRowLogData = New-Object -TypeName PSObject -Property ([Ordered]@{
                                'SiteURL'                 = $SiteURL
                                'DL'                      = $SkipTarget
                                'Operation Type'          = $WhatIfPreference ? 'Simulation' : 'Permission change'
                                'Operation'               = 'Skip'
                                'Target Group'            = 'N/A'
                                'Target Permission Level' = 'N/A'
                                'Operation Result'        = $("Skipped because $SkipReason")
                                'Timestamp'               = $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
                            })
                        if (-not $SkipWarning)
                        {
                            Write-Warning "Skipping object '$($Row.DL)' because $SkipReason"
                        }
                        Write-Host ''
                        $CSVLogData += $CSVRowLogData
                        Continue RowsLoop
                    }

                    # Set the splatting variable for Set-SPOObjectPermission function
                    $TargetObjectSplatting = @{}
                    $TargetObjectSplatting.SiteObject = $SiteObject
                    if ($Row.Inherits -eq 'Yes')
                    {
                        $TargetObjectSplatting.ResetRoleInheritance = $true
                    }
                    if ($TableName -in ('Site', 'VendorSite'))
                    {
                        $TargetObjectSplatting.Site = $true
                        if ($TableName -eq 'VendorSite')
                        {
                            Connect-PnPOnline -Url $Row.DL -ValidateConnection -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop
                            $SubSiteObject = Get-PnPWeb -Includes HasUniqueRoleAssignments, RoleAssignments
                            $SubSiteGroups = Get-PnPGroup
                            $VendorGroupName = ($Script:Vendors | Where-Object -FilterScript { $_.SubSiteURL -eq $Row.DL }).VendorGroupName
                            $VendorGroupsToIgnore = $SubSiteGroups | Where-Object -FilterScript { $_.Title.Contains('VD ') -and $_.Title -notlike "$($VendorGroupName)*" }
                            $TargetObjectSplatting.SiteObject = $SubSiteObject
                        }
                        else
                        {
                            $SubSiteGroups = $null
                        }
                    }
                    else
                    {
                        $SubSiteGroups = $null
                        $TargetObjectSplatting.List = ($SiteListsAndDocumentLibraries | Where-Object -FilterScript { $_.Title -eq $Row.DL } ) ?? $Row.DL
                    }

                    # Apply the changes
                    if ($TargetObjectSplatting.ResetRoleInheritance)
                    {
                        #* Enable inheritance
                        $CSVRowLogData = Set-SPOObjectPermission @TargetObjectSplatting -WhatIf:$WhatIfPreference
                        $CSVLogData += $CSVRowLogData
                    }
                    else
                    {
                        #* Loop through each Group to assign or remove permissions
                        $PermissionChangesCounter = 0
                        :GroupsLoop foreach ($Group in $Row.GroupsToUpdate)
                        {
                            # Update the progress bar
                            $ProgressBarsIdsCounter = ($ProgressBarsIdsCounter -ge 3) ? $ProgressBarsIdsCounter : 3
                            $PermissionChangesCounter++
                            $PermissionChangesPercentComplete = [Math]::Round(($PermissionChangesCounter / $($Row.GroupsToUpdate.Count)) * 100)
                            $PermissionChangesProgress = @{
                                Activity        = "Looping through groups ($($PermissionChangesPercentComplete)%)"
                                Status          = "Processing group '$($Group.GroupName)' ($PermissionChangesCounter of $($Row.GroupsToUpdate.Count))"
                                PercentComplete = $PermissionChangesPercentComplete
                                Id              = 3
                                ParentId        = 2
                            }
                            Write-Progress @PermissionChangesProgress

                            # Get groups based on the context
                            if ($SubSiteGroups)
                            {
                                # Match the group name with the one in the subsite
                                $GroupObject = $SubSiteGroups | Where-Object -FilterScript {
                                    $_.Title -eq $Group.GroupName -and
                                    $_.Title -notin $MissingPermissionGroups -and
                                    $_.Title -notin $VendorGroupsToIgnore.Title
                                }
                            }
                            else
                            {
                                $GroupObject = $Script:SiteGroups | Where-Object -FilterScript { $_.Title -eq $Group.GroupName -and $_.Title -notin $MissingPermissionGroups }
                            }

                            # Proceed if Group exists in the site
                            if ($GroupObject)
                            {
                                # Set the splatting variable for Set-SPOObjectPermission function
                                $TargetObjectSplatting.Group = $GroupObject
                                $TargetObjectSplatting.PermissionLevel = $Group.PermissionLevel

                                # Set permissions
                                $CSVRowLogData = Set-SPOObjectPermission @TargetObjectSplatting -WhatIf:$WhatIfPreference
                            }
                            else
                            {
                                # Go to the next group if the current one is a different Vendor group
                                if ($Group.GroupName -in $VendorGroupsToIgnore.Title)
                                {
                                    continue GroupsLoop
                                }

                                # Compose the log data
                                if ($Row.DL -match $Script:MainSite_Regex)
                                {
                                    $SkipTarget = 'Site'
                                }
                                elseif ($Row.DL -match $Script:SubSite_Regex)
                                {
                                    $SkipTarget = "SubSite $($Row.DL[-1])"
                                }
                                else
                                {
                                    $SkipTarget = $Row.DL
                                }
                                Write-Warning "Group '$($Group.GroupName)' not found in site '$($SiteURL)' - Skipping"
                                $CSVRowLogData = New-Object -TypeName PSObject -Property ([Ordered]@{
                                        'SiteURL'                 = $SiteURL
                                        'DL'                      = $SkipTarget
                                        'Operation Type'          = $WhatIfPreference ? 'Simulation' : 'Permission change'
                                        'Operation'               = 'Skip'
                                        'Target Group'            = $($Group.GroupName)
                                        'Target Permission Level' = $($Group.PermissionLevel)
                                        'Operation Result'        = 'Skipped (Group not found)'
                                        'Timestamp'               = $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
                                    })
                            }
                            $CSVLogData += $CSVRowLogData
                        }
                        Write-Progress -Activity "Looping through groups ($($PermissionChangesPercentComplete)%)" -Status 'Complete' -Id 3 -Completed
                    }
                    Write-Host ''

                    #EndRegion DD, DDC, VDM Main Site

                    #Region Vendor SubSites

                    if ($TableName -eq 'VendorSite' -and $Script:SiteType -eq 'VDM')
                    {
                        # Initialize variables
                        $CSVRowLogData = $null
                        $InheritanceParent = 'VendorSite'
                        $CurrentSubSitePermissionMatrix = Copy-PermissionsMappingObject -PermissionsMappingObject $CurrentSitePermissionMatrix
                        $CurrentSubSitePermissionMatrix = Convert-PermissionsMappingPlaceholders -PermissionsMappingObject $CurrentSubSitePermissionMatrix -SubSite $SubSiteObject
                        $SubSiteTablesToApply = ($CurrentSubSitePermissionMatrix | Where-Object -FilterScript { $_.psobject.Properties.Name -in ('VendorSiteDocumentLibraries', 'VendorSiteLists') })

                        $SubSiteListsAndDocumentLibraries = Get-PnPList -Includes HasUniqueRoleAssignments, RoleAssignments

                        $SubSiteTableCounter = 0
                        foreach ($SubSiteTable in  $SubSiteTablesToApply)
                        {
                            $SubSiteTableName = $SubSiteTable.psobject.Properties.Name

                            # Update the progress bar
                            $ProgressBarsIdsCounter = ($ProgressBarsIdsCounter -ge 3) ? $ProgressBarsIdsCounter : 3
                            $SubSiteTableCounter++
                            $SubSiteTablePercentComplete = [Math]::Round(($SubSiteTableCounter / $SubSiteTablesToApply.Count) * 100)
                            $SubSiteTableProgress = @{
                                Activity        = "Looping through subsite tables ($($SubSiteTablePercentComplete)%)"
                                Status          = "Processing table '$SubSiteTableName' ($SubSiteTableCounter of $($SubSiteTablesToApply.Count))"
                                PercentComplete = $SubSiteTablePercentComplete
                                Id              = 3
                                ParentId        = 2
                            }
                            Write-Progress @SubSiteTableProgress
                            Write-Host '-------------------------------------------------------------' -ForegroundColor Blue
                            Write-Host "Processing matrix subsite table '$($SubSiteTableName)'" -ForegroundColor Blue
                            Write-Host '-------------------------------------------------------------' -ForegroundColor Blue

                            #* Set to $true to show the Permissions Mapping table in a GridView and prompt the user to confirm if he wants to continue (default: Yes)
                            if ($ConfirmTable)
                            {
                                Confirm-PermissionsGridView -PermissionsMappingObject $CurrentSubSitePermissionMatrix -TableName $SubSiteTableName -InheritanceParent $InheritanceParent
                            }

                            # Loop through each row in the table
                            $SubSiteRowCounter = 0
                            :SubSiteRowsLoop foreach ($SubSiteRow in $SubSiteTable.$SubSiteTableName)
                            {
                                if ($SubSiteRow.DL)
                                {
                                    # Update the progress bar
                                    $ProgressBarsIdsCounter = ($ProgressBarsIdsCounter -ge 4) ? $ProgressBarsIdsCounter : 4
                                    $SubSiteRowCounter++
                                    $SubSiteRowsPercentComplete = [Math]::Round(($SubSiteRowCounter / $($SubSiteTable.$SubSiteTableName.Count)) * 100)
                                    $SubSiteRowsProgress = @{
                                        Activity        = "Looping through rows ($($SubSiteRowsPercentComplete)%)"
                                        Status          = "Processing '$($SubSiteRow.DL)' on SubSite '$(($Row.DL.Split('/')[-1] -Join '/').ToUpper())' ($SubSiteRowCounter of $($SubSiteTable.$SubSiteTableName.Count))"
                                        PercentComplete = $SubSiteRowsPercentComplete
                                        Id              = 4
                                        ParentId        = 3
                                    }
                                    Write-Progress @SubSiteRowsProgress
                                    Write-Host "### Processing object '$($SubSiteRow.DL)' ###" -ForegroundColor Cyan

                                    # Skip the row if the object is not present in the SubSite
                                    $SubSiteList = $SubSiteListsAndDocumentLibraries | Where-Object -FilterScript { $_.Title -eq $SubSiteRow.DL }
                                    if (-not $SubSiteList)
                                    {
                                        $CSVRowLogData = New-Object -TypeName PSObject -Property ([Ordered]@{
                                                'SiteURL'                 = $Row.DL
                                                'DL'                      = $SubSiteRow.DL
                                                'Operation Type'          = $WhatIfPreference ? 'Simulation' : 'Permission change'
                                                'Operation'               = 'Skip'
                                                'Target Group'            = 'N/A'
                                                'Target Permission Level' = 'N/A'
                                                'Operation Result'        = $('Skipped because not found')
                                                'Timestamp'               = $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
                                            })
                                        Write-Warning "Skipping object '$($SubSiteRow.DL)' because not found"
                                        Write-Host ''
                                        $CSVLogData += $CSVRowLogData
                                        Continue SubSiteRowsLoop
                                    }

                                    # Set the splatting variable for Set-SPOObjectPermission function
                                    $TargetObjectSplatting = @{}
                                    $TargetObjectSplatting.SiteObject = $SubSiteObject
                                    if ($SubSiteRow.Inherits -eq 'Yes')
                                    {
                                        $TargetObjectSplatting.ResetRoleInheritance = $true
                                    }
                                    $TargetObjectSplatting.List = $SubSiteList


                                    # Apply the changes
                                    if ($TargetObjectSplatting.ResetRoleInheritance)
                                    {
                                        #* Enable inheritance
                                        $CSVRowLogData = Set-SPOObjectPermission @TargetObjectSplatting -WhatIf:$WhatIfPreference
                                        $CSVLogData += $CSVRowLogData
                                    }
                                    else
                                    {
                                        #* Loop through each Group to assign or remove permissions
                                        $PermissionChangesCounter = 0
                                        :VendorSiteGroupsLoop foreach ($Group in $SubSiteRow.GroupsToUpdate)
                                        {
                                            # Update the progress bar
                                            $ProgressBarsIdsCounter = ($ProgressBarsIdsCounter -ge 5) ? $ProgressBarsIdsCounter : 5
                                            $PermissionChangesCounter++
                                            $PermissionChangesPercentComplete = [Math]::Round(($PermissionChangesCounter / $($SubSiteRow.GroupsToUpdate.Count)) * 100)
                                            $PermissionChangesProgress = @{
                                                Activity        = "Looping through groups ($($PermissionChangesPercentComplete)%)"
                                                Status          = "Processing group '$($Group.GroupName)' ($PermissionChangesCounter of $($SubSiteRow.GroupsToUpdate.Count))"
                                                PercentComplete = $PermissionChangesPercentComplete
                                                Id              = 5
                                                ParentId        = 4
                                            }
                                            Write-Progress @PermissionChangesProgress

                                            # Get groups based on the context
                                            $GroupObject = $null
                                            if ($SubSiteGroups)
                                            {
                                                # Check if we're trying to assign a Vendor group to a PO Library
                                                if ($Group.GroupName.StartsWith('VD') -and $SubSiteRow.DL -match ($Script:POLibraries -join '|'))
                                                {
                                                    # If the Vendor group is the default one, match it
                                                    if ($Group.GroupName -eq $VendorGroupName)
                                                    {
                                                        $GroupObject = $SubSiteGroups | Where-Object -FilterScript {
                                                            $_.Title -eq $VendorGroupName -and
                                                            $_.Title -notin $MissingPermissionGroups -and
                                                            $_.Title -notin $VendorGroupsToIgnore.Title
                                                        }
                                                    }
                                                    # Else the Vendor group is PO specific
                                                    elseif ($Group.GroupName.EndsWith($SubSiteRow.DL))
                                                    {
                                                        $GroupObject = $SubSiteGroups | Where-Object -FilterScript {
                                                            $_.Title.EndsWith($SubSiteRow.DL) -and
                                                            $_.Title -notin $MissingPermissionGroups -and
                                                            $_.Title -notin $VendorGroupsToIgnore.Title
                                                        }
                                                    }
                                                    # Else the Vendor group is PO specific but not valid for the current PO Library
                                                    else
                                                    {
                                                        continue VendorSiteGroupsLoop
                                                    }
                                                }
                                                else
                                                {
                                                    # Match the group name with the one in the subsite
                                                    $GroupObject = $SubSiteGroups | Where-Object -FilterScript {
                                                        $_.Title -eq $Group.GroupName -and
                                                        $_.Title -notin $MissingPermissionGroups -and
                                                        $_.Title -notin $VendorGroupsToIgnore.Title
                                                    }
                                                }
                                            }

                                            # Proceed if Group exists in the site
                                            if ($GroupObject)
                                            {
                                                # Set the splatting variable for Set-SPOObjectPermission function
                                                $TargetObjectSplatting.Group = $GroupObject
                                                $TargetObjectSplatting.PermissionLevel = $Group.PermissionLevel

                                                # Set permissions
                                                $CSVRowLogData = Set-SPOObjectPermission @TargetObjectSplatting -WhatIf:$WhatIfPreference
                                            }
                                            else
                                            {
                                                # Go to the next group if the current one is a different Vendor group
                                                if ($Group.GroupName -in $VendorGroupsToIgnore.Title)
                                                {
                                                    continue VendorSiteGroupsLoop
                                                }

                                                # Compose the log data
                                                Write-Warning "Group '$($Group.GroupName)' not found in subsite '$($SiteURL)' - Skipping"
                                                $CSVRowLogData = New-Object -TypeName PSObject -Property ([Ordered]@{
                                                        'SiteURL'                 = $Row.DL
                                                        'DL'                      = $SubSiteRow.DL
                                                        'Operation Type'          = $WhatIfPreference ? 'Simulation' : 'Permission change'
                                                        'Operation'               = 'Unknown'
                                                        'Target Group'            = $($Group.GroupName)
                                                        'Target Permission Level' = $($Group.PermissionLevel)
                                                        'Operation Result'        = 'Unknown'
                                                        'Timestamp'               = 'Unknown'
                                                    })
                                                $CSVRowLogData.Operation = 'Skip'
                                                $CSVRowLogData.'Operation Result' = 'Skipped (Group not found)'
                                                $CSVRowLogData.Timestamp = $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
                                            }
                                            $CSVLogData += $CSVRowLogData
                                        }
                                        Write-Progress -Activity "Looping through groups ($($PermissionChangesPercentComplete)%)" -Status 'Complete' -Id 5 -Completed
                                    }
                                }
                                else
                                {
                                    Write-Host '### No object object to process ###' -ForegroundColor Cyan
                                }
                                Write-Host ''
                            }
                            Write-Progress -Activity "Looping through rows ($($RowsPercentComplete)%)" -Status 'Complete' -Id 4 -Completed
                        }
                        Write-Progress -Activity "Looping through groups ($($PermissionChangesPercentComplete)%)" -Status 'Complete' -Id 3 -Completed
                    }

                    #EndRegion Vendor SubSites
                }
            }
            else
            {
                Write-Host '### No object object to process ###' -ForegroundColor Cyan
                Write-Host ''
            }
            Write-Progress -Activity "Looping through rows ($($RowsPercentComplete)%)" -Status 'Complete' -Id 2 -Completed
        }
        Write-Host ''
        Write-Progress -Activity "Looping through tables ($($PermissionSheetsPercentComplete)%)" -Status 'Complete' -Id 1 -Completed

        # Stop $SiteStopwatch to measure execution time for the current site
        $SiteStopwatch.Stop()
        $SiteElapsedTime = $(Get-Date -Date $($SiteStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
        $SiteExecutionEndDate = (Get-Date -Format 'dd/MM/yyyy - HH:mm:ss')
        Write-Host ("Site processing time: $($SiteElapsedTime)") -ForegroundColor Green
        $CSVLogData | Export-Csv -Path $CSVLogPath -Encoding UTF8 -Delimiter ';'
        Write-Host ("`nEnded processing Site '$($SiteURL)' at: $($SiteExecutionEndDate)`n") -ForegroundColor Green
        Stop-Transcript
        $CSVLogData = $null
    }
    Write-Progress -Activity "Looping through sites ($($SitesPercentComplete)%)" -Status 'Complete' -Id 0 -Completed
    $ConsoleOutputColor = 'Green'
    Write-Host "Script completed successfully.`n" -ForegroundColor Green
}
Catch
{
    # Complete all active progress bars and throw the error
    0..$ProgressBarsIdsCounter | ForEach-Object { Write-Progress -Activity 'Error caught' -Status 'Complete' -Id $_ -Completed }
    $ConsoleOutputColor = 'Red'
    Throw
}
Finally
{
    $WhatIfPreference = $false
    if ($CSVLogData)
    {
        $CSVLogData | Export-Csv -Path $CSVLogPath -Encoding UTF8 -Delimiter ';'
    }

    # Stop $ScriptStopwatch to script measure execution time
    if ($SiteStopwatch) { $SiteStopwatch.Stop() }
    $ScriptStopwatch.Stop()
    $ScriptElapsedTime = $(Get-Date -Date $($ScriptStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
    $ScriptExecutionEndDate = (Get-Date -Format 'dd/MM/yyyy - HH:mm:ss')
    Write-Host ('Script execution time: {0}' -f $ScriptElapsedTime) -ForegroundColor $ConsoleOutputColor
    Write-Host ("Script execution ended at: $($ScriptExecutionEndDate)`n") -ForegroundColor $ConsoleOutputColor
}

#EndRegion Main script

# If no error occurred, we still need to stop the transcript
Try { Stop-Transcript } Catch {}