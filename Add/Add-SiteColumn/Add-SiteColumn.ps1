function Add-PnPSiteColumn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$ColumnName,

        [Parameter(Mandatory = $true)] #ToDo: Add ValidateSet
        [string]$ColumnType,

        [Parameter(Mandatory = $true)]
        [string]$Group,

        [Parameter(Mandatory = $false)]
        [string]$ContentType
    )

    try {
        # Connect to the SharePoint site
        $SiteUrl = $SiteUrl.TrimEnd('/')
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

        # Check if the column already exists
        $Field = Get-PnPField -Identity $ColumnName -ErrorAction SilentlyContinue

        if ($Field) {
            Write-Host "Column '$ColumnName' already exists in '$SiteUrl'." -ForegroundColor Yellow
        }
        else {

            # Add the site column
            $fieldCreationInformation = @{
                DisplayName  = $ColumnName
                InternalName = $ColumnName.Replace(' ', '')
                Group        = $Group
                Type         = $ColumnType
            }
            $Field = Add-PnPField @fieldCreationInformation
            Write-Host "Column '$ColumnName' added successfully to '$SiteUrl'." -ForegroundColor Green
        }
        # Add the column to the specified content type if ContentType is provided
        if ($ContentType) {
            Add-PnPFieldToContentType -Field $Field -ContentType $ContentType
            Write-Host "Column '$ColumnName' added successfully to content type '$ContentType'." -ForegroundColor Green
        }
    }
    catch {
        Throw
    }
}

$Parameters = @{
    ColumnName  = 'VendorName'
    ColumnType  = 'Text'
    Group       = 'TCM-Reply-Custom'
    ContentType = 'ClientDocumentList Content Type'
}
Add-PnPSiteColumn @Parameters -SiteUrl 'https://tecnimont.sharepoint.com/sites/A2201DigitalDocuments_testC'