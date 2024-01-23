# Import the module
Import-Module AMS-Utilities -Force
#Import-Module '.\AMS-Utilities\1.0.0\AMS-Utilities' -Force
#Import-Module 'C:\Users\ST-442\OneDrive\Desktop\WorkDesk\AMS-Utilities\AMS-Utilities\1.0.0\AMS-Utilities' -Force

################## TEST ##################
$SiteUrl = 'https://tecnimont.sharepoint.com/sites/poc_VDM'
$VDMSiteConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection
$List = 'Vendor Documents List'

# Create List object
$TestList = [SPOList]::New($List, $VDMSiteConnection)

# Load all List items
$ListItemsTest = $TestList.GetAllItems('Display') # Or Internal

# Create a column mapping to choose which columns to load and with which custom name
$ListColumnsMapping = [PSCustomObject]@{
    TCM_DN = 'Title'
    Rev    = 'Revision Number'
}
# Get all List items with specified columns
$ListItemsTest = $TestList.GetAllItems($ListColumnsMapping)

# Return output for first item of the list
$ListItemsTest[0]

# Get List vua function
$ListItemsTest = Get-SPOList -DisplayName $List -SPOConnection $VDMSiteConnection -Columns Display
