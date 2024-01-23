<#
    ToDo:
        - VDM Site URL validation
        - DD Site URL validation
#>

# Match a SharePoint Main Site or Sub Site URL
$MainOrSubSite_Regex = '(?i)^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[\w&-]+(/[\w&-]+)?/?$' # OLD: '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$'
$BaseURL = 'https://contoso.sharepoint.com' # Don't match
$BaseURL -match $MainOrSubSite_Regex # False
$MainSiteURL = 'https://contoso.sharepoint.com/Sites/MainSite' # Match
$MainSiteURL -match $MainOrSubSite_Regex # True
$SubSiteURL = 'https://contoso.sharepoint.com/Sites/MainSite/SubSite' # Match
$SubSiteURL -match $MainOrSubSite_Regex # True
$SubSiteSPOObject = 'https://contoso.sharepoint.com/Sites/MainSite/SubSite/SPOObject' # Don't match
$SubSiteSPOObject -match $MainOrSubSite_Regex # False

# Match a SharePoint Main Site URL
$MainSite_Regex = '(?i)^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[\w&-]+/?$' # OLD:'^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+/?$'
$BaseURL = 'https://contoso.sharepoint.com' # Don't match
$BaseURL -match $MainSite_Regex # False
$MainSiteURL = 'https://contoso.sharepoint.com/Sites/MainSite' # Match
$MainSiteURL -match $MainSite_Regex # True
$SubSiteURL = 'https://contoso.sharepoint.com/Sites/MainSite/SubSite' # Don't match
$SubSiteURL -match $MainSite_Regex # False
$SubSiteURLOrSPOObject = 'https://contoso.sharepoint.com/Sites/MainSite/SubSiteURLOrSPOObject' # Don't match
$SubSiteURLOrSPOObject -match $MainSite_Regex # False

# Match only a SharePoint Sub Site URL
$SubSite_Regex = '(?i)^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[\w&-]+/[\w&-]+/?$' #OLD: '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+/[\w-]+/?$'
$BaseURL = 'https://contoso.sharepoint.com' # Don't match
$BaseURL -match $SubSite_Regex # False
$MainSiteURL = 'https://contoso.sharepoint.com/Sites/MainSite' # Don't match
$MainSiteURL -match $SubSite_Regex # False
$SubSiteURL = 'https://contoso.sharepoint.com/Sites/MainSite/SubSite' # Match
$SubSiteURL -match $SubSite_Regex # True
$SubSiteSPOObject = 'https://contoso.sharepoint.com/Sites/MainSite/SubSite/SPOObject' # Don't match
$SubSiteSPOObject -match $MainOrSubSite_Regex # False

# Match a GUID (es. Power Automate Flow Id)
$FlowId = '1a2fe951-fb15-4cd3-b2a3-97804187fb33'
[guid]::TryParse($FlowId, $([ref][guid]::Empty)) -eq $true

# Match a PO Number #! PC11 has strings
$PONumber_Regex = '^\d{10}$' #or '^\d{10}(-\d)?$' for PO Number with '-1' suffix
$PONumber = '1234567890'
$PONumber -match $PONumber_Regex # True

# Match a valid PowerShell variable name
$PSVariable_Regex = '^\s*\$([a-zA-Z0-9_]+)\s*='
function Get-VariableName()
{
    $AssignedVariableName = if ($($MyInvocation.Line) -match $PSVariable_Regex) { $Matches[1] }
    Write-Host "The name of the variable this function is assigned to is:`n$AssignedVariableName"
}
$MyVariable = Get-VariableName

# Match Discipline Group or Document Library Name (one or more words separated by a space and a dash)
$Discipline_Regex = '\b[A-Z]+ - [A-Z]+\b'
'A - CIVIL' -match $Discipline_Regex # True
'AUTO - AUTOMATION' -match $Discipline_Regex # True
