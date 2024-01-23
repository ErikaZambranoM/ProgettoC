#Parameters
$SourceSiteURL = 'https://tecnimont.sharepoint.com/sites/vdm_4191'
$DestinationSiteURL = 'https://tecnimont.sharepoint.com/sites/vdm_K461'
$PageName = 'Home.aspx'

#Connect to Source Site
$SourceSiteConnection = Connect-PnPOnline -Url $SourceSiteURL -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop
$SourceSiteRelativeURL = $SourceSiteURL -replace 'https://tecnimont.sharepoint.com', ''

#Export the Source page
Get-PnPFile -Url "$($SourceSiteRelativeURL)/SitePages/$($PageName)" -Path "$($Env:USERPROFILE)\Downloads" -Filename $PageName -Connection $SourceSiteConnection -AsFile -Force

#Import the page to the destination site
$DestinationSiteConnection = Connect-PnPOnline -Url $DestinationSiteURL -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop
$Folder = Get-PnPFolder -Url "$($DestinationSiteURL)/SitePages" -Connection $DestinationSiteConnection
Add-PnPFile -Path "$($Env:USERPROFILE)\Downloads\$PageName" -Folder $($Folder.ServerRelativeUrl) -NewFileName $PageName -Connection $DestinationSiteConnection