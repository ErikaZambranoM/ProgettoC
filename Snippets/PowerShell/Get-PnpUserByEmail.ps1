$SiteUrl = 'https://tecnimont.sharepoint.com/sites/DDWave2'
$UserEmail = 'x_barfed002@tecnimont.it'

$SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue

$UserEndPoint = "$SiteUrl/_api/web/SiteUsers?`$filter=Email eq '$UserEmail'"
$UserResponse = Invoke-PnPSPRestMethod -Method Get -Url $userEndPoint -Connection $SPOConnection
$UserDisplayName = $UserResponse.Value.Title
$UserDisplayName