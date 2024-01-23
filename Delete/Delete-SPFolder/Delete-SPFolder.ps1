$SiteUrl = 'https://tecnimont.sharepoint.com/sites/DDWave2'
$ParentFolderSiteRelativeURL = 'StagingArea'
$FolderName = 'test'

Connect-PnPOnline -Url $SiteUrl -UseWebLogin
Remove-PnPFolder -Name $FolderName -Folder $ParentFolderSiteRelativeURL -Recycle