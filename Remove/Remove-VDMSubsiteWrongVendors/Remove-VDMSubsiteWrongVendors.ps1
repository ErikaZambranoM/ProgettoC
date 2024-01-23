#! CHECK BEFORE RUNNIG THIS SCRIPT !#
# ToDo: Specify role parameter - ADD LOG


Pause
Pause

Connect-PnPOnline -Url 'https://tecnimont.sharepoint.com/sites/vdm_4285' -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

$vendors = Get-PnPListItem -List 'Vendors' -PageSize 5000 | ForEach-Object {
    [pscustomobject] @{
        ID        = $_['ID']
        GroupName = $_['VD_GroupName']
        SiteUrl   = $_['VD_SiteUrl']
    }
}

foreach ($vendor in $vendors) {
    $subConn = Connect-PnPOnline -Url $vendor.SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $groupList = Get-PnPGroup -Connection $subConn
    $vendorGroup = $groupList | Where-Object -FilterScript { $_.Title.Contains('VD ') }
    foreach ($group in $vendorGroup) {
        if (!($group.Title.Contains($vendor.GroupName))) {
            $groupId = (Get-PnPGroup -Identity $group.Title -Connection $subConn).Id
            try {
                Set-PnPGroupPermissions -Identity $groupId -RemoveRole 'MT Readers' -Connection $subConn
                Write-Host "Gruppo $($group.Title) rimosso da $($vendor.SiteUrl)" -ForegroundColor Green
            }
            catch { Write-Host "Gruppo $($group.Title) NON rimosso da $($vendor.SiteUrl) - $($_)" -ForegroundColor Red }
        }
    }
}