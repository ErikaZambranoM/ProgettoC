$csv = Import-Csv -Path 'C:\Temp\prj_43U4\43U4_ClientCode.CSV' -Delimiter ';'

Connect-PnPOnline -Url 'https://tecnimont.sharepoint.com/sites/vdm_43U4' -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

Write-Host 'Caricamento Process Flow...' -ForegroundColor Cyan
$pfs = Get-PnPListItem -List 'Process Flow Status List' -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID         = $_['ID']
        TCM_DN     = $_['VD_DocumentNumber']
        Rev        = $_['VD_RevisionNumber']
        ClientCode = $_['VD_ClientDocumentNumber']
        VDL_ID     = $_['VD_VDL_ID']
    }
}

Write-Host 'Caricamento VDL...' -ForegroundColor Cyan
$VDL = Get-PnPListItem -List 'Vendor Documents List' -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID         = $_['ID']
        TCM_DN     = $_['VD_DocumentNumber']
        Rev        = $_['VD_RevisionNumber']
        ClientCode = $_['VD_ClientDocumentNumber']
    }
}

$log = New-Item -Path 'C:\Temp\prj_43U4\list1.csv' -Force -ItemType File
Add-Content $log 'TCM_DN;Rev;NewTCM_DN;NewCC'

Write-Host 'Inizio ciclo...' -ForegroundColor Cyan
ForEach ($row in $csv) {
    Write-Host "$($row.ClientCode) " -NoNewline -ForegroundColor Blue
    $vdlItem = $VDL | Where-Object -FilterScript { $_.ClientCode -eq $row.ClientCode }
    $pfsItem = $pfs | Where-Object -FilterScript { $_.VDL_ID -eq $vdlItem.ID }

    $result = [PSCustomObject]@{
        TCM_DN    = $pfsItem.TCM_DN
        Rev       = $pfsItem.Rev
        NewTCM_DN = $vdlItem.TCM_DN
        NewCC     = $vdlItem.ClientCode
    }
    Add-Content $log "$($result.TCM_DN);$($result.Rev);$($result.NewTCM_DN);$($result.NewCC)"
    Write-Host 'Aggiunto.' -ForegroundColor Green
}