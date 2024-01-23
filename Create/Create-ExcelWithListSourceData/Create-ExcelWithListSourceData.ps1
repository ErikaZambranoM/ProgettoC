<#
    AL MOMENTO FUNZIONA SOLO SU VD

    This script open on Edge the HomePage and SiteAssets of a VD Project and the (already downloaded) VendorDocumentList_Connected.xlsx file.
    Manual steps:
        On PowerQuery Editor, change the site and the List, then close without loading the result and upload it on SiteAssets
        Then use console output values to manually create the link button on the HomePage

    To Do:
        - Download Excel Template if not present on Downloads folder
            - Change template param to TemplatePrjCode
        - Make Template params not usable together
        - Adapt params to be used on DD
        - Add param for browser or use on default browser
#>

Param (
    [Parameter(Mandatory = $true)]
    [String]$PrjCode,

    [Parameter(Mandatory = $false)]
    [String]$VDLSourceExcelTemplate = 'https://tecnimont.sharepoint.com/sites/vdm_4305/_layouts/15/download.aspx?sourceurl=https%3a//tecnimont.sharepoint.com/sites/vdm_4305/SiteAssets/VendorDocumentList_Connected.xlsx',

    [Parameter(Mandatory = $false)]
    [String]$DDSourceExcelTemplate = ''
)

Connect-PnPOnline -Url $SiteURL -UseWebLogin
$ListTitle = 'Vendor Documents List'
$List = Get-PnPList -Identity $ListTitle

$PrjCode = $PrJCode.ToUpper()
$TemplateFilePath = "C:\Users\$env:USERNAME\Downloads\VendorDocumentList_Connected.xlsx"
$SiteURL = 'https://tecnimont.sharepoint.com/sites/vdm_' + $PrjCode
Start-Process $TemplateFilePath
Start-Process 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' $SiteURL
Start-Process 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' ($SiteURL + '/SiteAssets')
$BaseExcelLink = ('https://tecnimont.sharepoint.com/sites/vdm_{0}/_layouts/15/download.aspx?sourceurl=https%3a//tecnimont.sharepoint.com/sites/vdm_{0}/SiteAssets/VendorDocumentList_Connected.xlsx' -f $PrjCode)
$LinkButtonText = 'Export Vendor Document List'
$LinkButtonText | Set-Clipboard
Start-Sleep -Milliseconds 500
$BaseExcelLink | Set-Clipboard
Start-Sleep -Milliseconds 500
$PrjCode | Set-Clipboard


Write-Host ''
$List.Id.Guid
Write-Host ''
$BaseExcelLink
Write-Host ''
Write-Host $LinkButtonText