param(
    [parameter(Mandatory = $true)][string]$SiteUrl, #URL del sito
    [parameter(Mandatory = $true)][string]$ListName, #URL del sito
    [parameter(Mandatory = $false)][switch]$CountOnly #Only returns number of items on list without producing export
)

# Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

# Get all list items
Write-Host "Caricamento '$($ListName)'..." -ForegroundColor Cyan
$listItems = Get-PnPListItem -List 'Client Document List' -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID           = $_['ID']
        TCM_DN       = $_['Title']
        Rev          = $_['IssueIndex']
        ComDueDate   = $_['CommentDueDate']
        TRN          = $_['LastTransmittal']
        TRNDate      = $_['LastTransmittalDate']
        Path         = $_['DocumentsPath']
        LasClientTrn = $_['LastClientTransmittal']
    }
}
Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

Write-Host "Trovati $($listItems.Count) elementi." -ForegroundColor Cyan
If (!$CountOnly) {
    $ExecutionDate = Get-Date -Format 'yyyyMMdd-HHmmss'
    $filePath = "$($PSScriptRoot)\Log\Export_$($siteCode)-$($ExecutionDate).csv";
    if (!(Test-Path -Path $filePath)) { New-Item $filePath -Force -ItemType File | Out-Null }
    $listItems | Export-Csv -Path $filePath -Delimiter ';' -NoTypeInformation
    Write-Host "[SUCCESS] Log generato nel percorso $($filePath)" -ForegroundColor Green
}