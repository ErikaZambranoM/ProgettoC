param(
    [parameter(Mandatory = $true)][string]$SiteUrl #URL del sito
)

$urlSplit = $SiteUrl.Split('/')
$ExecutionDate = Get-Date -Format 'yyyyMMdd-HHmmss'

if ($SiteUrl.ToLower().Contains('digitaldocumentsc')) { $listType = 'DDc' }
elseif ($SiteUrl.ToLower().Contains('digitaldocuments')) { $listType = 'DD' }
elseif ($SiteUrl.ToLower().Contains('vdm_')) { $listType = 'VD' }

# Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop
$siteCode = (Get-PnPWeb).Title.Split(' ')[0]
$listArray = Get-PnPList | Where-Object -FilterScript { $_.BaseType -eq 'DocumentLibrary' -and $_.DefaultViewUrl -notmatch '[\s_]' -and $_.Title -match '\s-\s' }

$filePath = "$($PSScriptRoot)\export\export_$($siteCode)-$($listType)-$($ExecutionDate).csv"
if (!(Test-Path -Path $filePath)) { New-Item $filePath -Force -ItemType File | Out-Null }

$rowCounter = 0
Write-Host 'Inizio operazione...' -ForegroundColor Cyan
ForEach ($list in $listArray) {
    Write-Progress -Activity 'Elaborazione' -Status "$($rowCounter+1)/$($listArray.Length) - $($list.Title)" -PercentComplete (($rowCounter++ / $listArray.Length) * 100)

    $cursorTop = [Console]::CursorTop

    # Get all list items
    Write-Host "Caricamento '$($list.Title)'..." -NoNewline -ForegroundColor Cyan
    $listItems = Get-PnPListItem -List $list.Title -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            Name = $_['FileLeafRef']
            Path = $_['FileRef'] -replace "/$($urlSplit[3])/$($urlSplit[4])/", ''
            Type = $_['FSObjType'] #1 folder, 0 file
        }
    }
    [Console]::SetCursorPosition(0, $cursorTop)

    try {
        $listItems | Where-Object -FilterScript { $_.Type -eq 1 } | Select-Object Name, Path | Export-Csv -Path $filePath -Delimiter ';' -NoTypeInformation -Append
        Write-Host '[SUCCESS]' -ForegroundColor Green
    }
    catch { Write-Host "[ERROR] - $($_)" -ForegroundColor Red }
}
Write-Progress -Activity 'Elaborazione' -Completed
Write-Host "File: $($filePath)" -ForegroundColor Green