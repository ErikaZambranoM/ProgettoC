
# AGGIUNGERE PARAMETRO -OLDER/NEWER
# DEFAULT: MANTIENE IL PIU' VECCHIO

param (
    [Parameter(Mandatory = $True)][String]$SiteUrl,
    [Parameter(Mandatory = $True)][String]$ListName,
    [Parameter(Mandatory = $True)][String]$ColumnName
)

Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

Write-Host "Caricamento '$($ListName)'..." -ForegroundColor Cyan
$listItems = Get-PnPListItem -List $($ListName) -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID          = $_['ID']
        $ColumnName = $_[$ColumnName]
    }
}
Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

$csv = $listItems | Group-Object -Property $ColumnName | Where-Object -FilterScript { $_.Count -gt 1 }

$rowCounter = 0
Write-Host 'Inizio operazione...' -ForegroundColor Cyan
ForEach ($row in $csv) {
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Controllo' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

    ForEach ($item in ($row.Group | Select-Object -Skip 1 )) {
        try {
            Remove-PnPListItem -List $ListName -Identity $item.ID -Force -Recycle | Out-Null
            Write-Host "[SUCCESS] - List: $($ListName) - ID: $($item.ID) - Value: $($item.$ColumnName) - REMOVED" -ForegroundColor Green
        }
        catch {
            #if (!($_.Contains("Item does not exist."))) {
            Write-Host "[SUCCESS] - List: $($ListName) - ID: $($item.ID) - Value: $($item.$ColumnName) - REMOVED - $($_)" -ForegroundColor Red
        }
        #}
    }
}
Write-Progress -Activity 'Controllo' -Completed
Write-Host 'Operazione completata.' -ForegroundColor Cyan