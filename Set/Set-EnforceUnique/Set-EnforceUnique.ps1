# AGGIUNGERE VALIDAZIONE LISTE SUPPORTATE (PROTOCOL REGISTRY)

$columnName = 'Title'

$csvPath = (Read-Host -Prompt 'CSV Path').Trim('"')
$csv = Import-Csv -Path $csvPath -Delimiter ';'

$rowCounter = 0
Write-Host 'Inizio operazione' -ForegroundColor Cyan
ForEach ($row in $csv) {
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Controllo' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

    Connect-PnPOnline -Url $row.URL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    if ($row.URL.ToLower().Contains('digitaldocumentsc')) {
        $listName = 'ClientProtocolRegistry'
    }
    else {
        $listName = 'ProtocolRegistry'
    }

    $output = Get-PnPField -List $listName -Identity $columnName | Select-Object EnforceUniqueValues, Indexed

    Write-Host "Site: $($row.URL.Split('/')[-1]) - EnforceUniqueValues: $($output.EnforceUniqueValues)" -ForegroundColor Cyan

    if ($output.EnforceUniqueValues -ne $true) {
        try {
            if ($output.Indexed -eq $false) {
                Set-PnPField -List $listName -Identity $columnName -Values @{ Indexed = $true }
            }
            Set-PnPField -List $listName -Identity $columnName -Values @{ EnforceUniqueValues = $true }
            Write-Host "[SUCCESS] - List: $($listName) - Column: $($columnName) - UPDATED" -ForegroundColor Green
        }
        catch { Write-Host "[SUCCESS] - List: $($listName) - Column: $($columnName) - FAILED - $($_)" -ForegroundColor Red }
    }
    else {
        Write-Host "[WARNING] - List: $($listName) - Column: $($columnName) - ALREADY TRUE" -ForegroundColor Yellow
    }
}
Write-Progress -Activity 'Controllo' -Completed
Write-Host 'Operazione completata' -ForegroundColor Cyan