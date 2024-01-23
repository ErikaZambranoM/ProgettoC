# Da usare dopo aver trovato l'id dell'item nel cestino che si desidera ripristinare tramite script ExtractInfoFromRecycleBin.ps1
param(
    [parameter(Mandatory = $true)][string]$SiteUrl #URL del sito
)

try {
    $CSVPath = (Read-Host -Prompt 'CSV Path o ID elemento').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv -Path $CSVPath -Delimiter ';'
        # Validazione colonne
		$validCols = @('Id')
		$validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
			if ($_ -in $validCols) { $validCounter++ }
		}
		if ($validCounter -lt $validCols.Count) {
			Write-Host "Colonne obbligatorie mancanti: $($validCols -join ', ')" -ForegroundColor Red
			Exit
		}
    }
    else {
        $csv = [PSCustomObject]@{
            Id    = $CSVPath
            Count = 1
        }
    }

    Connect-PnPOnline -Url $SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    $itemCounter = 0
    Write-Host 'Inizio ripristino...' -ForegroundColor Cyan
    ForEach ($item in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Ripristino' -Status "$($itemCounter+1)/$($csv.Count)" -PercentComplete (($itemCounter++ / $csv.Count) * 100) }
        try {
            Restore-PnPRecycleBinItem -Identity $item.Id -Force | Out-Null
            Write-Host "[SUCCESS] - ID: $($item.Id) - RESTORED" -ForegroundColor Green
        }
        catch { Write-Host "[WARNING] - ID: $($item.Id) - FAILED - $($_)" -ForegroundColor Yellow }
    }
    Write-Host 'Ripristino completato.' -ForegroundColor Cyan
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Ripristino' -Completed } }