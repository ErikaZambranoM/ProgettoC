# Questo script aggiunge un permission level a un sito
# Privileges to be added are hardcoded in command Add-PnPRoleDefinition (Riga 60)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message -ForegroundColor Cyan }
}

try {
    $plName = Read-Host -Prompt 'New Permission Level name'
    $plCopy = Read-Host -Prompt 'Permission Level to copy'

    # Caricamento CSV / Sito
    $CSVPath = (Read-Host -Prompt 'CSV Path o Site URL').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv -Path $CSVPath -Delimiter ';'
        # Validazione colonne
		$validCols = @('SiteURL')
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
        $csv = [PSCustomObject] @{
            SiteURL = $CSVPath
            Count   = 1
        }
    }

    $rowCounter = 0
    Write-Log 'Inizio operazione...'
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($row.SiteURL)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

        # Connessione al sito
        Connect-PnPOnline -Url $row.SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
        Write-Log "Sito: $($row.SiteURL.Split('/')[-1])"
        
        # Recupera i permission level esistenti
        $roles = Get-PnPRoleDefinition

        # Cerca se già esistente
        $found = $roles | Where-Object -FilterScript { $_.Name -eq $plName }

        if ($null -ne $found) { Write-Log "[WARNING] - Permission Level: $($plName) - ALREADY EXISTS" }
        else {
            try {
                # Crea il nuovo permission level
                Add-PnPRoleDefinition -RoleName $plName -Clone $plCopy -Include EnumeratePermissions, BrowseDirectories -Description $plName -WarningAction Stop | Out-Null
                Write-Log "[SUCCESS] - Permission Level: $($plName) - CREATED"
            }
            catch { Write-Log "[ERROR] - Permission Level: $($plName) - FAILED - $($_)" }
        }
    }
    Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }