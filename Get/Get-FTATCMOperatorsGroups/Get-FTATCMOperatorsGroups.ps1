# Funzione di log to CSV
Function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    $ExecutionDate = Get-Date -Format 'MM_dd_yyyy'
    $logPath = "$($PSScriptRoot)\logs\$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;Site;App Name;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - Site: ', ';').Replace(' - App: ', ';').Replace(' - Package: ', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}


Try {

    # Caricamento CSV o sito singolo
    $CSVPath = Read-Host -Prompt 'CSV Path o Site Url'
    if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
    elseif ($CSVPath -ne '') {
        $csv = [PSCustomObject]@{
            Site  = $CSVPath
            Count = 1
        }
    }

    else { Exit }
}
catch {
    Write-Log '[ERROR] Input error'
    Throw
}



# Create a new CSV file path on the Desktop
$csvFilePath = Join-Path -Path "$($PSScriptRoot)\logs" -ChildPath 'TCMOperatorsGroup.csv'


Write-Log 'Inizio operazione...'

Foreach ($row in $csv) {
    Try {

        Write-Log "Sito: $($row.Site)"
        Write-Host
        Connect-PnPOnline -Url $row.Site -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop -ValidateConnection
        $TCMOpGroups = Get-PnPListItem -List 'FTAEvolutionSettings' | Select-Object -ExpandProperty FieldValues | ForEach-Object {
            $_['TCMOperatorsGroup']
        } | Select-Object -Unique

        $csvRows = @()
        $csvRows += $row.Site

        foreach ($Group in $TCMOpGroups) {
            $csvRows += $Group
        }

        $csvContent = $csvRows -join ';'
        $csvContent | Out-File -FilePath $csvFilePath -Append

        Write-Host 'CSV file updated with group names.'
        Write-Log "Groups have been added for the site $($row.Site)"
    }

    Catch {

        if ($_.InvocationInfo.InvocationName -eq 'Connect-PnPOnline') {
            Write-Log '[ERROR] Errore durante connessione al sito'
            Continue
        }

        Write-Log '[ERROR] Script in errore'
        Throw
    }

    Write-Host '[SUCCESS] CSV creato con successo'
}
