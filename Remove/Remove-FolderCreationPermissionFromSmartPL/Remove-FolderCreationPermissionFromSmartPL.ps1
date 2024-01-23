#Script che disabilita la possibilità di creare una nuova cartella in tutte le DL di tutti sottositi dei Vendor contenenti "SmartPL" nel nome

# Funzione di log to CSV
Function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
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

Write-Log 'Inizio operazione...'

Foreach ($row in $csv) {
    Try {

        Write-Log "Sito: $($row.Site)"
        Write-Host
        Connect-PnPOnline -Url $row.Site -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop -ValidateConnection
        $vendors = Get-PnPListItem -List 'Vendors' | Select-Object -ExpandProperty FieldValues | ForEach-Object {
            $_['VD_SiteUrl']
        }
        foreach ($vendor in $vendors) {


            try {

                Connect-PnPOnline -Url $vendor -UseWebLogin -WarningAction SilentlyContinue -ErrorAction Stop -ValidateConnection
                Write-Log "[SUCCESS] Connessione effettuata al sito $($vendor)"
                Write-Host
                $DLSmartPL = Get-PnPList | Where-Object { $_.Title -like '*SmartPL*' }

                # If there are SmartPL lists, process them
                if ($DLSmartPL.Count -ne 0) {
                    # Iterate through SmartPL lists and disable folder creation
                    Write-Log "Trovate le seguenti DL: $($DLSmartPL.Title)"
                    foreach ($DL in $DLSmartPL) {
                        if ($DL.EnableFolderCreation -eq $True) {
                            Write-Log "Rimozione della FolderCreation dalla DL $($DL.Title)"
                            Set-PnPList -Identity $DL.Title -EnableFolderCreation $false | Out-Null
                            Write-Log '[SUCCESS] FolderCreation Disabilitata '
                            Write-Host
                        }
                        else {
                            Write-Log '[WARNING] FolderCreation già disabilitata '
                            Write-Host
                        }
                    }
                }
                else {
                    Write-Log "[WARNING] No SmartPL lists found for vendor $($vendor)"
                    Write-Host
                }
            }
            catch {
                Write-Log "[ERROR] Errore durante la ricerca o l'elaborazione della lista SmartPL per il vendor $($vendor)"

            }
        }


        Write-Log '[SUCCESS] Script eseguito con successo'
    }
    Catch {

        if ($_.InvocationInfo.InvocationName -eq 'Connect-PnPOnline') {
            Write-Log '[ERROR] Errore durante connessione al sito'
            Continue
        }

        Write-Log '[ERROR] Script in errore'
        Throw
    }
}
