<#
    Questo script permetti di copiare i valori da una colonna all'altra usando il PnP Batch.
    Per funzionare, è necessario avere i permessi per l'uso di PnP con login Interactive.
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
    [Parameter(Mandatory = $true)][string]$SiteUrl,
    [Parameter(Mandatory = $true)][string]$ListName,
    [Parameter(Mandatory = $true)][string]$SourceField,
    [Parameter(Mandatory = $true)][string]$DestinationField,
    [Int]$BatchSize = 1000
)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $siteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog "Timestamp;Type;ListName;TCM_DN;Rev;Action;Value"
    }
    $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($Message.Contains("[SUCCESS]")) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains("[ERROR]")) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains("[WARNING]")) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(" - List: ", ";").Replace(" - TCM_DN: ", ";").Replace(" - Rev: ", ";").Replace(" - DeptCode: ", ";").Replace(" - Desc: ", ";").Replace(" - ", ";").Replace(": ", ";")
    Add-Content $logPath "$FormattedDate;$Message"
}

# Funzione che scrive il batch su SharePoint
function Update-ListUsingBatch {
    param (
        [Parameter(Mandatory=$true)][System.Object]$Batch
    )

    $amount = $Batch.RequestCount
    Write-Log "Batch $($amount) record pronto. Avvio scrittura..."
    try {
        Invoke-PnPBatch -Batch $Batch
        Write-Host "[SUCCESS] - List: $($ListName) - Number of record: $($amount) - PATCHED" -ForegroundColor Green
    }
    catch { Write-Log "[ERROR] - List: $($ListName) - Number of record: $($amount) - $($_)" }
}

try {
    # Connessione interactive
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    $siteCode = (Get-PnPWeb).Title.Split(' ')[0]

    # Caricamento lista
    Write-Host "Caricamento '$($ListName)'..." -ForegroundColor Cyan
    $listItems = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID        = $_["ID"]
            SrcField  = $_[$SourceField]
            DestField = $_[$DestinationField]
        }
    }
    Write-Log "Caricamento lista completato."

    # Filtro per gli elementi non vuoti e con valori diversi
    [Array]$filtered = $listItems | Where-Object -FilterScript { $_.SrcField -ne $null -and $_.SrcField -ne $_.DestField }

    # Creazione batch per l'aggiornamento
    Write-Log "Inizio operazione..."
    $recordBlock = New-PnPBatch
    $counter = 0
    $batchCounter = 1
    $itemMax = $filtered.Count -lt $BatchSize ? ($filtered.Count) : ($BatchSize)
    for ($i = 0; $i -lt $filtered.Count; $i++) {
        if ($filtered.Count -gt 1) { Write-Progress -Activity "Batch" -Status "Block: $($batchCounter) - Current: $($counter+1)/$($itemMax) - Done: $($BatchSize * $batchCounter)/$($filtered.Count)" -PercentComplete (($counter++ / $itemMax) * 100) }

        # Aggiunta record al batch
        try {
            Set-PnPListItem -List $ListName -Identity $filtered[$i].ID -Values @{ $DestinationField = $filtered[$i].SrcField } -UpdateType SystemUpdate -Batch $recordBlock
            Write-Log "[SUCCESS] - List: $($ListName) - ID: $($filtered[$i].ID) - $($SourceField): $($filtered[$i].SrcField) - $($DestinationField): $($filtered[$i].DestField) - UPDATED"
        }
        catch { Write-Log "[ERROR] - List: $($ListName) - ID: $($filtered[$i].ID) - $($SourceField): $($filtered[$i].SrcField) - $($DestinationField): $($filtered[$i].DestField) - $($_)" }

        # Scrittura batch ogni $BatchSize record
        if ($counter -eq $BatchSize) {
            Update-ListUsingBatch -Batch $recordBlock 
            $counter = 0
            $itemMax = (($filtered.Count - ($BatchSize * $batchCounter)) -lt $BatchSize) ? ($filtered.Count - ($BatchSize * $batchCounter)) : ($BatchSize)
            $batchCounter++
            $recordBlock = New-PnPBatch
            Write-Log "Generazione nuovo batch..."
        }
    }
    
    Update-ListUsingBatch -Batch $recordBlock

    Write-Log "Operazione completata."
}
catch { Throw }
finally { if ($filtered.Count -gt 1) { Write-Progress -Activity "Batch" -Completed } }