# Questo script serve per ricreare la libreria del PO malformata.
param (
    [Parameter(Mandatory = $true)][String]$SiteUrl # URL del Sito
)

# Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message -ForegroundColor Cyan }
}

# Caricamento ID / CSV / Documento
$docNumber = Read-Host -Prompt 'TCM Document Number'
$index = Read-Host -Prompt 'Index'
$doc = [PSCustomObject]@{
    TCM_DN = $docNumber
    Index  = $index
}

$listName = 'Process Flow Status List'
$configList = 'Configuration List'
$keyName = 'CloseTasksFlowUrl'

# Costanti del Flusso PA
$method = 'POST'

$headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
$headers.Add('Content-Type', 'application/json; charset=utf-8')
$headers.Add('Accept', 'application/json')

Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop

# Cerca l'URI del flusso della creazioen sulla Configuration List
Write-Log "Lettura '$($configList)'..."
$config = Get-PnPListItem -List $configList -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID    = $_['ID']
        Key   = $_['Title']
        Value = $_['VD_ConfigValue']
    }
}
$uri = ($config | Where-Object -FilterScript { $_.Key -eq $keyName }).Value
if ($null -ne $uri) { Write-Log "'$($keyName)' trovata." }
else {
    Write-Host "'$($keyName)' non trovata." -ForegroundColor Red
    Exit
}

# Caricacamento della VDL
Write-Log "Caricamento '$($listName)'..."
$VDLItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID     = $_['ID']
        TCM_DN = $_['VD_DocumentNumber']
        Rev    = $_['VD_RevisionNumber']
        Index  = $_['VD_Index']
    }
}
Write-Log 'Caricamento lista completato.'

# Filtro documenti con il PO di riferimento
$items = $VDLItems | Where-Object -FilterScript { $_.TCM_DN -eq $doc.TCM_DN -and $_.Index -eq $doc.Index }

if ($null -eq $items) { Write-Log "[ERROR] - List: $($ListName) - Document: $($doc.TCM_DN)/$($doc.Rev) - NOT FOUND" }
else {
    $rowCounter = 0
    Write-Log 'Inizio pulizia...'
    ForEach ($item in $items) {
        if ($items.Length -gt 1) { Write-Progress -Activity 'Pulizia' -Status "$($item.TCM_DN)-$($item.Rev)" -PercentComplete (($rowCounter++ / $items.Length) * 100) }
        # Genera il body da inviare al flusso
        $body = '{
            "SiteUrl": "' + $SiteUrl + '",
            "ProcessID": ' + $item.ID + '
        }'

        # Chiama il flusso
        try {
            $encodedBody = [System.Text.Encoding]::UTF8.GetBytes($Body)
            Invoke-RestMethod -Uri $uri -Method $method -Headers $headers -Body $encodedBody | Out-Null
            $msg = "[SUCCESS] - List: $($listName) - Document: $($item.TCM_DN)/$($item.Rev) - CLEANING TASKS"
        }
        catch { $msg = "[ERROR] - List: $($listName) - Document: $($item.TCM_DN)/$($item.Rev) - FAILED - $($_)" }
        Write-Log -Message $msg
        Start-Sleep -Seconds 1
    }
    Write-Progress -Activity 'Pulizia' -Completed
    Write-Log 'Pulizia completata.'
}