<#
    Questo script serve per ricreare la libreria del PO malformata.
    Check current running flow to avoid trottling

    TODO:
    - Check running executions to adjust the waiting time in the loop.
#>
#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param (
    [Parameter(Mandatory = $true)][String]$ProjectCode, # URL del Sito
    [Switch]$BypassCheck
)

# Funzione di log to CSV
function Write-Log
{
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message -ForegroundColor Cyan }
}

try
{
    # Caricamento ID / CSV / Documento
    $PONumber = (Read-Host -Prompt 'PO Number, CSV Path or TCM Document Number').Trim('"')

    # "^\d+(-1)?$" checks if the string contains only numbers and dots
    if ($PONumber.ToLower().Contains('.csv') )
    {
        If (Test-Path -Path $PONumber) { $csv = Import-Csv -Path $PONumber -Delimiter ';' }
        Else
        {
            Write-Host "File '$($PONumber)' non trovato." -ForegroundColor Red
            Exit
        }
        # Validazione colonne
        $validCols = @('TCM_DN', 'Index')
        $validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
            if ($_ -in $validCols) { $validCounter++ }
        }
        if ($validCounter -lt $validCols.Count)
        {
            Write-Host "Colonne obbligatorie mancanti: $($validCols -join ', ')" -ForegroundColor Red
            Exit
        }
    }
    elseif ($PONumber -eq '')
    {
        Write-Host 'MODE: ALL LIST' -ForegroundColor Red
        Pause
    }
    elseif ($PONumber -notmatch '^\d+(-1)?$')
    {
        $index = Read-Host -Prompt 'Index'
        $doc = [PSCustomObject] @{
            TCM_DN = $PONumber
            Index  = $index
        }
    }

    $siteUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
    $listName = 'Vendor Documents List'
    $configList = 'Configuration List'
    $keyName = 'FlowUrl_DocumentStructure'
    $waitingTime = 15

    # Costanti del Flusso PA
    $method = 'POST'

    # Definizione Header
    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('Accept', 'application/json')

    Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

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
    else
    {
        Write-Host "'$($keyName)' non trovata." -ForegroundColor Red
        Exit
    }

    # Definizione Header in base all'uso Power Automate o Azure Function
    if ($uri.ToLower().Contains('azurewebsites.net')) { $headers.Add('Content-Type', 'text/plain; charset=utf-8') }
    else { $headers.Add('Content-Type', 'application/json; charset=utf-8') }

    # Caricacamento della VDL
    Write-Log "Caricamento '$($listName)'..."
    $VDLItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID            = $_['ID']
            TCM_DN        = $_['VD_DocumentNumber']
            Rev           = $_['VD_RevisionNumber']
            Index         = $_['VD_Index']
            PO            = $_['VD_PONumber']
            VendorWhoSent = $_['VD_VendorWhoSent']
            DocPath       = $_['VD_DocumentPath']
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Filtro documenti con il PO di riferimento # "^\d+(-1)?$" checks if the string contains only numbers and dots
    $items = @()

    if ($PONumber.ToLower().EndsWith('.csv'))
    {
        # Loop through the CSV records and check their presence in the SPO list
        foreach ($row in $csv)
        {
            $found = $false
            foreach ($item in $VDLItems)
            {
                if ($item.TCM_DN -eq $row.TCM_DN -and $item.Index -eq $row.Index)
                {
                    $items += $item
                    $found = $true
                    break
                }
            }
            if (-not $found) { Write-Host "No matching record found for $($row.TCM_DN), $($row.Index)" }
        }
    }
    elseif ($PONumber -eq '') { $items = $VDLItems | Where-Object -FilterScript { $_.DocPath -eq $null } }
    elseif ($PONumber -match '^\d+(-1)?$') { $items = $VDLItems | Where-Object -FilterScript { $_.PO -eq $PONumber } }
    else { $items = $VDLItems | Where-Object -FilterScript { $_.TCM_DN -eq $doc.TCM_DN -and $_.Index -eq $doc.Index } }

    if ($null -eq $items) { Write-Log "[ERROR] - List: $($listName) - PO: $($PONumber) - NOT FOUND" }
    else
    {
        $itemCounter = 0
        Write-Log 'Inizio creazione...'
        ForEach ($item in $items)
        {
            if ($items.Length -gt 1) { Write-Progress -Activity 'Creazione' -Status "$($itemCounter+1)/$($items.Length) - $($item.TCM_DN)" -PercentComplete (($itemCounter++ / $items.Length) * 100) }
            if ($null -eq $item.VendorWhoSent -or $BypassCheck)
            {
                # Genera il body da inviare al flusso
                $body = '{"siteUrl": "' + $siteUrl + '","vdlItemId": ' + $item.ID + ',"isManual": 1}'

                # Chiama il flusso / Azure Function
                try
                {
                    $encodedBody = [System.Text.Encoding]::UTF8.GetBytes($Body)

                    # Local uri when using local Azure Function
                    #$uri = 'http://localhost:7116/api/CreateVendorFolderStructure'

                    Invoke-RestMethod -Uri $uri -Method $method -Headers $headers -Body $encodedBody | Out-Null
                    Write-Log "[SUCCESS] - List: $($listName) - Doc: $($item.TCM_DN) - Index: $($item.Index) - CREATION STARTED"
                    Write-Log '[WARNING] - Check correct permissions and status on the document.'
                }
                catch { Write-Log "[ERROR] - List: $($listName) - Doc: $($item.TCM_DN) - Index: $($item.Index) - FAILED - $($_)" }

                # Sleep tra una chiamata e l'altra
                if ($items.Length -gt 1 -and !($uri.ToLower().Contains('azurewebsites.net')))
                {
                    Write-Log "Waiting $($waitingTime) seconds..."
                    Start-Sleep -Seconds $waitingTime
                }
            }
            else { Write-Log "[WARNING] - List: $($listName) - Doc: $($item.TCM_DN) - Index: $($item.Index) - DOC NOT IN PLACEHOLDER - SKIPPED" }
        }
        Write-Log 'Operazione completata.'
    }
}
catch { Throw }
finally { if ($items.Length -gt 1) { Write-Progress -Activity 'Creazione' -Completed } }