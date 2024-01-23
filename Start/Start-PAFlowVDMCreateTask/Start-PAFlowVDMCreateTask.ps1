<#
    Questo script serve per creare un planner task per uno o più documenti.
    Per funzionare necessita di un file CSV con i seguenti campi:
        - TCM_DN: TCM Document Number
        - Index: Index (Non 'Revision Number')
#>
param (
    [Parameter(Mandatory = $true)][String]$SiteUrl # URL del Sito
    #[Parameter(Mandatory = $true)][String]$MRCode # MR Code
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

function Search-Key {
    param(
        [Parameter(Mandatory = $true)][Array]$Config,
        [Parameter(Mandatory = $true)][String]$Key
    )

    $found = ($Config | Where-Object -FilterScript { $_.Key -eq $Key }).Value

    if ($null -ne $found) { Write-Log "'$($Key)' trovata." }
    else {
        Write-Host "'$($Key)' non trovata." -ForegroundColor Red
        Exit
    }
    return $found
}

function Invoke-PA {
    param(
        [Parameter(Mandatory = $true)][String]$Uri,
        [Parameter(Mandatory = $true)][String]$Body
    )
    $method = 'POST'

    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('Content-Type', 'application/json; charset=utf-8')
    $headers.Add('Accept', 'application/json')
    $encodedBody = [System.Text.Encoding]::UTF8.GetBytes($Body)

    Invoke-RestMethod -Uri $uri -Method $method -Headers $headers -Body $encodedBody | Out-Null
}
Try {
    $CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv $CSVPath -Delimiter ';'
        $RequiredColumns = 'TCM_DN', 'Index'
        $CSVColumns = ($csv | Get-Member -MemberType NoteProperty -Name $RequiredColumns).Name
        $CSV_HasRequiredColumns = Compare-Object -ReferenceObject $RequiredColumns -DifferenceObject $CSVColumns
        if ($CSV_HasRequiredColumns.Count -ne 0) { Throw "[ERROR] - CSV: $($CSVPath) - NOT ALL REQUIRED COLUMNS FOUND (TCM_DN, Index)" }
    }
    else {
        $index = Read-Host -Prompt 'Index'
        $csv = [PSCustomObject] @{
            TCM_DN = $CSVPath
            Index  = $index
            Count  = 1
        }
    }

    $VDL = 'Vendor Documents List'
    $PFS = 'Process Flow Status List'
    $configList = 'Configuration List'
    $keyName = 'FlowUrl_TaskCreation'
    $keyName2 = 'Project_TeamsID'
    $keyName3 = 'FlowURL_DisciplineNotifications'

    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    # Cerca l'URI del flusso della creazioen sulla Configuration List
    Write-Log "Lettura '$($configList)'..."
    $config = Get-PnPListItem -List $configList -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID    = $_['ID']
            Key   = $_['Title']
            Value = $_['VD_ConfigValue']
        }
    }
    $uri = Search-Key -Config $config -Key $keyName
    $teamsId = Search-Key -Config $config -Key $keyName2
    $disciplineNotify = Search-Key -Config $config -Key $keyName3

    # Caricacamento della VDL
    Write-Log "Caricamento '$($VDL)'..."
    $VDLItems = Get-PnPListItem -List $VDL -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID                 = $_['ID']
            TCM_DN             = $_['VD_DocumentNumber']
            Rev                = $_['VD_RevisionNumber']
            Index              = $_['VD_Index']
            DocTitle           = $_['VD_EnglishDocumentTitle']
            PONumber           = $_['VD_PONumber']
            MRCode             = $_['VD_MRCode']
            DisciplineOwnerTCM = $_['VD_DisciplineOwnerTCM'].LookupValue
            DisciplinesTCM     = [Array]$_['VD_DisciplinesTCM'].LookupValue -join '","'
            VendorName         = $_['VD_VendorName'].LookupValue
            VendorSiteUrl      = $_[$($_.FieldValues.Keys.Where({ $_ -eq 'VendorName_x003a_Site_x0020_Url' }))].LookupValue
            Path               = $_['VD_DocumentPath']
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Caricacamento della Process Flow
    Write-Log "Caricamento '$($PFS)'..."
    $PFSItems = Get-PnPListItem -List $PFS -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID              = $_['ID']
            VDL_ID          = $_['VD_VDL_ID']
            Status          = $_['VD_DocumentStatus']
            CommentsEndDate = $_['VD_CommentsEndDate']
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Filtro documenti
    [Array]$items = @()
    ForEach ($row in $csv) {
        $items += $VDLItems | Where-Object -FilterScript { $_.TCM_DN -in $row.TCM_DN -and $_.Index -in $row.Index }
    }

    if ($null -eq $items) { Write-Log "[ERROR] - List: $($VDL) - NOT DOCUMENTS FOUND" }
    else {
        $rowCounter = 0
        Write-Log 'Inizio operazione...'
        ForEach ($item in $items) {
            if ($items.Count -gt 1) { Write-Progress -Activity 'Creazione' -Status "$($rowCounter+1)/$($items.Count) - $($item.TCM_DN)/$($item.Rev)" -PercentComplete (($rowCounter++ / $items.Count) * 100) }

            # Filtro item sulla Process Flow Status
            $PFSItem = $PFSItems | Where-Object -FilterScript { $_.VDL_ID -eq $item.ID }

            if ($null -eq $PFSItem) { $msg = "[ERROR] - List: $($PFS) - Document: $($item.TCM_DN)/$($item.Rev) - NOT FOUND" }
            elseif ($PFSItem.Status -eq 'Commenting') {

                $pathSplit = $item.Path.Split('/')
                $disciplinesArray = '"' + $item.DisciplineOwnerTCM + '"'
                if ($item.DisciplinesTCM -ne '') { $disciplinesArray += ',"' + $item.DisciplinesTCM + '"' }

                # Genera il body da inviare al flusso
                $body = '{
                "ChosenDisciplines": [' + $disciplinesArray + '],
                "RootSiteUrl": "' + $SiteUrl + '",
                "ProjectTeamsID": "' + $teamsId + '",
                "DocumentNumber": "' + $pathSplit[-1] + '",
                "MRCode": "' + $item.MRCode + '",
                "PONumber": "' + $item.PONumber + '",
                "VendorSiteUrl": "' + $item.VendorSiteUrl + '",
                "VendorName": "' + $item.VendorName + '",
                "VDL_DocumentNumber": "' + $item.TCM_DN + '",
                "VDL_RevisionNumber": "' + $item.Rev + '",
                "EnglishDocumentTitle": "' + $item.DocTitle + '",
                "CommentsEndDate": "' + $PFSItem.CommentsEndDate + '",
                "VDL_Index": ' + $item.Index + ',
                "VDL_ID": ' + $item.ID + ',
                "FlowURL_DisciplineNotifications": "' + $disciplineNotify + '"
            }'

                # Chiama il flusso
                try {
                    Invoke-PA -Uri $uri -Body $body
                    $msg = "[SUCCESS] - List: $($VDL) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - CREATION STARTED"
                    Start-Sleep -Seconds 1
                }
                catch {
                    $msg = "[ERROR] - List: $($VDL) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - FAILED - $($_)"
                    throw
                }
            }
            else { continue }
            Write-Log -Message $msg
        }
        Write-Progress -Activity 'Creazione' -Completed
        Write-Log 'Operazione completata.'
    }
}
Catch { Throw }