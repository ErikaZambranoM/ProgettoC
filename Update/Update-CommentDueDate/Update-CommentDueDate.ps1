#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

<#
    - CSV must contain the following columns: TCM_DN, Rev, CommentDueDate
    - Date must be provided in format 'MM/dd/yyyy'
    - We recommend to review your CSV on an editor (such as notepad++) before launching the script
#>


param(
    [parameter(Mandatory = $true)][string]$SiteCode, # Codice del sito
    [Switch]$System #System Update (opzionale)
)

# Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $SiteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID;Action;Key;Value;OldValue'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Doc: ', ';').Replace('/', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

try {
    # Funzione SystemUpdate
    $system ? ( $updateType = 'SystemUpdate' ) : ( $updateType = 'Update' ) | Out-Null

    # URL del sito
    $tcmUrl = "https://tecnimont.sharepoint.com/sites/$($SiteCode)DigitalDocuments"
    $clientUrl = "https://tecnimont.sharepoint.com/sites/$($SiteCode)DigitalDocumentsC"
    $VDMUrl = "https://tecnimont.sharepoint.com/sites/VDM_$($SiteCode)"

    # URI Flow
    $flowURI = "https://prod-171.westeurope.logic.azure.com:443/workflows/f33d2976d48541d3898de11facb658c8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=DZ4bT8oNEX4SI1j5a-2bq8gXNQoNBvyKaqQbsubk_fM"

    # Indentifica il nome della Lista
    $CDL = 'Client Document List'
    $RTPList = 'Review Task Panel'
    $RTAList = 'Review Task Archive'
    $columnName = 'CommentDueDate'

    # Definizione Header
    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('Accept', 'application/json')
    $headers.Add('Content-Type', 'application/json; charset=utf-8')

    # Caricamento CSV/Documento/Tutta la lista
    $CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) { 
        $csv = Import-Csv -Path $CSVPath -Delimiter ';'
        # Validazione colonne
        $validCols = @('TCM_DN', 'Rev', $columnName)
        $validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
            if ($_ -in $validCols) { $validCounter++ }
        }
        if ($validCounter -lt $validCols.Count) {
            Write-Host "Mandatory column(s) missing: $($validCols -join ', ')" -ForegroundColor Red
            Exit
        }
    }
    elseif ($CSVPath -ne '') {
        $Rev = Read-Host -Prompt 'Issue Index'
        $newDate = Read-Host -Prompt $columnName
        $csv = [PSCustomObject] @{
            TCM_DN         = $CSVPath
            Rev            = $Rev
            CommentDueDate = $newDate
            Count          = 1
        }
    }

    $create = $Host.UI.PromptForChoice("Create new task", "", @("&Yes", "&No"), 1)

    # Connessione al sito
    $tcmConn = Connect-PnPOnline -Url $tcmUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    try { $VDMConn = Connect-PnPOnline -Url $VDMUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection }
    catch { Write-Host "[WARNING] - Sito VDM non esistente." -ForegroundColor Yellow }
    $clientConn = Connect-PnPOnline -Url $clientUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $listsArray = Get-PnPList -Connection $clientConn

    # Carica la Client Document List
    Write-Log "Caricamento '$($CDL)'..."
    $CDItems = Get-PnPListItem -List $CDL -PageSize 5000 -Connection $clientConn | ForEach-Object {
        [PSCustomObject]@{
            ID                = $_['ID']
            TCM_DN            = $_['Title']
            Rev               = $_['IssueIndex']
            ClientCode        = $_['ClientCode']
            ID_DL             = $_['IDDocumentList']
            Title             = $_['DocumentTitle']
            RFI               = $_['ReasonForIssue']
            ClientDiscipline  = $_['Client_Discipline']
            CommentDueDate    = $_['CommentDueDate']
            MS                = $_['DocumentClass']
            CDC               = $_['OwnerDocumentClass']
            DocPath           = $_['DocumentsPath']
            StagingPath       = $_['StagingDocumentsPath']
            SourceEnvironment = $_['DD_SourceEnvironment']
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Carica Review Task Panel
    Write-Log "Caricamento '$($RTPList)'..."
    $RTPItems = Get-PnPListItem -List $RTPList -PageSize 5000 -Connection $clientConn | ForEach-Object {
        [PSCustomObject]@{
            ID     = $_['ID']
            ID_CDL = $_['IDClientDocumentList']
            Rev    = $_['IssueIndex']
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Carica Review Task Archive
    Write-Log "Caricamento '$($RTAList)'..."
    $RTAItems = Get-PnPListItem -List $RTAList -PageSize 5000 -Connection $clientConn | ForEach-Object {
        [PSCustomObject]@{
            ID     = $_['ID']
            ID_CDL = $_['IDClientDocumentList']
            Rev    = $_['IssueIndex']
        }
    }
    Write-Log 'Caricamento lista completato.'

    $rowCounter = 0
    Write-Log 'Inizio correzione...'
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

        Write-Host "Doc: $($row.TCM_DN)/$($row.Rev)" -ForegroundColor Blue
        $item = $CDItems | Where-Object -FilterScript { ($_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev) -or ($_.ClientCode -eq $row.TCM_DN -and $_.Rev -eq $row.Rev) }
        if ($null -eq $item) { Write-Log "[ERROR] - List: $($CDL) - Doc: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
        elseif ($item.Length -gt 1) { Write-Log "[WARNING] - List: $($CDL) - Doc: $($row.TCM_DN)/$($row.Rev) - DUPLICATED" }
        else {
            # Aggiorno l'item lato TCM
            if ($item.SourceEnvironment -eq 'VendorDocuments') {
                $listName = 'Vendor Documents List'
                $srcConn = $VDMConn
            }
            else {
                $listName = 'DocumentList'
                $srcConn = $tcmConn
            }

            try {
                Set-PnPListItem -List $listName -Identity $item.ID_DL -Values @{
                    $columnName = $row.CommentDueDate
                } -UpdateType $updateType -Connection $srcConn | Out-Null
                Write-Log "[SUCCESS] - List: $($listName) - ID: $($item.ID_DL) - Doc: $($item.TCM_DN)/$($item.Rev) - UPDATE"
            }
            catch { Write-Log "[ERROR] - List: $($listName) - ID: $($item.ID_DL) - Doc: $($item.TCM_DN)/$($item.Rev) - FAILED $($_)" }
            # VERIFICA SE IL DOCUMENTO è APERTO IN STAIGING AREA
            if ($null -ne $item.StagingPath) {
                $pathSTASplit = $item.StagingPath.Split('/')
                $relSTAPath = ($pathSTASplit[5..$pathSTASplit.Length] -join '/')
                $DLRelSTAPath = '/' + ($pathSTASplit[3..5] -join '/')
                $DLSTA = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelSTAPath }
                Write-Host "[WARNING] - Documento in Staging Area" -ForegroundColor Yellow
                # Aggiorna la Revisione su Staging Area
                $fileListSTA = Get-PnPFolderItem -FolderSiteRelativeUrl $relSTAPath -ItemType File -Recursive -Connection $clientConn | Select-Object Name, ServerRelativeUrl
                if ($null -eq $fileListSTA) { Write-Log "[WARNING] - List: $($DLSTA.Title) - Doc: $($item.TCM_DN)/$($item.Rev) - EMPTY OR NOT FOUND" }
                else {
                    foreach ($file in $fileListSTA) {
                        $fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem -Connection $clientConn
                        If ($null -ne $fileAtt.FieldValues.$columnName) {
                            try {
                                Set-PnPListItem -List $DLSTA.Title -Identity $fileAtt.Id -Values @{
                                    $columnName = $row.CommentDueDate
                                } -UpdateType $updateType -Connection $clientConn | Out-Null
                                Write-Log "[SUCCESS] - List: $($DLSTA.Title) - FileName: $($file.Name) - UPDATED"
                            }
                            catch { Write-Log "[ERROR] - List: $($DLSTA.Title) - FileName: $($file.Name) - FAILED" }
                        }
                    }
                }
            }
            # Calcolo variabili per aggiornamento dei file
            $pathSplit = $item.DocPath.Split('/')
            $relPath = ($pathSplit[5..$pathSplit.Length] -join '/')
            $DLRelPath = '/' + ($pathSplit[3..5] -join '/')
            $DL = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }

            # Aggiorno l'item sulla CDL
            try {
                Set-PnPListItem -List $CDL -Identity $item.ID -Values @{
                    $columnName = $row.CommentDueDate
                } -UpdateType $updateType -Connection $clientConn | Out-Null
                Write-Log "[SUCCESS] - List: $($CDL) - ID: $($item.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - UPDATE"
            }
            catch { Write-Log "[ERROR] - List: $($CDL) - ID: $($item.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - FAILED" }
            # Aggiorno la Revisione sui file
            $fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $relPath -ItemType File -Recursive -Connection $clientConn | Select-Object Name, ServerRelativeUrl

            if ($null -eq $fileList) { Write-Log "[WARNING] - List: $($DL.Title) - Doc: $($item.TCM_DN)/$($item.Rev) - EMPTY OR NOT FOUND" }
            else {
                foreach ($file in $fileList) {
                    $fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem -Connection $clientConn
                    If ($null -ne $fileAtt.FieldValues.$columnName) {
                        try {
                            Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
                                $columnName = $row.CommentDueDate
                            } -UpdateType $updateType -Connection $clientConn | Out-Null
                            Write-Log "[SUCCESS] - List: $($DL.Title) - FileName: $($file.Name) - UPDATED"
                        }
                        catch { Write-Log "[ERROR] - List: $($DL.Title) - FileName: $($file.Name) - FAILED" }
                    }
                }
            }

            if ($SiteCode -eq "4355") {
                # Filtra il documento nel Review Task Archive
                $RTARecords = $RTAItems | Where-Object -FilterScript { $_.ID_CDL -eq $item.ID }
                if ($null -ne $RTARecords) {
                    foreach ($task in $RTARecords) {
                        try {
                            Set-PnPListItem -List $RTAList -Identity $task.ID -Values @{
                                $columnName = $row.CommentDueDate
                            } -UpdateType $updateType -Connection $clientConn | Out-Null
                            Write-Log "[SUCCESS] - List: $($RTAList) - ID: $($task.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - UPDATE"
                        }
                        catch { Write-Log "[ERROR] - List: $($RTAList) - ID: $($task.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - FAILED" }
                    }
                }
                # Se non lo trova nell'Archive, filtra il Review Task Panel
                else {
                    $RTPRecords = $RTPItems | Where-Object -FilterScript { $_.ID_CDL -eq $item.ID }
                    if ($null -eq $RTPRecords) { Write-Log "[WARNING] - List: $($RTPList) - Doc: $($item.TCM_DN)/$($item.Rev) - NOT FOUND" }
                    else {
                        foreach ($task in $RTPRecords) {
                            try {
                                Set-PnPListItem -List $RTPList -Identity $task.ID -Values @{
                                    $columnName = $row.CommentDueDate
                                } -UpdateType $updateType -Connection $clientConn | Out-Null
                                Write-Log "[SUCCESS] - List: $($RTPList) - ID: $($task.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - UPDATE"
                            }
                            catch { Write-Log "[ERROR] - List: $($RTPList) - ID: $($task.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - FAILED" }
                        }
                    }
                }

                if ($create -eq 0) {
                    if ($null -ne $RTARecords) {
                        foreach ($task in $RTARecords) {
                            try {
                                Remove-PnPListItem -List $RTAList -Identity $task.ID -Recycle -Force -Connection $clientConn | Out-Null
                                Write-Log "[SUCCESS] - List: $($RTAList) - ID: $($task.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - DELETED"
                            }
                            catch { Write-Log "[ERROR] - List: $($RTAList) - ID: $($task.ID) - Doc: $($item.TCM_DN)/$($item.Rev) - FAILED" }
                        }
                    }

                    if ($null -eq $RTPRecords) {
                        $convDate = Get-Date ([DateTime]::ParseExact($row.CommentDueDate, 'MM/dd/yyyy', $null)) -Format "yyyy-MM-dd"
                        $body = '{
                            "IDClientDocument": "' + $item.ID + '",
                            "IDDocument": "' + $item.ID_DL + '",
                            "LastClientTransmittal": "",
                            "ClientCode": "' + $item.ClientCode + '",
                            "ReasonForIssue": "' + $item.RFI + '",
                            "DocumentTitle": "' + $item.Title + '",
                            "PONumber": "",
                            "ClientDiscipline": "' + $item.ClientDiscipline + '",
                            "CommentDueDate": "' + $convDate + '",
                            "CommentRequest": "True",
                            "IssueIndex": "' + $item.Rev + '",
                            "ClientSiteUrl": "' + $clientUrl + '",
                            "TCMSiteUrl": "' + $tcmUrl + '",
                            "VendorSiteUrl": "' + $VDMUrl + '",
                            "SourceEnvironment": "' + $item.SourceEnvironment + '",
                            "DocumentClass": "' + (($item.SourceEnvironment -eq 'VendorDocuments') ? ($item.CDC) : ($item.MS)) + '"
                        }'

                        try {
                            $encodedBody = [System.Text.Encoding]::UTF8.GetBytes($body)
                            Invoke-RestMethod -Uri $flowURI -Method "POST" -Headers $headers -Body $encodedBody | Out-Null
                            Write-Log "[SUCCESS] - List: $($RTPList) - NEW - Doc: $($item.TCM_DN)/$($item.Rev) - CREATION STARTED"
                        }
                        catch { Write-Log "[ERROR] - List: $($RTPList) - NEW - Doc: $($item.TCM_DN)/$($item.Rev) - $($_)" }
                    }
                }
            }
        }
    }
    Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }