<#Crea la cartella Native quando non si riesce a creare#>
param (
    [Parameter(Mandatory = $true)][String]$ProjectCode
)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $ProjectCode
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
    $Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Previous: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

try {
    # Caricamento CSV/Documento/Tutta la lista
    $CSVPath = (Read-Host -Prompt 'CSV Path o Last Transmittal').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
    elseif ($CSVPath -ne '') {
        $csv = [PSCustomObject]@{
            LastTrn = $CSVPath
            Count   = 1
        }
    }
    else {
        Write-Host 'MODE: ALL LIST' -ForegroundColor Red
        Pause
    }

    $TCMUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
    $ClientUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
    $VDMUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
    $workFolder = "C:\Temp\$($ProjectCode)"

    if (!(Test-Path -Path $workFolder)) { New-Item -Path $workFolder -ItemType Directory -Force | Out-Null }

    $DL = 'DocumentList'
    $CDL = 'Client Document List'
    $VDL = 'Vendor Documents List'
    $NDList = 'NativeDocuments'

    $ClientConn = Connect-PnPOnline -Url $ClientUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $TCMConn = Connect-PnPOnline -Url $TCMUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $VDMConn = Connect-PnPOnline -Url $VDMUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

    Write-Log "Caricamento '$($CDL)'..."
    $CDLItems = Get-PnPListItem -List $CDL -PageSize 5000 -Connection $ClientConn | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            Rev        = $_['IssueIndex']
            ClientCode = $_['ClientCode']
            LastTrn    = $_['LastTransmittal']
            IDSrcList  = $_['IDDocumentList']
            SrcEnv     = $_['DD_SourceEnvironment']
            DocPath    = $_['DocumentsPath']
        }
    } | Sort-Object LastTrn -Descending
    Write-Log 'Caricamento lista completato.'

    Write-Log "Caricamento '$($DL)'..."
    $DLItems = Get-PnPListItem -List $DL -PageSize 5000 -Connection $TCMConn | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            Rev        = $_['IssueIndex']
            ClientCode = $_['ClientCode']
            LastTrn    = $_['LastTransmittal']
            DocPath    = $_['DocumentsPath']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Write-Log "Caricamento '$($VDL)'..."
    $VDLItems = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $VDMConn | ForEach-Object {
        [PSCustomObject]@{
            ID      = $_['ID']
            TCM_DN  = $_['VD_DocumentNumber']
            Rev     = $_['VD_RevisionNumber']
            LastTrn = $_['LastTransmittal']
            DocPath = $_['VD_DocumentPath']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Write-Log "Caricamento '$($NDList)'..."
    $NDItems = Get-PnPListItem -List $NDList -PageSize 5000 -Connection $ClientConn | ForEach-Object {
        [PSCustomObject]@{
            ID      = $_['ID']
            Name    = $_['FileLeafRef']
            RelPath = $_['FileRef']
        }
    }
    Write-Log 'Caricamento lista completato.'

    if ($CSVPath -eq '') { $csv = $CDLItems }

    $itemCounter = 0
    Write-Log 'Inizio controllo...'
    ForEach ($item in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($itemCounter+1)/$($csv.Length) - $($item.LastTrn)" -PercentComplete (($itemCounter++ / $csv.Length) * 100) }
        
        $folder = $NDItems | Where-Object -FilterScript { $_.Name -eq $item.LastTrn }

        #if ($null -eq $folder) {
        if ($true) {
            try {
                Add-PnPFolder -Folder $NDList -Name $item.LastTrn -Connection $ClientConn | Out-Null
                Write-Log "[WARNING] - List: $($NDList) - Name: $($item.LastTrn) - CREATED"
            }
            catch {}

            [Array]$documents = $CDLItems | Where-Object -FilterScript { $_.LastTrn -eq $item.LastTrn }
            $nativeFolder = "$($NDList)/$($item.LastTrn)"

            if ($null -eq $documents) { Write-Log "[ERROR] - List: $($CDL) - Last Transmittal: $($item.LastTrn) - NOT FOUND" }
            else {
                foreach ($document in $documents) {

                    if ($document.SrcEnv -eq 'DigitalDocuments') {
                        $DLItem = $DLItems | Where-Object -FilterScript { $_.ID -eq $document.IDSrcList }

                        if ($null -ne $DLItem) {
                            $folderRelPath = $DLItem.DocPath.Replace($TCMUrl + '/', '') + '/' + $DLItem.TCM_DN
                            $nativeDocs = Get-PnPFolderItem -FolderSiteRelativeUrl $folderRelPath -ItemType File -Connection $TCMConn
                            $mainConn = $TCMConn
                        }
                    }
                    else {
                        $VDLItem = $VDLItems | Where-Object -FilterScript { $_.ID -eq $document.IDSrcList }

                        if ($null -ne $VDLItem) {
                            $pathSplit = $VDLItem.DocPath.Split('/')
                            $subSite = $pathSplit[0..5] -join '/'
                            $subConn = Connect-PnPOnline -Url $subSite -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
                            $nativeRelPath = $VDLItem.DocPath.Replace($subSite + '/', '') + '/Native'
                            $OFVRelPath = $VDLItem.DocPath.Replace($subSite + '/', '') + '/OFV'
                            $nativeDocs = Get-PnPFolderItem -FolderSiteRelativeUrl $nativeRelPath -ItemType File -Connection $subConn
                            $ofvDocs = Get-PnPFolderItem -FolderSiteRelativeUrl $OFVRelPath -ItemType File -Connection $subConn
                            ForEach ($file in $ofvDocs) {
                                if (!($file.Name.ToLower().Contains('.pdf'))) {
                                    try {
                                        Get-PnPFile -Url $($file.ServerRelativeUrl) -AsFile -Path "$($workFolder)" -Filename $file.Name -Connection $subConn | Out-Null
                                        Add-PnPFile -Path "$($workFolder)\$($file.Name)" -Folder $nativeFolder -Connection $ClientConn | Out-Null
                                        Remove-Item -Path "$($workFolder)\$($file.Name)" -Force | Out-Null
                                        $msg = "[SUCCESS] - List: $($NDList) - File: OFV/$($file.Name) - UPLOADED"
                                    }
                                    catch { $msg = "[ERROR] - List: $($NDList) - File: OFV/$($file.Name) - UPLOAD FAILED - $($_)" }
                                    Write-Log $msg
                                }
                            }
                            $mainConn = $VDMConn
                        }
                    }
                    ForEach ($file in $nativeDocs) {
                        try {
                            Get-PnPFile -Url $($file.ServerRelativeUrl) -AsFile -Path "$($workFolder)" -Filename $file.Name -Connection $mainConn | Out-Null
                            Add-PnPFile -Path "$($workFolder)\$($file.Name)" -Folder $nativeFolder -Connection $ClientConn | Out-Null
                            Remove-Item -Path "$($workFolder)\$($file.Name)" -Force | Out-Null
                            $msg = "[SUCCESS] - List: $($NDList) - File: NATIVE/$($file.Name) - UPLOADED"
                        }
                        catch { $msg = "[ERROR] - List: $($NDList) - File: NATIVE/$($file.Name) - UPLOAD FAILED - $($_)" }
                        Write-Log $msg
                    }
                }
            }
        }
        else { Write-Log "[SUCCESS] - List: $($NDList) - Name: $($item.LastTrn) - SKIPPED" }
    }
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed }
    Remove-Item -Path $workFolder -Force | Out-Null
    Write-Log 'Operazione completata.'
}
Catch { Throw }