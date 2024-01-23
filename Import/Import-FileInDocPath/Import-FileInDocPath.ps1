#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
    [parameter(Mandatory = $true)][String]$SiteUrl # URL sito
)

# Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $siteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID/Doc;Action;Key;Value;OldValue'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Doc: ', ';').Replace(' - Folder: ', ';').Replace(' - File: ', ';').Replace('/', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

try {
    if ($SiteUrl.ToLower().EndsWith('digitaldocumentsc')) { $listName = 'Client Document List' }
    elseif ($SiteUrl.ToLower().EndsWith('digitaldocuments')) { $ListName = 'DocumentList' }
    elseif ($SiteUrl.ToLower().Contains('vdm_')) { $ListName = 'Vendor Documents List' }

    # Caricamento CSV / Documento
    $CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv -Path $CSVPath -Delimiter ';'
        # Validazione colonne
        $validCols = @('TCM_DN', 'Rev')
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
        $rev = Read-Host -Prompt 'Issue Index'
        $csv = [PSCustomObject] @{
            TCM_DN = $CSVPath
            Rev    = $rev
            Count  = 1
        }
    }

    $localPath = (Read-Host -Prompt 'Local file path').Trim('"')

    # Connessione al sito
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    $siteCode = (Get-PnPWeb).Title.Split(' ')[0]

    Write-Log "Caricamento '$($ListName)'..."
    $listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
        if ($SiteUrl.ToLower().Contains('vdm_')) {
            [PSCustomObject]@{
                ID         = $_['ID']
                TCM_DN     = $_['VD_DocumentNumber']
                Rev        = $_['VD_RevisionNumber']
                ClientCode = $_['VD_ClientDocumentNumber']
                DocPath    = $_['VD_DocumentPath']
            }
        }
        else {
            [PSCustomObject]@{
                ID         = $_['ID']
                TCM_DN     = $_['Title']
                Rev        = $_['IssueIndex']
                ClientCode = $_['ClientCode']
                DocPath    = $_['DocumentsPath']
            }
        }
    }
    Write-Log 'Caricamento lista completato.'

    $rowCounter = 0
    Write-Log 'Inizio operazione...'
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)-$($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

        $item = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

        if ($null -eq $item) { Write-Log "[ERROR] - List: $($listName) - Doc: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
        if ([Array]$item.Length -gt 1) { Write-Log "[WARNING] - List: $($listName) - Doc: $($row.TCM_DN)/$($row.Rev) - DUPLICATED" }
        else {
            $pathSplit = $item.DocPath.Split('/')
            $folderRelPath = ($pathSplit[5..($pathSplit.Length)] -join '/')

            try {
                $files = Get-ChildItem -Path $localPath -Recurse -ErrorAction Stop | Where-Object -FilterScript { $_.Name.StartsWith($item.TCM_DN) -and $_.Name.EndsWith("$($item.Rev).pdf") -or $_.Name.StartsWith($item.ClientCode) -and $_.Name.EndsWith("$($item.Rev).pdf") }
                if ($null -eq $files) { Write-Log "[WARNING] - LocalFolder: $($localPath) - Doc: $($item.TCM_DN)/$($item.Rev) - NOT FOUND" }
                else {
                    ForEach ($file in $files) {
                        try {
                            $newFileName = $file.Name -replace $item.ClientCode, $item.TCM_DN
                            Add-PnPFile -Path $file.FullName -Folder $folderRelPath -NewFileName $newFileName | Out-Null
                            Write-Log "[SUCCESS] - List: $($pathSplit[5]) - File: $($newFileName) - UPLOADED"
                        }
                        catch { Write-Log "[ERROR] - List: $($pathSplit[5]) - File: $($newFileName) - FAILED - $($_)" }
                    }
                }
            }
            catch {
                Write-Host "[ERROR] - LocalFolder: $($clientTrnPath) - $($_)" -ForegroundColor Red
                Exit
            }
        }
    }
    Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }