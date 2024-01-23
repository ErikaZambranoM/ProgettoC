# Il CSV deve contenere le colonne TCM_DN e Rev

param(
    [parameter(Mandatory = $true)][string]$siteCode #URL del sito
)

#Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $siteCode
    )

    $Path = "$($PSScriptRoot)\logs\$($Code)-$(Get-Date -Format 'yyyy_MM_dd').csv";

    if (!(Test-Path -Path $Path)) {
        $newLog = New-Item $Path -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;TCM_DN;Rev;Action;Value'
    }

    if ($Message.Contains('[SUCCESS]')) {
        Write-Host $Message -ForegroundColor Green
    }
    elseif ($Message.Contains('[ERROR]')) {
        Write-Host $Message -ForegroundColor Red
    }
    elseif ($Message.Contains('[WARNING]')) {
        Write-Host $Message -ForegroundColor Yellow
    }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }

    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Message = $Message.Replace(' - List: ', ';').Replace(' - TCM_DN: ', ';').Replace(' - Rev: ', ';').Replace(' - ID: ', ';').Replace(' - Path: ', ';').Replace(' - ', ';').Replace(': ', ';').Replace("'", ';')
    Add-Content $Path "$FormattedDate;$Message"
}

$SiteUrl = "https://tecnimont.sharepoint.com/sites/$($siteCode)DigitalDocuments"
$listName = 'DocumentList'

try {
    $CSVPath = (Read-Host -Prompt "ID / CSV Path / TCM Document Number").Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv -Path $CSVPath -Delimiter ','
        #Validazione Colonne
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
    elseif ($CSVPath -match '^[\d]+$') {
        $csv = [PSCustomObject]@{
            ID    = $CSVPath
            Count = 1
        }
    }
    elseif ($CSVPath -ne '') {
        $rev = Read-Host -Prompt 'Issue Index'
        $csv = [PSCustomObject]@{
            TCM_DN = $CSVPath
            Rev    = $rev
            Count  = 1
        }
    }
    else { Exit }
}
catch {
    throw
}

#Connessione al sito
$Conn = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop
Write-Host "Caricamento '$($listName)'..." -ForegroundColor Cyan
$itemList = Get-PnPListItem -List $listName -PageSize 5000 -Connection $Conn | ForEach-Object {
    [PSCustomObject]@{
        ID              = $_['ID']
        TCM_DN          = $_['Title']
        Rev             = $_['IssueIndex']
        Status          = $_['DocumentStatus']
        IsCurrent       = $_['IsCurrent']
        LastTransmittal = $_['LastTransmittal']
        DocumentsPath   = $_['DocumentsPath']
    }
}
Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

$rowCounter = 0
Write-Host 'Inizio pulizia...' -ForegroundColor Cyan
ForEach ($row in $csv) {
    if ($csv.Count -gt 1) {
        Write-Progress -Activity 'Pulizia' -Status "$($rowCounter+1/$csv.Length) - $($row.TCM_DN)" -PercentComplete (($rowCounter++ / $csv.Count) * 100)
    }

    $delItem = $itemList | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }
    $PathSplit = $delItem.DocumentsPath.Split('/')
    $relativePath = $PathSplit[5..7] -join '/'
    try {
        Remove-PnPListItem -List $listName -Identity $delItem.ID -Recycle -Force | Out-Null
        $msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($delItem.TCM_DN) - Rev: $($delItem.Rev) - DELETED - ID: '$($delItem.ID)'"
    }
    catch { $msg = "[ERROR] - List: $($listName) - TCM_DN: $($delItem.TCM_DN) - Rev: $($delItem.Rev) - FAILED - ID: '$($delItem.ID)'" }
    try {    
        Remove-PnPFolder -Name "$($delItem.TCM_DN)" -Folder "$($relativePath)" -Recycle -Force -Connection $Conn | Out-Null
        $msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($delItem.TCM_DN) - Rev: $($delItem.Rev) - DELETED - ID: '$($delItem.ID)'"
    }
    catch { $msg = "[ERROR] - List: $($listName) - TCM_DN: $($delItem.TCM_DN) - Rev: $($delItem.Rev) - FAILED - ID: '$($delItem.ID)'" }
    Write-Log -Message $msg 
}
if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed }