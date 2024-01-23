#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }
#Questro è un file temporaneo che va a verificare se l'elenco rivecuto in input ha valorizzato il campo "Client Document Class" a IFI

#cosa mi serve: sito, lista, valore da checcare

param(
    [Parameter(Mandatory = $true)][String]$SiteUrl,
    [Parameter(Mandatory = $true)][String]$ListName,
    [Parameter(Mandatory = $true)][String]$FieldInternalName,
    [Parameter(Mandatory = $true)][String]$ValoreAttCheck
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
    #connessione al sito
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    #verifica l'esistenza dell'attributo
    try {
        $internalName = (Get-PnPField -List $ListName -Identity $FieldInternalName).InternalName
        if ($internalName -notin ($csv | Get-Member -MemberType NoteProperty).Name) {
            $csv | Add-Member -NotePropertyName $internalName -NotePropertyValue $null
        }
    }
    catch {
        throw
    }
}
catch {
    throw
}
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
#Scarica la lista del sito
$listItem = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID                 = $_['ID']
        TCM_DN             = $_['Title']
        Revision           = $_['VD_RevisionNumber']
        ClientCode         = $_['VD_ClientDocumentNumber']
        $FieldInternalName = $_[$($FieldInternalName)]
    }
}

foreach ($row in $csv) {
    $filter = $listItem | Where-Object -FilterScript { $_.ClientCode -eq $row.TCM_DN -and $_.Revision -eq $row.Rev }
    if ($filter.$FieldInternalName -eq $ValoreAttCheck) {
        Write-Log "VALIDATE : IL DOC $($filter.TCM_DN) - REV $($filter.Revision) - $($FieldInternalName) = $($filter.$FieldInternalName)"
    }
    else {        
        Write-Log "ERROR : IL DOC $($filter.TCM_DN) - REV $($filter.Revision) - $($FieldInternalName) = $($filter.$FieldInternalName)"
    }
    #Settare l'attributo con il valore desiderato
    #Set-PnPListItem -List $ListName -Identity $filter.ID -Values @{$FieldInternalName = "IFI"}
}