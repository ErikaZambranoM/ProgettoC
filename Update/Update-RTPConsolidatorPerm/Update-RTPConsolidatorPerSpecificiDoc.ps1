<#
    Questo script concede i permessi a tutti gli utenti inseriti del campo Consolidator sul Review Task Panel.

    AGGIORNARE IL FILTRO PRIMA DI ESEGUIRE.
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param (
    [Parameter(Mandatory = $true)][String]$ProjectCode
)
# Funzione di log to CSV
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
    $Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Doc: ', ';').Replace('/', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = (Read-Host -Prompt 'CSV Path o ClientCode').Trim('"')
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
else {
    $Rev = Read-Host -Prompt 'Issue Index'
    $csv = [PSCustomObject] @{
        ClientCode = $CSVPath
        Rev        = $Rev
        Count      = 1
    }
}
 
try {
    $siteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
    $listName = 'Review Task Panel' 
    $permEdit = 'MT Contributors'
    #$permRead = 'MT Readers'
    $newConsolidator = Read-Host -Prompt "Mail"
    [String]$newConsolidatorName = Read-Host -Prompt "New Consolidator Name"
 

    Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    Write-Host "Caricamento '$($listName)'..." -ForegroundColor Cyan
    $listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID           = $_['ID']
            ClientCode   = $_['ClientCode']
            Rev          = $_['IssueIndex']
            CDL_ID       = $_['IDClientDocumentList']
            ActionReview = $_['ActionReview']
            Consolidator = $_['Consolidator'].Email
        }
    }
    Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

    $rowCounter = 0
    Write-Log 'Inizio correzione...'
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

        Write-Host "Doc: $($row.ClientCode)" -ForegroundColor Blue
        $items = $listItems | Where-Object -FilterScript { $_.ClientCode -eq $row.ClientCode -and $_.Rev -eq $row.Rev }
        if ($null -eq $items) { Write-Log "[ERROR] - List: $($listName) - Doc: $($row.ClientCode) - NOT FOUND" }
        else {
            foreach ($item in $items) {
                #Se inveve si vuole aggiungere il consolidato, aggiungere all'array quello nuovo
                [array]$emailsOldCon = $item.Consolidator
                # Cambio campo Consolidator
                if (!($item.Consolidator.Contains($newConsolidator))) {
                    try {
                        Set-PnPListItem -List $listName -Identity $item.ID -Values @{
                            Consolidator = $newConsolidatorName
                        } -UpdateType SystemUpdate | Out-Null
                        Write-Log "[SUCCESS] - List: $($listName) - ID: $($item.ID) - $($newConsolidator) UPDATED"
                    }
                    catch { Write-Log "[ERROR] - List: $($listName) - ID: $($item.ID) - $($_)" }
                    # rimuove i permessi per i vecchi Consolidator
                    foreach ($mail in $emailsOldCon) {
                        try {
                            Set-PnPListItemPermission -List $listName -Identity $item.ID -User $mail -RemoveRole $permEdit | Out-Null
                            Write-Log "[SUCCESS] - List: $($listName) - ID: $($item.ID) - User: $($mail) - PERMISSION REMOVED"
                        }
                        catch { Write-Log "[ERROR] - List: $($listName) - ID: $($item.ID) - User: $($mail) - $($_)" }                       
                    }
                    # Assegna i permessi per il nuovo Consolidator
                    try {
                        Set-PnPListItemPermission -List $listName -Identity $item.ID -User $newConsolidator -AddRole $permEdit | Out-Null
                        Write-Log "[SUCCESS] - List: $($listName) - ID: $($item.ID) - User: $($newConsolidator) - PERMISSION UPDATED"
                    }
                    catch { Write-Log "[ERROR] - List: $($listName) - ID: $($item.ID) - User: $($newConsolidator) - $($_)" }
                }
            }           
        }
        Write-Host 'Aggiornamento completato.' -ForegroundColor Cyan
    }
}
catch { Throw }
finally { if ($filtered.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }