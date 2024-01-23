<#
    Rimuove i documenti da un transmittal.
    
    NON FUNZIONANTE!
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
    [parameter(Mandatory = $true)][String]$ProjectCode # Codice sito
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
    $DDUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
    $ClientUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
    $VDMUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"

    while ( $srcSite.ToUpper() -ne "DD" -or $srcSite.ToUpper() -ne "CLIENT" -or $srcSite.ToUpper() -ne "VDM") { $srcSite = (Read-Host -Prompt "Source Site (DD, CLIENT, VDM)").ToUpper() }
    while ( $destSite.ToUpper() -ne "DD" -or $destSite.ToUpper() -ne "CLIENT" -or $destSite.ToUpper() -ne "VDM") { $destSite = (Read-Host -Prompt "Destination Site (DD, CLIENT, VDM)").ToUpper() }

    # Caricamento ID / CSV / Documento
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
        $rev = Read-Host -Prompt "Issue Index"
        $csv = [PSCustomObject]@{
            TCM_DN = $CSVPath
            Rev    = $rev
            Count  = 1
        }
    }

    $ddConn = Connect-PnPOnline -Url $DDUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $clientConn = Connect-PnPOnline -Url $ClientUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $vdmConn = Connect-PnPOnline -Url $VDMUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

    if ($srcSite -eq "DD" -or $destSite -eq "DD") {
        $sourceList = "DocumentList"
        Write-Log "Caricamento '$($sourceList)'..."
        $sourceListItems = Get-PnPListItem -List $sourceList -PageSize 5000 -Connection $ddConn | ForEach-Object {
            [PSCustomObject]@{
                ID            = $_['ID']
                TCM_DN        = $_['Title']
                Rev           = $_['IssueIndex']
                LastTrn       = $_['LastTransmittal']
                LastClientTrn = $_['LastClientTransmittal']
            }
        }
    }
    elseif ($srcSite -eq "CLIENT" -or $destSite -eq "CLIENT") {
        $sourceList = "Client Document List"
        Write-Log "Caricamento '$($sourceList)'..."
        $sourceListItems = Get-PnPListItem -List $sourceList -PageSize 5000 -Connection $clientConn | ForEach-Object {
            [PSCustomObject]@{
                ID            = $_['ID']
                TCM_DN        = $_['Title']
                Rev           = $_['IssueIndex']
                LastTrn       = $_['LastTransmittal']
                LastClientTrn = $_['LastClientTransmittal']
            }
        }
    }
    elseif ($srcSite -eq "VDM" -or $destSite -eq "VDM") {
        Write-Log "Caricamento '$($sourceList)'..."
        $sourceListItems = Get-PnPListItem -List $sourceList -PageSize 5000 -Connection $ddConn | ForEach-Object {
            [PSCustomObject]@{
                ID            = $_['ID']
                TCM_DN        = $_['Title']
                Rev           = $_['IssueIndex']
                LastTrn       = $_['LastTransmittal']
                LastClientTrn = $_['LastClientTransmittal']
            }
        }
    }

}
catch { Throw }
finally {}