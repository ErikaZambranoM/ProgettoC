#Questo script restituisce le revisioni per documento
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$ProjectCode
    
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
        Add-Content $newLog "Timestamp;Type;ListName;ID;Action;Key;Value;OldValue"
    }
    $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($Message.Contains("[SUCCESS]")) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains("[ERROR]")) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains("[WARNING]")) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(" - ID: ", ";").Replace(" - ClientCode: ", ";").Replace(" - IssueIndex: ", ";").Replace(" - Assignee", ";").Replace(" -Consolidator ", ";")
    Add-Content $logPath "$FormattedDate;$Message"
}
# URL del sito
$siteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsc"


# Indentifica il nome della Lista
$RTP = "Review Task Panel"

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = (Read-Host -Prompt "CSV Path").Trim('"')
$PathRev = (Read-Host -Prompt "Path where to download the list Rev").Trim('"')
$csv = Import-Csv -Path $CSVPath -Delimiter ","

$Conn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
# Ottieni i dati dalla lista
Write-Log "Caricamento Lista"
$List = Get-PnPListItem -List $RTP -PageSize 5000 -Connection $Conn | ForEach-Object {
    [PSCustomObject]@{
        ID           = $_["ID"]
        ClientCode   = $_["ClientCode"]
        IssueIndex   = $_["IssueIndex"]
        Assignee     = $_["Assignee"]
        Consolidator = $_["Consolidator"]
    }
}
Write-Log "Lista Caricata"
Write-Log "Creazione file con le Revisioni"
$revPath = "$($PathRev)\$($ProjectCode)Rev.csv";
if (!(Test-Path -Path $revPath)) {
    $newLog = New-Item $revPath -Force -ItemType File
    Add-Content $newLog "ID; ClientCode; IssueIndex; Assignee; Consolidator "
}
#confronta ogni riga.TCM del csv sulla lista
foreach ($row in $csv) {
    Write-Log "Iniziamo:"
    Write-Log "$($row.clientcode)"
    $filter = $List | Where-Object -FilterScript { ($_.ClientCode -eq $row.clientcode) -and ($_.IssueIndex -eq $row.rev) }
    Write-Log "Per il documentento $($filter.Title) ci sono $($filter.count) revisioni"
    foreach ($item in $filter) {
        $rev = "$($item.ID);$($item.ClientCode);$($item.IssueIndex);$($item.Assignee);$($item.Consolidator)" 
        Write-Log "[SUCCESS] $($rev)"
        Add-Content $revPath $rev
    }  
}