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
    $Message = $Message.Replace(" - List: ", ";").Replace(" - ID: ", ";").Replace(" - Doc: ", ";").Replace("/", ";").Replace(" - ", ";")
    Add-Content $logPath "$FormattedDate;$Message"
}
# URL del sito
$siteUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"

# Indentifica il nome della Lista
$VDL = "Vendor Documents List"

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = (Read-Host -Prompt "CSV Path").Trim('"')
$PathRev = (Read-Host -Prompt "Path where to download the list Rev").Trim('"')
$field = (Read-Host -Prompt "Field")
$csv = Import-Csv -Path $CSVPath -Delimiter ";"

$vdmConn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
# Ottieni i dati dalla lista
Write-Log "Caricamento Lista"
$VD = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $vdmConn | ForEach-Object {
    [PSCustomObject]@{
        ID                = $_["ID"]
        TCM_DN            = $_["VD_DocumentNumber"]
        VD_RevisionNumber = $_["VD_RevisionNumber"]
        VD_Index          = $_["VD_Index"]
        LastTrn           = $_['LastTransmittal']
        LastClientTrn     = $_['LastClientTransmittal']
        $field            = $_[$($field)]
    }
}
Write-Log "Lista Caricata"
Write-Log "Creazione file con le Revisioni"
$revPath = "$($PathRev)\$($ProjectCode)Rev.csv";
if (!(Test-Path -Path $revPath)) {
    $newLog = New-Item $revPath -Force -ItemType File
    Add-Content $newLog "TCM_DN; Rev; Index; ClientCode; LastTrn; LastClientTransmittal;$($field) "
}
#confronta ogni riga.TCM del csv sulla lista
foreach ($row in $csv) {
    Write-Log "Iniziamo:"
    Write-Log "$($row.TCM_DN)"
    $filter = $VD | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN }
    Write-Log "Per il documentento $($filter.TCM_DN) ci sono $($filter.count) revisioni"
    foreach ($item in $filter) {
        Write-Log "Per il documentento $($filter.TCM_DN) ci sono $($filter.count) revisioni"
        $rev = "$($item.TCM_DN);$($item.VD_RevisionNumber);$($item.VD_Index);$($row.ClientCode);$($item.LastTrn);$($item.LastClientTransmittal);$($item.$field)" 
        Write-Log "[SUCCESS] $($rev)"
        Add-Content $revPath $rev
    }  
}