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

# Funzione SystemUpdate
$system ? ( $updateType = "SystemUpdate" ) : ( $updateType = "Update" ) | Out-Null

# URL del sito
$siteUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($SiteCode)"

# Indentifica il nome della Lista
$VDL = "Vendor Documents List"

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = (Read-Host -Prompt "CSV Path").Trim('"')
$csv = Import-Csv -Path $CSVPath -Delimiter ";"


$vdmConn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
# Ottieni i dati dalla lista
$VDL = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $vdmConn | Where-Object -FilterScript { $_.VD_ClientDocumentNumber -eq $null }
$filedaAggiornare = $VDL | Where-Object -FilterScript { $_.VD_ClientDocumentNumber -eq $null }
forEach ($item in $filedaAggiornare) {
    [PSCustomObject]@{
        ID                      = $_["ID"]
        TCM_DN                  = $_["VD_DocumentNumber"]
        VD_ClientDocumentNumber = $_["VD_ClientDocumentNumber"]
    }
}

#confronta ogni riga.TCM del csv sulla lista
foreach ($row in $csv) {
    $filter = $filedaAggiornare | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM }

}
