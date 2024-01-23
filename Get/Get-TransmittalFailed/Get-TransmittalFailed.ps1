#Questo script verifica i Transmittal partendo dalla Lista della Registry su DD e VDM
#verifica i documenti collegati a ogni transmittal che ha filtrato sulla Queue details registry e 
#controlla nella Document List i documenti se hanno la colonna LAst Tranmittal popolata e il numero del tranmittal
#Al rigo 83 si può modificare la condizione a seconda della colonna che si vuole filtrare
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$siteUrl
    
)
$SplitPjC = $siteUrl.Split('/')
$ProjectCode = $SplitPjC[4..5]
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
    $Message = $Message.Replace(" - List: ", ";").Replace(" - ID: ", ";").Replace(" - Doc: ", ";").Replace("TransmittalID", ";").Replace(" - ", ";")
    Add-Content $logPath "$FormattedDate;$Message"
}
# URL del sito
$detailregistry = "TransmittalQueueDetails_Registry"
$registry = "TransmittalQueue_Registry"
$PathDownload = (Read-Host -Prompt "Path where download the list Transmittal").Trim('"')

if ($siteUrl.Contains("vdm_")) {
    $VDL = "Vendor Documents List"
    $VDMConn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

    $ProjectCode = (get-pnpweb -Connection $VDMConn).Title.Split(' ')[0]
    # Ottieni i dati dalla lista
    Write-Log "Caricamento Lista $($VDL)"
    $VDItems = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $VDMConn | ForEach-Object {
        [PSCustomObject]@{
            ID            = $_["ID"]
            TCM_DN        = $_["VD_DocumentNumber"]
            Rev           = $_["VD_RevisionNumber"]
            ClientCode    = $_['ClientCode']
            VD_Index      = $_["VD_Index"]
            LastTrn       = $_['LastTransmittal']
            LastClientTrn = $_['LastClientTransmittal']
        }
    }
    Write-Log "Lista $($VDL) Caricata"
    $DDConn = $VDMConn
}
else {
    # Indentifica il nome della Lista
    $DD = "DocumentList"
    $DDConn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    # Ottieni i dati dalla lista
    Write-Log "Caricamento Lista $($DD)"
    $DD = Get-PnPListItem -List $DD -PageSize 5000 -Connection $DDConn | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            Rev        = $_['IssueIndex']
            ClientCode = $_['ClientCode']
            LastTrn    = $_['LastTransmittal']
        }
    }
    Write-Log "Lista $($DD) Caricata"
}

Write-Log "Caricamento Lista $($registry)"
$Reg = Get-PnPListItem -List $registry -PageSize 5000 -Connection $DDConn | ForEach-Object { If ($_.FieldValues.TransmittalStatus -ne "*Sent*") {
        [PSCustomObject]@{
            ID                = $_['ID']
            TransmittalID     = $_['Transmittal_x0020_Number']
            Title             = $_['Title']
            Created           = $_['Created']
            DocumentsCount    = $_['DocumentsCount']
            TransmittalStatus = $_['TransmittalStatus']
            Modified          = $_['Modified']
        }
    }
}
Write-Log "Lista $($registry) Caricata"
Write-Log "Caricamento Lista $($detailregistry)"
$DetailReg = Get-PnPListItem -List $detailregistry -PageSize 5000 -Connection $DDConn | ForEach-Object {
    [PSCustomObject]@{
        ID            = $_['ID']
        TransmittalID = $_['TransmittalID']
        Title         = $_['Title']
        IssueIndex    = $_['IssueIndex']
        Created       = $_['Created_x0020_Date']
        
    }
}
Write-Log "Lista $($detailregistry) Caricata"
Write-Log "Creazione file excel con Transmittal"
$Path = "$($PathDownload)\$($ProjectCode).csv";
if (!(Test-Path -Path $Path)) {
    $newLog = New-Item $Path -Force -ItemType File
    Add-Content $newLog "TCM_DN; ClientCode; TransmittalID; "
}
foreach ($doc in $Reg) {
    Write-Log "Iniziamo:"
    Write-Log "$($doc.TransmittalID)"
    $item = $DetailReg | Where-Object -FilterScript { $_.TransmittalID -eq $doc.Title }
    #confronta ogni riga dei transmittal con Status != Sent sulla lista DocumentList
    $Transmittal = "Tranmittal $($doc.TransmittalID) - TransmittalStatus $($doc.TransmittalStatus) - Totale Documenti $($doc.DocumentsCount) - Created $($doc.Created)" 
    Add-Content $Path $Transmittal
    foreach ($i in $item) {
        Write-Log "Documentento $($i.Title), Rev $($i.IssueIndex)"
        if ($siteUrl.Contains("vdm_")) {
            $filter = $VDItems | Where-Object -FilterScript { $_.TCM_DN -eq $i.Title -and $_.Rev -eq $i.IssueIndex }
            Write-Log "Per il documentento $($filter.TCM_DN) Rev $($filter.Rev) "
            $rev = "$($filter.TCM_DN); $($filter.ClientCode); $($filter.LastTrn); $($filter.Created_x0020_Date)" 
            Write-Log "[SUCCESS] $($rev)"
            Add-Content $Path $rev 
        }
        else {
            $filter = $DD | Where-Object -FilterScript { $_.TCM_DN -eq $i.Title -and $_.Rev -eq $i.IssueIndex }
            Write-Log "Per il documentento $($filter.TCM_DN) Rev $($filter.Rev) "
            $rev = "$($filter.TCM_DN); $($filter.ClientCode); $($filter.LastTrn); $($filter.Created_x0020_Date)" 
            Write-Log "[SUCCESS] $($rev)"
            Add-Content $Path $rev
        }
    }
}
