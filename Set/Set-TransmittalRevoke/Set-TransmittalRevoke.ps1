<# 
WORK IN PROGRESS
Questo script permette di fare il revoque di un transmittal con più documento di un solo documento
LATO VDM
TransmittalQueueDetails_Registry > rimuove il record del documento
Comment Status Report > rimuove il record del documento
TransmittalQueue_Registry > DocumentCount= Valore-1
Vendor Documents List > Pulisce gli attributi:
last Transmittal
last Transmittal Date
Comment Due Date

Process Flow Status List -> no action

LATO CLIENT
Rimuovere record CDL
cancellare la revisione del documento inerente al transmittal
#>


Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$codiceSito,
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the TransmittalNumber')]
    [string]$transmittalNumber,
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the DocumentNumber')]
    [string]$TCM_DN,
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the Revision')]
    [string]$revision


)
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $codiceSito

    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;Level;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - TCMN: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

$SitoVDM = "https://tecnimont.sharepoint.com/sites/vdm_$($codiceSito)"
#$SitoDD = "https://tecnimont.sharepoint.com/sites/$($codiceSito)DigitalDocuments"
$SitoDDc = "https://tecnimont.sharepoint.com/sites/$($codiceSito)DigitalDocumentsc"




$ConnectVDM = Connect-PnPOnline -Url $SitoVDM -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
#$ConnectDD = Connect-PnPOnline -Url $SitoDD -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

<# LATO VDM
TransmittalQueueDetails_Registry > Rimuovere il record del documento

TransmittalQueue_Registry > DocumentCount= Valore-1

Vendor Documents List > Pulisce gli attributi:
last Transmittal
last Transmittal Date
Comment Due Date
#>

$TRN_QDR = Get-PnPListItem -List 'TransmittalQueueDetails_Registry' -Connection $ConnectVDM -PageSize 5000 | ForEach-Object {
    [pscustomobject]@{
        ID            = $_['ID']
        Revision      = $_['IssueIndex']
        TCMN          = $_['Title']
        TransmittalID = $_['TransmittalID']

    }
}
# Filtro elementi della lista contenenti il TransmittalID
$docTRN = $TRN_QDR | Where-Object { $_.TransmittalID -contains $transmittalNumber }
foreach ($doc in $docTRN) {
    try {
        Remove-PnPListItem -List 'TransmittalQueueDetails_Registry' -Identity $doc.ID -Recycle -Force -ErrorAction Stop | Out-Null
        Write-Log "-List - TransmittalQueueDetails_Registry Changes - Documento: $($doc.TCMN) - Rev: $($doc.Revision) - ID: $($doc.ID) Rimosso"
    }
    catch {
        Write-Log "[ERROR]-List - TransmittalQueueDetails_Registry Documento: $($doc.TCMN) - Rev: $($doc.Revision) - ID: $($doc.ID)"
    }
}

#TransmittalQueue_Registry
$TRN_QueueR = Get-PnPListItem -List 'TransmittalQueue_Registry' -Connection $ConnectVDM -PageSize 5000 | ForEach-Object {
    [pscustomobject]@{
        ID             = $_['ID']
        TCMN           = $_['Title']
        DocumentsCount = $_['DocumentsCount']

    }
}

# Imposta il valore del DocumentCount (valore attuale -1)
try {
    $TRNToMod = $TRN_QueueR | Where-Object -FilterScript { $_.TCMN -eq $transmittalNumber }
    $neWDocCount = $TRNToMod.DocumentsCount - 1
    Set-PnPListItem -List 'TransmittalQueue_Registry' -Identity $TRNToMod.ID -Values @{'DocumentsCount' = $($neWDocCount) }
    Write-Log "-List - TransmittalQueue_Registry Changes - Documento: $($TRNToMod.TCMN) - NewDocumentsCount: $($neWDocCount) - ID: $($TRNToMod.ID) Aggiornato"

}
catch {
    Write-Log "[ERROR] -List - TransmittalQueue_Registry Changes - Documento: $($TRNToMod.TCMN) - NewDocumentsCount: $($neWDocCount) - ID: $($TRNToMod.ID)"
}

<#Vendor Documents List > Pulisce gli attributi:
last Transmittal
last Transmittal Date
Comment Due Date#>
$VDLItems = Get-PnPListItem -List 'Vendor Documents List' -Connection $ConnectVDM -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID                  = $_['ID']
        TCMN                = $_['Title']
        Rev                 = $_['IssueIndex']
        lastTransmittal     = $_['LastTransmittal']
        lastTransmittalDate = $_['LastTransmittalDate']
        commentDueDate      = $_['CommentDueDate']
    }
}

try {
    $Item_Mod = $VDLItems | Where-Object -FilterScript { $_.lastTransmittal -eq $transmittalNumber }

    foreach ($item in $Item_Mod) {
        Write-Log "-List: Vendor Documents List Changes - Documento: $($item.TCMN) - ID: $($item.ID) -  LastTransmittal: $($item.lastTransmittal) - LastTransmittalDate : $($item.LastTransmittalDate) - CommentDueDate : $($item.commentDueDate) verranno Rimossi"
        Set-PnPListItem -List 'Vendor Documents List' -Identity $item.ID -Values @{'LastTransmittal' = $null; 'LastTransmittalDate' = $null; 'CommentDueDate' = $nul }
        Write-Log 'rimossi gli attributi: LastTransmittal; LastTransmittalDate ; CommentDueDate'
    }
}
catch {
    Write-Log " [ERROR] -List: Vendor Documents List - Documento: $($item.TCMN) - ID: $($item.ID) -  LastTransmittal: $($item.lastTransmittal) - LastTransmittalDate : $($item.LastTransmittalDate) - CommentDueDate : $($item.commentDueDate)"
}


<# LATO CLIENT
Rimuovere record > CDL
cancellare cartella inerente alla revisione
#>
$ConnectDDc = Connect-PnPOnline -Url $SitoDDc -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

$CDLItems = Get-PnPListItem -List 'Client Document List' -Connection $ConnectDDc -PageSize 5000 | ForEach-Object {
    [pscustomobject]@{
        ID              = $_['ID']
        Revision        = $_['IssueIndex']
        TCMN            = $_['Title']
        LastTransmittal = $_['LastTransmittal']
        DocumentsPath   = $_['DocumentsPath']

    }
}
# Filtro elementi della lista contenenti il TransmittalID
$CDLItemRemove = $CDLItems | Where-Object { $_.LastTransmittal -contains $transmittalNumber }
foreach ($doc in $CDLItemRemove) {
    try {
        if ($doc.Revision -eq $revision) {
            Remove-PnPListItem -List 'Client Document List' -Identity $doc.ID -Recycle -Force -ErrorAction Stop | Out-Null
            Write-Log "-List : Client Document List Changes - Documento: $($doc.TCMN) - Rev: $($doc.Revision) - ID: $($doc.ID) Rimosso"
        }
    }
    catch {
        Write-Log "[ERROR]-List : Client Document List Documento: $($doc.TCMN) - Rev: $($doc.Revision) - ID: $($doc.ID)"
    }
    try {
        $ldaSottrarre = ($revision.Length + 1)
        $lpath = $doc.DocumentsPath.Length
        $llink = $lpath - $ldaSottrarre
        $DocumentsPath = $doc.DocumentsPath.Substring(0, $($llink))
        Write-Host "DocumentPath: $($DocumentsPath)"
        Remove-PnPFolder -Folder $DocumentsPath -Name $revision -Recycle -Force -ErrorAction Stop | Out-Null
        $msg = "[SUCCESS] - Folder:  $DocumentsPath - TCM_DN: $($doc.TCM_DN) - Rev: $($doc.Rev) - DELETED - Folder"
    }
    catch {
        $msg = "[WARNING] - Folder:  $DocumentsPath - TCM_DN: $($doc.TCM_DN) - Rev: $($doc.Rev) - NOT FOUND - Folder"
    }

    Write-Log -Message $msg
}