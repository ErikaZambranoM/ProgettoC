#Revoke Lato Client - Work in progress
# Scaricare dalla Lista Detail registy tutti documenti con le rispettove rev
# filtrare i documenti sulla Client Document List e svuotare gli attributi: Last Client Transmittal, Last Client Transmittal Date
# Cartella = spostare i documenti dalla Originals e cancellare quelli nella ROOT
# Results from TransmittalFromClient_Archive = cancellare PDF
# mettere il transmittal sulla registry come deleted
# DD: DL sui documenti con numero e revisioni uguali ai documenti filtrati cancellare la cartella CMTD dei documenti e 
# Svuotare i campi: Approval User, Actual Date, Last Client Transmittal, Last Client Transmittal Date, Client Acceptance Status
# VDM: VDL sui documenti con numero e revisioni uguali ai documenti filtrati cancellare la cartella CMTD dei documenti e 
# Scuotare i campi: Last Client Transmittal, Last Client Transmittal Date, Client Acceptance Status, User Approval Comments

#parametri

Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$codiceSito,
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the TransmittalNumber')]
    [string]$transmittalNumber
)

# Funzione registro Log
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
$SitoDD = "https://tecnimont.sharepoint.com/sites/$($codiceSito)DigitalDocuments"
$SitoDDc = "https://tecnimont.sharepoint.com/sites/$($codiceSito)DigitalDocumentsc"
$CDL = 'Client Document List'
$CLTRN = 'ClientTransmittalQueueDetails_Registry'
$DL = 'DocumentList'
$VDL = 'Vendor Docuemnts List'
$ListTRNpdf = 'TransmittalToClient_Archive'


#connessione DDClient
Connect-PnPOnline -Url $SitoDDc -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

#filtro sulla listra Details Registry
Write-Host "Caricamento Lista $($CLTRN)"
$ListDocs_TRN = Get-PnPListItem -List $CLTRN -PageSize 5000 | Where-Object -FilterScript { $_.FieldValues.TransmittalID -eq $transmittalNumber } | ForEach-Object {
    [pscustomobject]@{
        ID            = $_['ID']
        Revision      = $_['IssueIndex']
        TCMN          = $_['Title']
        TransmittalID = $_['TransmittalID']

    }
}
Write-Host "Caricamento Lista $($CDL)"
$CDList = Get-PnPListItem -List $CDL -PageSize 5000 | Where-Object -FilterScript { $_.FieldValues.LastClientTransmittal -eq $transmittalNumber } | 
ForEach-Object {
    [PSCustomObject]@{
        ID                        = $_['ID']
        Revision                  = $_['IssueIndex']
        TCMN                      = $_['Title']
        ClientCode                = $_['ClientCode']
        LastClientTransmittal     = $_['LastClientTransmittal']
        LastClientTransmittalDate = $_['LastClientTransmittalDate']
        DocumentsPath             = $_['DocumentsPath']
    }
}
if ($ListDocs_TRN.count -ne $CDList.count) {
    Write-Host "Attenzione il numero dei documenti presenti nel transmittal non coincide con il numero di documenti presenti nella $($CDL)"
    Exit
}
else {
    Write-Host "Totale elementi del transmittal $($ListDocs_TRN.count)"
}

#modifica documenti DDClient
try {
    foreach ($Doc in $CDList) {
        if ($Doc.LastClientTransmittal -eq $transmittalNumber) {
            Set-PnPListItem -List $CDL -Identity $Doc.ID -Values @{
                $Doc.LastClientTransmittal     = $null
                $Doc.LastClientTransmittalDate = $null
            }
            Write-Log "[SUCCESS] - List: $($CDL) - Doc: $($Doc.TCMN)/$($Doc.Revision) - ID $($Doc.ID) - UPDATED`n$($Doc.LastClientTransmittalDate):"
        }
    }
    $pathSplit = $Doc.DocPath.Split('/')
    $folderRelPath = ($pathSplit[5..($pathSplit.Length)] -join '/')
    $fileName = "$($Doc.ClientCode)_$($Doc.rev).pdf"
    $sourceUrl = "$($folderRelPath)/$($fileName)"
    $TargetUrl = ($pathSplit[5..6] -join '/')
    #cancellare i docuemnti nella root
    #Remove-PnPFile -SiteRelativeUrl _catalogs/themes/15/company.spcolor -Recycle
    #Remove-PnPFile -SiteRelativeUrl $sourceUrl -Recycle
    #nome documento= $clientCode + _ + $rev + .pdf
    #spostare i documenti dalla cartell Originals alla Root e cancellare la cartella Originals
    #Move-PnPFile -SourceUrl "Shared Documents/company.docx" -TargetUrl "SubSite2/Shared Documents" -NoWait
    #Move-PnPFile -SourceUrl $sourceUrls -TargetUrl $TargetUrl -NoWait
}
catch {
    throw
}

#Connessione a DD
Connect-PnPOnline -Url $SitoDD -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

#filtro lista TransmittalID
$DDList = Get-PnPListItem -List $DL -PageSize 5000 | Where-Object -FilterScript { $_.FieldValues.LastClientTransmittal -eq $transmittalNumber } |
ForEach-Object {
    [PSCustomObject]@{
        ID                        = $_['ID']
        Revision                  = $_['IssueIndex']
        TCMN                      = $_['Title']
        LastClientTransmittal     = $_['LastClientTransmittal']
        LastClientTransmittalDate = $_['LastClientTransmittalDate']
        ApprovalUser              = $_['ApprovalUser']
        ActualDate                = $_['Actual Date']
        ClientAcceptanceStatus    = $_['ClientAcceptanceStatus']
        DocumentsPath             = $_['DocumentsPath']
    }
}
# Svuotare i campi: Approval User, Actual Date, Last Client Transmittal, Last Client Transmittal Date, Client Acceptance Status
try {
    foreach ($Doc in $DDList) {
        $Doc.LastClientTransmittal = $null
        $Doc.LastClientTransmittalDate = $null
        $Doc.ApprovalUser = $null
        $Doc.ActualDate = $null
        $Doc.ClientAcceptanceStatus = $null
    }
    
    $pathSplit = $Doc.DocPath.Split('/')
    $folderRelPath = ($pathSplit[5..($pathSplit.Length)] -join '/')
    #cancellare cartella CMTD

}
catch {
    throw
}

# VDM: VDL sui documenti con numero e revisioni uguali ai documenti filtrati cancellare la cartella CMTD dei documenti e 



#Connessione a VDM
Connect-PnPOnline -Url $SitoVDM -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

#filtro lista TransmittalID
$VDList = Get-PnPListItem -List $VDL -PageSize 5000 | Where-Object -FilterScript { $_.FieldValues.LastClientTransmittal -eq $transmittalNumber } |
ForEach-Object {
    [PSCustomObject]@{
        ID                        = $_['ID']
        Revision                  = $_['IssueIndex']
        TCMN                      = $_['Title']
        LastClientTransmittal     = $_['LastClientTransmittal']
        LastClientTransmittalDate = $_['LastClientTransmittalDate']
        ApprovalUser              = $_['ApprovalUser']
        ActualDate                = $_['Actual Date']
        ClientAcceptanceStatus    = $_['ClientAcceptanceStatus']
        DocumentsPath             = $_['DocumentsPath']
    }
}
#Svuotare i campi: Last Client Transmittal, Last Client Transmittal Date, Client Acceptance Status, User Approval Comments
try {
    foreach ($Doc in $VDList) {
        $Doc.LastClientTransmittal = $null
        $Doc.LastClientTransmittalDate = $null
        $Doc.ApprovalUser = $null
        $Doc.ClientAcceptanceStatus = $null
    }
    
    $pathSplit = $Doc.DocPath.Split('/')
    $folderRelPath = ($pathSplit[5..($pathSplit.Length)] -join '/')
    #cancellare cartella CMTD

}
catch {
    throw
}

Write-Host "Ora che è tutto completato con successo, metti il transmittal $($transmittalNumber) in stato revoke/deleted sulla lista ClientTransmittalRegistry"