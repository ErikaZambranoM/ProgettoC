<#
Questo script imposta i permission levels del gruppo del vendor a seconda dello stato della folder
#>

Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$sitoVendor,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the PO')]
    [string]$listName
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
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Status', ';').Replace(' - Level: ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}
$MTReaders = 'MT Readers'
$MTContributors = 'MT Contributors - Vendor'
# Connessione al sito
$SPOConnection = Connect-PnPOnline -Url $sitoVendor -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction Continue
Write-Log "Caricamento lista $($listName)"
$POFolders = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID                      = $_['ID']
        Name                    = $_['VD_DocumentNumber']
        folderName              = $_['FileLeafRef']
        VD_Status               = $_['VD_DocumentStatus']
        VDMVendorGroup          = $_['VD_AGCCVendorGroup']
        VDMDisciplineOwnerGroup = $_['VD_AGCCDisciplineOwnerGroup']
    }
}
Write-Log "Lista $($listName) caricata"
$counter = 0
foreach ($folder in $POFolders) {
    $counter++
    $folderName = "$($listName)/$($folder.folderName)"
    $VendorGroup = $folder.VDMVendorGroup#Nome gruppo a seconda del VDMVendorGroup
    if ($folder.VD_Status -eq "Placeholder" -or $folder.VD_Status -eq "Rejected") {
        try {
            #Aggiunge il gruppo con PL MT Contributors - Vendor
            Set-PnPFolderPermission -List $listName -Identity $folderName -Group $VendorGroup -AddRole $MTContributors -Connection $SPOConnection | Out-Null
            Write-Log "[SUCCESS] Gruppo $($VendorGroup) - Status : $($folder.VD_Status) -Level: $($MTContributors) - AGGIORNATO"
        }
        catch {
            Write-Log "[ERROR] $($folderName)"
            throw
        }
    }
    elseif ($folder.VD_Status -eq "Closed" -or $folder.VD_Status -eq "Received") {
        try {
            #Aggiunge il gruppo MT Readers - Vendor
            Set-PnPFolderPermission -List $listName -Identity $folderName -Group $VendorGroup -AddRole $MTReaders -Connection $SPOConnection | Out-Null
            Write-Log "[SUCCESS] Gruppo $($VendorGroup) - Status : $($folder.VD_Status) - Level: $($MTReaders) - AGGIORNATO"
        }
        catch {
            Write-Log "[ERROR] $($folderName)"
            throw
        }
    }
    else {
        try {
            #Rimuove il gruppo MT Readers - Vendor
            Set-PnPFolderPermission -List $listName -Identity $VendorGroup -Group $VendorGroup -RemoveRole $MTContributors -Connection $SPOConnection | Out-Null
            Write-Log "[SUCCESS] Gruppo $($OldGroup) - Status : $($folder.VD_Status) - Level: $($MTContributors) - RIMOSSO"
        }
        catch {
            Write-Log "[ERROR] IL GRUPPO $($OldGroup) NON É PRESENTE"
            throw
        }
    }
}
Write-Host 'Aggiornamento Completato'


<######################################################################################
# FIELD VALUES
Discipline: VD_AGCCDisciplineOwnerGroup
Document Status: VD_DocumentStatus

            GRUPPI E PERMISSION LEVELS

    "$($codiceVendor) - Vendor Documents Clients - MT Readers"
    "$($codiceVendor) - Vendor Documents Owners -  Full Control"
    "$($codiceVendor) - Vendor Documents Project Admins - MT Readers"
    "$($codiceVendor) - Vendor Documents Project Readers - MT Readers"
    "$($codiceVendor) - Vendor Documents VDL Editors - MT Readers"
    DS Discipline - MT Contributors
    VD Vendor - MT Contributors - Vendor
    Vendor Documents - User Full Control
######################################################################################>