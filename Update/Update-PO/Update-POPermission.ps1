<#
IN Progress
Questo script va a lavorare sul sotto sito del vendo, SOLO per i documenti in stato PLACEHOLDER
va a settare il gruppo segregato per il PO, gruppo VD e Vendor Documents con i permissionLevels :
Se si vuole settare il gruppo Vendor Documents VDL Editors e l'user Vendor documents scommentare le righe dalla 84 a 89
Vendor Documents VDL Editors : MT Readers
Vendor Documents : User Full Control

---- TODO -----
impostare i permission level in base allo stato del documento - if (Placeholder or Rejected){$_['VD_DocumentStatus'] -eq 'Placeholder' -or $_['VD_DocumentStatus'] -eq 'Rejected'} VD Vendor : MT Contributors - else VD Vendor : MR Readers
Aggiungere le discipline ai gruppi con i PL
#>

Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$sitoVendor,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the PO')]
    [string]$listName,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the OLD Group')]
    [string]$OldGroup,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the NEW Group')]
    [string]$NewGroup

)
function Write-Log
{
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $codiceSito
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath))
    {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;Level;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else
    {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Level: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}
$VendorGroup = 'Vendor Documents VDL Editors'
$MTReaders = 'MT Readers'
$MTContributors = 'MT Contributors - Vendor'
$UserVD = 'Vendor Documents'
$UserVDPL = 'Full Control'
# Connessione al sito
$SPOConnection = Connect-PnPOnline -Url $sitoVendor -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction Continue
Write-Log "Caricamento lista $($listName)"
$POFolders = Get-PnPListItem -List $listName -PageSize 5000 | Where-Object -FilterScript { $_['VD_DocumentStatus'] -eq 'Placeholder' }
$placeholder = $POFolders | ForEach-Object {
    [PSCustomObject]@{
        ID                      = $_['ID']
        Name                    = $_['VD_DocumentNumber']
        folderName              = $_['FileLeafRef']
        VDMVendorGroup          = $_['VD_AGCCVendorGroup']
        VDMDisciplineOwnerGroup = $_['VD_AGCCDisciplineOwnerGroup']
    }
}
Write-Log "Lista $($listName) caricata"
$counter = 0
foreach ($folder in $placeholder)
{
    $counter++
    $folderName = "$($listName)/$($folder.folderName)"
    #Setta il nuovo gruppo in VDMVendorGroup
    try
    {
        Set-PnPListItem -List $listName -Identity $folder.ID -Values @{VD_AGCCVendorGroup = $NewGroup }
        Write-Log "[SUCCESS] Aggiornato VDMVendorGroup a $($NewGroup) - AGGIORNATO"
        #Aggiunge il gruppo "New" con PL MT Contributors - Vendor
        Set-PnPFolderPermission -List $listName -Identity $folderName -Group $NewGroup -AddRole $MTContributors -Connection $SPOConnection | Out-Null
        Write-Log "[SUCCESS] Gruppo $($NewGroup) - Level: $($MTContributors) - AGGIORNATO"
        #Aggiunge il gruppo Vendor con PL MT Readers
        #Set-PnPFolderPermission -List $listName -Identity $folderName -Group $VendorGroup -AddRole $MTReaders -SystemUpdate -Connection $SPOConnection
        #Write-Log "[SUCCESS] Gruppo $($VendorGroup) - Level:$($MTReaders) - AGGIORNATO"
        #Aggiunge l'user Vendor Documents con PL Full Control
        #Set-PnPFolderPermission -List $listName -Identity $folderName -User $UserVD -AddRole $UserVDPL -SystemUpdate -Connection $SPOConnection
        #Write-Log "[SUCCESS] - User: $($UserVD) - Level: $($UserVDPL) - AGGIORNATO"
    }
    catch
    {
        Write-Log "[ERROR] $($folderName)"
        throw
    }
    #Rimuove il gruppo "OLD"
    try
    {
        Set-PnPFolderPermission -List $listName -Identity $folderName -Group $OldGroup -RemoveRole $MTContributors -Connection $SPOConnection | Out-Null
        Write-Log "[SUCCESS] Gruppo $($OldGroup) - Level: $($MTContributors) - RIMOSSO"
    }
    catch
    {
        Write-Log "[ERROR] IL GRUPPO $($OldGroup) NON É PRESENTE"
        throw
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