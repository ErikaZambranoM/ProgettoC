#Questo script va a lavorare sel sotto sito del vendo, va ad impostare impostare i permessi corretti a secondo del PO

Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$codiceSito,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the Vendor Code')]
    [string]$codiceVendor,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the PO')]
    [string]$PO,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the OLD Group')]
    [string]$OldGroup,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the NEW Group')]
    [string]$NewGroup

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
        Add-Content $newLog 'Timestamp;Type;ListName;Code;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Code: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}
<#
$Sito= "https://tecnimont.sharepoint.com/sites/vdm_$($codiceSito)" #prod
$sitoVendor = "$($sito)/V_$($codiceVendor)/$($PO)"#prod
$VendorGroup = "Vendor Documents VDL Editors"
$VendorGroupPL = "MT Readers"
$GruppoVDPL = "MT Contributors - Vendor"
$UserVD= "Vendor Documents"
$UserVDPL= "Full Control"
#>

$Sito = "https://tecnimont.sharepoint.com/sites/$($codiceSito)_VDM" #POC
$sitoVendor = "$($sito)/$($codiceVendor)"#POC
$listName = $PO
$VendorGroup = 'POC Vendor Documents Management VDL Editors'
$VendorGroupPL = 'MT Readers'
$MTReaders = 'MT Readers'
$MTContributor = 'MT Contributors - Vendor'
$UserVD = 'Flow Cloud System User'
$UserVDPL = 'Full Control'
# Connessione al sito
$SPOConnection = Connect-PnPOnline -Url $sitoVendor -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction Continue

$POFolders = Get-PnPListItem -List $listName -PageSize 5000 -Connection $SPOConnection | ForEach-Object {
    [PSCustomObject]@{
        ID                      = $_['ID']
        Name                    = $_['VD_DocumentNumber']
        VDMVendorGroup          = $_['VD_AGCCVendorGroup']
        VDMDisciplineOwnerGroup = $_['VD_AGCCDisciplineOwnerGroup']
    }
}
$counter = 0
foreach ($folder in $POFolders) {
    $counter = $counter + 1
    Write-Log "Folder $($counter) è $($folder)"
    if ($folder.VDMVendorGroup -eq $OldGroup) {
        Write-Progress -Activity 'Aggiornamento' -Status "$($folder.Name)"
        Write-Log "$($folder.Name) and $($counter)"
        $folderName = "$($listName)/$($folder.Name)-000"
        Write-Log "Folder $($folderName)"
        try {
            #Setta in nuovo gruppo nel VDMVendorGroup
            Set-PnPListItem -List $listName -Identity $counter -Values @{VD_AGCCVendorGroup = $NewGroup }
            Write-Log "[SUCCESS] Aggiornato VD_AGCCVendorGroup a $($NewGroup) - AGGIUNTO"
            $permessi = Get-PnPListItemPermission -List $listName -Identity $counter -Connection $SPOConnection
            $permessi | ForEach-Object {
                #permessi presenti nella cartella e i permission levels
                $permessiL = $permessi.Permissions
                if ($permessiL.PrincipalName.Contains($OldGroup)) {
                    Write-Log "[SUCCESS] Gruppo $($OldGroup)"
                    if ($permessiL.PermissionLevels.Name -eq $MTContributor) {
                        Write-Log "con PermissionLevels $($MTContributor) rimosso"
                        Set-PnPGroupPermissions -List $listName -Identity $counter -RemoveRole $MTContributor -Connection $SPOConnection | Out-Null
                    }
                    else {
                        Write-Log "[SUCCESS] Gruppo $($permessiL.PrincipalName) con PErmissionLevels: $($MTReaders)"
                        Set-PnPGroupPermissions -List $listName -Identity $counter -RemoveRole $MTReaders -Connection $SPOConnection | Out-Null
                    }

                }
                if ($permessiL.PrincipalName.Contains($NewGroup)) {
                    if ($permessiL.PermissionLevels.Name -eq $MTContributor) {
                        <# Verificare che $NewGroup ha i permission levels come $MTContributor #>
                        Write-Log "[WARNING] I permessi per $($NewGroup) sono corretti"
                    }
                }
                if ($permessiL.PrincipalName.Contains($VendorGroup)) {
                    if ($permessiL.PermissionLevels.Name -eq $VendorGroupPL) {
                        Write-Log "[WARNING] I permessi per $($VendorGroup) sono corretti $($permessiL.PermissionLevels.Name)"
                    }
                }
                if ($permessiL.PrincipalName.Contains($UserVD)) {
                    if ($permessiL.PermissionLevels.Name -eq $UserVDPL) {
                        Write-Log "[WARNING] I permessi per $($UserVDPL) sono corretti"
                    }

                }
                else {
                    Write-Host 'Gruppi mancanti: '

                    if (!($permessiL.PrincipalName.Contains($VendorGroup))) {
                        Write-Log "[SUCCESS] $($VendorGroup) - AGGIUNTO"
                        Set-PnPFolderPermission -List $listName -Identity $folderName -Group $VendorGroup -AddRole $VendorGroupPL -SystemUpdate
                    }
                    if (!($permessiL.PrincipalName.Contains($NewGroup))) {
                        Write-Log "[SUCCESS] $($NewGroup) - AGGIUNTO"
                        Set-PnPFolderPermission -List $listName -Identity $folderName -Group $NewGroup -AddRole $MTContributor -SystemUpdate
                    }
                    if (!($permessiL.PrincipalName.Contains($UserVD))) {
                        Write-Log "[SUCCESS] $($UserVD) - AGGIUNTO"
                        Set-PnPFolderPermission -List $listName -Identity $folderName -User $UserVD -AddRole $UserVDPL -SystemUpdate
                    }
                }

            }
        }
        catch { Write-Log "[WARNING] - Folder:$($folderName) - VDMVendorGroup is empty" }
    }
}
Disconnect-PnPOnline