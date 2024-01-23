# Questo script consente di inserire un gruppo di sicurezza e garantire i permessi per un certo PO

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
    [parameter(Mandatory = $true)][String]$SiteUrl, # URL sito
    [parameter(Mandatory = $true)][String]$ListName, # PO Number
    [parameter(Mandatory = $true)][String]$GroupName # Group name da inserire
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
        Add-Content $newLog 'Timestamp;Type;ListName;ID;Action;Key;Value;OldValue'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Folder: ', ';').Replace(' - Status: ', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

try {
    # Connessione al sito
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    $siteCode = (Get-PnPWeb).Title.Split(' ')[0]
    $listGroup = Get-PnPGroup -Identity $GroupName

    # Caricamento document library del PO
    Write-Log "Caricamento '$($ListName)'..."
    $listItems = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID              = $_['ID']
            Name            = $_['FileLeafRef']
            TCM_DN          = $_['VD_DocumentNumber']
            Status          = $_['VD_DocumentStatus']
            DisciplineOwner = $_['VD_AGCCDisciplineOwnerGroup']
            Vendor          = $_['VD_AGCCVendorGroup']
        }
    }
    Write-Log 'Libreria caricata con successo.'

    # Filtro cartelle da aggiornare
    $filtered = $listItems | Where-Object -FilterScript { $_.Status -ne $null }

    $rowCounter = 0
    Write-Log 'Inizio operazione...'

    # Rimozione gruppo Vendor sbagliato
    if ($GroupName.StartsWith("VD") -and $filtered[0].Vendor -ne $GroupName) {
        $context = Get-PnPContext
        $list = Get-PnPList -Identity $ListName
        $oldGroup = Get-PnPGroup -Identity $filtered[0].Vendor
        Set-PnPList -Identity $ListName -BreakRoleInheritance -CopyRoleAssignments | Out-Null
        $list.RoleAssignments.GetByPrincipal($oldGroup).DeleteObject() | Out-Null
        $context.ExecuteQuery()
        Set-PnPList -Identity $ListName -ResetRoleInheritance | Out-Null
        Write-Log "[SUCCESS] - Library: $($ListName) - Group: $($filtered[0].Vendor) - REMOVED"
    }

    foreach ($folder in $filtered) {
        if ($filtered.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($filtered.Count)" -PercentComplete (($rowCounter++ / $filtered.Count) * 100) }

        Write-Host "Folder: $($folder.Name)" -ForegroundColor Blue

        $folderName = "$($ListName)/$($folder.Name)"

        try {
            if ($GroupName.StartsWith("DS")) {   
                if ($folder.DisciplineOwner -ne $GroupName) {
                    $previousGroup = Get-PnPGroup -Identity $folder.DisciplineOwner
                    Set-PnPListItem -List $ListName -Identity $folder.ID -Values @{ VD_AGCCDisciplineOwnerGroup = $GroupName } -UpdateType SystemUpdate | Out-Null
                    Write-Log "[WARNING] - Library: $($ListName) - Folder: $($folder.TCM_DN) - UPDATED DisciplineOwener Group"

                    if ($folder.Status -eq "Closed") {
                        Set-PnPFolderPermission -List $ListName -Identity $folderName -Group $listGroup.Id -AddRole 'MT Readers' -SystemUpdate
                        Set-PnPFolderPermission -List $ListName -Identity $folderName -Group $previousGroup.Id -RemoveRole 'MT Readers' -SystemUpdate
                    }
                    else {
                        Set-PnPFolderPermission -List $ListName -Identity $folderName -Group $listGroup.Id -AddRole 'MT Contributors' -SystemUpdate
                        Set-PnPFolderPermission -List $ListName -Identity $folderName -Group $previousGroup.Id -RemoveRole 'MT Contributors' -SystemUpdate
                    }
                    Write-Log "[SUCCESS] - Library: $($ListName) - Folder: $($folder.TCM_DN) - Status: $($folder.Status)"
                }
            }
            else {
                # Aggiornamento attributo Vendor Group
                if ($folder.Vendor -ne $GroupName) {
                    Set-PnPListItem -List $ListName -Identity $folder.ID -Values @{ VD_AGCCVendorGroup = $GroupName } -UpdateType SystemUpdate | Out-Null
                    Write-Log "[WARNING] - Library: $($ListName) - Folder: $($folder.TCM_DN) - UPDATED Vendor Group"
                }

                # Aggiornamento permessi in base allo status
                if ($folder.Status -eq 'Placeholder') {
                    Set-PnPFolderPermission -List $ListName -Identity $folderName -Group $listGroup.Id -AddRole 'MT Contributors - Vendor' -RemoveRole 'MT Readers' -SystemUpdate
                    $folderNative = "$($ListName)/$($folder.Name)/Native"
                    $folderUpload = "$($ListName)/$($folder.Name)/Upload Area"
                    Set-PnPFolderPermission -List $ListName -Identity $folderNative -Group $listGroup.Id -AddRole 'MT Contributors' -RemoveRole 'MT Readers' -SystemUpdate
                    Write-Log "[SUCCESS] - Library: $($ListName) - Folder: $($folder.TCM_DN)/Native"
                    Set-PnPFolderPermission -List $ListName -Identity $folderUpload -Group $listGroup.Id -AddRole 'MT Contributors' -RemoveRole 'MT Readers' -SystemUpdate
                    Write-Log "[SUCCESS] - Library: $($ListName) - Folder: $($folder.TCM_DN)/Upload"
                }
                elseif (($folder.Status -eq 'Received' ) -or ($folder.Status -eq 'Closed')) {
                    Set-PnPFolderPermission -List $ListName -Identity $folderName -Group $listGroup.Id -AddRole 'MT Readers' -SystemUpdate
                }
                elseif (($folder.Status -eq 'Commenting') -or ($folder.Status -eq 'Comment Complete')) {
                    try { Set-PnPFolderPermission -List $listName -Identity $folderName -Group $listGroup -RemoveRole 'MT Readers' -SystemUpdate } catch {}
                }
                Write-Log "[SUCCESS] - Library: $($ListName) - Folder: $($folder.TCM_DN) - Status: $($folder.Status)"
            }
        }
        catch { Write-Log "[ERROR] - Library: $($ListName) - Folder: $($folder.TCM_DN) - $($_)" }
    }
    Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($filtered.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }