<#
	Consente di aggiornare i permessi delle Discipline sulle cartelle di VDM 
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
    [parameter(Mandatory = $true)][string]$SiteCode #URL del sito
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

# Funzione che restituisce array di ID delle Discipline
function Find-ID {
    param (
        [Parameter(Mandatory = $true)][String[]]$Items,
        [Array]$List = $discipline
    )

    $array = @()
    $Items | ForEach-Object {
        $currItem = $_.Trim()
        $found = $discipline | Where-Object -FilterScript { $_.Title -eq $currItem }
        if ($null -ne $found) { $array += $found.ID }
        else {
            Write-Host "Discipline '$($_)' not found." -ForegroundColor Red
            Exit
        }
    }
    return $array
}

try {
    $siteUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($SiteCode)"
    $VDL = 'Vendor Documents List'
    $PFS = 'Process Flow Status List'
    $DMList = 'Distribution Matrix'
    $dList = 'Disciplines'

    # Caricamento CSV o DM Code
    $CSVPath = (Read-Host -Prompt 'CSV Path o DM Code').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv -Path $CSVPath -Delimiter ';'
        # Validazione colonne
        $validCols = @('Title')
        $validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
            if ($_ -in $validCols) { $validCounter++ }
        }
        if ($validCounter -lt $validCols.Count) {
            Write-Host "Colonne obbligatorie mancanti: $($validCols -join ', ')" -ForegroundColor Red
            Exit
        }
    }
    else {
        $csv = [PSCustomObject]@{
            Title = $CSVPath
            Count = 1
        }
    }

    # Connessione al sito e calcolo del site Code
    $mainConn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

    # Caricamento lista Distribution Matrix
    Write-Log "Caricamento '$($DMList)'..."
    $DMItems = Get-PnPListItem -List $DMList -PageSize 5000 -Connection $mainConn | ForEach-Object {
        $dArray = @()
        $dArray += $_['VD_DisciplineOwnerTCM'].LookupValue
        if ($null -ne $_['VD_DisciplinesTCM'].LookupValue) { $dArray += $_['VD_DisciplinesTCM'].LookupValue }
        if ($null -ne $_['VD_DisciplinesTCM_RoleX'].LookupValue) { $dArray += $_['VD_DisciplinesTCM_RoleX'].LookupValue }
        if ($null -ne $_['VD_DisciplinesTCMO_RoleX'].LookupValue) { $dArray += $_['VD_DisciplinesTCM_ RoleO'].LookupValue }
        [PSCustomObject]@{
            ID                 = $_['ID']
            Title              = $_['Title']
            DisciplineOwnerTCM = $_['VD_DisciplineOwnerTCM'].LookupValue
            DisciplinesTCM     = $_['VD_DisciplinesTCM'].LookupValue
            DisciplinesTCMX    = $_['VD_DisciplinesTCM_RoleX'].LookupValue
            DisciplinesTCMO    = $_['VD_DisciplinesTCM_RoleO'].LookupValue
            DisciplinesArray   = $dArray
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Caricamento lista Disciplines
    Write-Log "Caricamento '$($dList)'..."
    $discipline = Get-PnPListItem -List $dList -PageSize 5000 -Connection $mainConn | ForEach-Object {
        [PSCustomObject]@{
            ID    = $_['ID']
            Title = $_['Title']
        }
    }
    Write-Log 'Caricamento lista completato.'

    # Caricamento VDL
    Write-Log "Caricamento '$($VDL)'..."
    $VDLItems = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $mainConn | ForEach-Object {
        $dArray = @()
        $dArray += $_['VD_DisciplineOwnerTCM'].LookupValue
        if ($null -ne $_['VD_DisciplinesTCM'].LookupValue) { $dArray += $_['VD_DisciplinesTCM'].LookupValue }
        [PSCustomObject]@{
            ID                     = $_['ID']
            TCM_DN                 = $_['VD_DocumentNumber']
            Rev                    = $_['VD_RevisionNumber']
            DistributionMatrixCode = $_['VD_DistributionMatrix_Calc']
            DisciplineOwnerTCM     = $_['VD_DisciplineOwnerTCM'].LookupValue
            DisciplinesTCM         = $_['VD_DisciplinesTCM'].LookupValue
            DisciplinesArray       = $dArray
            DocPath                = $_['VD_DocumentPath']
        }
    }
    Write-Log 'Caricamento lista completato.'
    
    # Caricamento PFS
    Write-Log "Caricamento '$($PFS)'..."
    $PFSItems = Get-PnPListItem -List $PFS -PageSize 5000 -Connection $mainConn | ForEach-Object {
        [PSCustomObject]@{
            ID     = $_['ID']
            TCM_DN = $_['VD_DocumentNumber']
            Rev    = $_['VD_RevisionNumber']
            Status = $_['VD_DocumentStatus']
            VDL_ID = $_['VD_VDL_ID']
        }
    }
    Write-Log 'Caricamento lista completato.'


    Write-Log 'Inizio operazioni...'
    foreach ($DMCode in $csv.Title) {
        Write-Host "Distribution Matrix Code: $($DMCode)" -ForegroundColor DarkYellow
        
        $found = $DMItems | Where-Object -FilterScript { $_.Title -eq $DMCode }
    
        if ($null -eq $found ) {
            Write-Host "[ERROR] - List: $($DMList) - Code: $($DMCode) - NOT FOUND" -ForegroundColor Red
            continue
        }

        # Filtro DisciplineOwnerTCM
        $dOwner = Find-ID -Items $found.DisciplineOwnerTCM

        $dTCM = @()
        # Filtro DisciplinesTCM
        if ($found.DisciplinesTCM) { $dTCM += Find-ID -Items $found.DisciplinesTCM }
        if ($found.DisciplinesTCMX) { $dTCM += Find-ID -Items $found.DisciplinesTCMX }
        if ($found.DisciplinesTCMO) { $dTCM += Find-ID -Items $found.DisciplinesTCMO }

        # Filtro VDL per Distribution Matrix Code
        $filtered = $VDLItems | Where-Object -FilterScript { $_.DistributionMatrixCode -eq $DMCode }

        $rowCounter = 0
        ForEach ($row in $filtered) {
            if ($filtered.Count -gt 1) { Write-Progress -Activity 'DM Code' -Status "$($rowCounter+1)/$($filtered.Count) - $($DMCode)" -PercentComplete (($rowCounter++ / $filtered.Count) * 100) }

            # Controllo discipline da aggiornare
            $compare = Compare-Object -ReferenceObject $found.DisciplinesArray -DifferenceObject $row.DisciplinesArray
            if ($null -ne $compare) {
                Write-Host "Doc: $($row.TCM_DN)/$($row.Rev)" -ForegroundColor Blue

                # Aggiornamento VDL
                try {
                    Set-PnPListItem -List $VDL -Identity $row.ID -Values @{
                        VD_DisciplineOwnerTCM = $dOwner
                        VD_DisciplinesTCM     = $dTCM
                    } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
                    Write-Log "[SUCCESS] - List: $($VDL) - Doc: $($row.TCM_DN)/$($row.Rev) - UPDATED"
                }
                catch { Write-Log "[ERROR] - List: $($VDL) - Doc: $($row.TCM_DN)/$($row.Rev) - $($_)" }

                # Filtro PFS
                $PFSItem = $PFSItems | Where-Object -FilterScript { $_.VDL_ID -eq $row.ID }

                if ($null -eq $PFSItem) { Write-Log "[ERROR] - List: $($PFS) - Doc: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
                else {
                    # Aggiornamento PFS
                    try {
                        Set-PnPListItem -List $PFS -Identity $PFSItem.ID -Values @{
                            VD_DisciplineOwnerTCM = $dOwner
                            VD_DisciplinesTCM     = $dTCM
                        } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
                        Write-Log "[SUCCESS] - List: $($PFS) - Doc: $($row.TCM_DN)/$($row.Rev) - UPDATED"
                    }
                    catch { Write-Log "[ERROR] - List: $($PFS) - Doc: $($row.TCM_DN)/$($row.Rev) - $($_)" }

                    # Variabili posizione del documento
                    $pathSplit = $row.DocPath.Split('/')
                    $subSite = $pathSplit[0..5] -join '/'
                    $POLibrary = $pathSplit[6]
                    $folderPath = "$($pathSplit[6])/$($pathSplit[7])"

                    # Connessione al sottosito del Vendor
                    $subConn = Connect-PnPOnline -Url $subSite -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

                    # Ciclo per ogni disciplina da aggiungere / togliere
                    foreach ($item in $compare) {
                        $dGroup = Get-PnPGroup -Identity "DS $($item.InputObject)" -Connection $subConn

                        try {
                            if ($item.SideIndicator -eq '<=') {
                                if ($PFSItem.Status -eq 'Closed') {
                                    Set-PnPFolderPermission -List $POLibrary -Identity $folderPath -Group $dGroup.Id -AddRole 'MT Readers' -SystemUpdate -Connection $subConn
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/FromClient" -Group $dGroup.Id -AddRole 'MT Readers' -SystemUpdate -Connection $subConn } catch {}
                                }
                                else {
                                    Set-PnPFolderPermission -List $POLibrary -Identity $folderPath -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/Native" -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/Upload Area" -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/ATTACHMENTS" -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/OFV" -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/CMTD" -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/FromClient" -Group $dGroup.Id -AddRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                }
                                Write-Log "[SUCCESS] - List: $($POLibrary) - Folder: $($folderPath) - Status: $($PFSItem.Status) - Discipline: $($dGroup.LoginName) - ADDED"
                            }
                            elseif ($item.SideIndicator -eq '=>') {
                                if ($PFSItem.Status -eq 'Closed') {
                                    Set-PnPFolderPermission -List $POLibrary -Identity $folderPath -Group $dGroup.Id -RemoveRole 'MT Readers' -SystemUpdate -Connection $subConn
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/FromClient" -Group $dGroup.Id -RevoveRole 'MT Readers' -SystemUpdate -Connection $subConn } catch {}
                                }
                                else {
                                    Set-PnPFolderPermission -List $POLibrary -Identity $folderPath -Group $dGroup.Id -RemoveRole 'MT Contributors' -SystemUpdate -Connection $subConn
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/Native" -Group $dGroup.Id -RevoveRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/Upload Area" -Group $dGroup.Id -RevoveRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/ATTACHMENTS" -Group $dGroup.Id -RevoveRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/OFV" -Group $dGroup.Id -RevoveRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/CMTD" -Group $dGroup.Id -RevoveRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                    try { Set-PnPFolderPermission -List $POLibrary -Identity "$($folderPath)/FromClient" -Group $dGroup.Id -RevoveRole 'MT Contributors' -SystemUpdate -Connection $subConn } catch {}
                                }
                                Write-Log "[SUCCESS] - List: $($POLibrary) - Folder: $($folderPath) - Status: $($PFSItem.Status) - Discipline: $($dGroup.LoginName) - REMOVED"
                            }
                        }
                        catch {
                            Write-Log "[ERROR] - List: $($POLibrary) - Folder: $($folderPath) - Status: $($PFSItem.Status) - Discipline: $($dGroup.LoginName) - $($_)"
                            Throw
                        } 
                    }
                }
            }
        }
    }
    Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($filtered.Count -gt 1) { Write-Progress -Activity 'DM Code' -Completed } }