$siteUrl = 'https://tecnimont.sharepoint.com/sites/vdm_K484'
$subSite = 'https://tecnimont.sharepoint.com/sites/vdm_K484/V_1554'
$PONumber = '7500117153'
#$relPath = '/sites/vdm_K484/V_1554/7500117153/'
$PFS = 'Process Flow Status List'
$Vendor = 'VD TrilliumPumpsItalySpa'

try {
    $csv = Import-Csv -Path 'C:\Users\ST-471\Downloads\7500117153.csv' -Delimiter ';'

    $mainConn = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ReturnConnection
    $subConn = Connect-PnPOnline -Url $subSite -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ReturnConnection
    $vendorGroup = Get-PnPGroup -Identity $Vendor -Connection $subConn

    Write-Host "Caricamento '$($PFS)'..." -ForegroundColor Cyan
    $PFSItems = Get-PnPListItem -List $PFS -PageSize 5000 -Connection $mainConn | ForEach-Object {
        [PSCustomObject] @{
            ID     = $_['ID']
            TCM_DN = $_['VD_DocumentNumber']
            Rev    = $_['VD_RevisionNumber']
            Index  = $_['VD_Index']
            Status = $_['VD_DocumentStatus']
        }
    }
    Write-Host "Caricamento lista completato." -ForegroundColor Cyan

    Write-Host "Caricamento '$($PONumber)'..." -ForegroundColor Cyan
    $POItems = Get-PnPListItem -List $PONumber -PageSize 5000 -Connection $subConn | ForEach-Object {
        [PSCustomObject] @{
            ID     = $_['ID']
            Name   = $_['FileLeafRef']
            PFS_ID = $_['VD_AGCCProcessFlowItemID']
        }
    }
    Write-Host "Caricamento libreria completato." -ForegroundColor Cyan

    Write-Host 'Inizio operazioni...' -ForegroundColor Cyan
    $rowCounter = 0
    ForEach ($row in $csv) {
        Write-Progress -Activity 'Update' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100)

        Write-Host "Doc: $($row.TCM_DN)/$($row.Index)" -ForegroundColor Blue

        $PFSFilter = $PFSItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Index -eq $row.Index }

        if ($null -eq $PFSFilter) { Write-Host "[ERROR] - List: $($PFS) - ITEM NOT FOUND" -ForegroundColor Red }
        elseif ($PFSFilter.Status -eq "Placeholder") { Continue }
        else {
            $folder = $POItems | Where-Object -FilterScript { $_.PFS_ID -eq $PFSFilter.ID }

            if ($null -eq $folder) { Write-Host "[ERROR] - List: $($PONumber) - FOLDER NOT FOUND" -ForegroundColor Red }
            else {
                $folderRelPath = "$($PONumber)/$($folder.Name)"
                try {
                    Set-PnPListItem -List $PONumber -Identity $folder.ID -Values @{
                        VD_DocumentStatus = $PFSFilter.Status
                    } -UpdateType SystemUpdate -Connection $subConn | Out-Null

                    if ($PFSFilter.Status -eq 'Received' -or $PFSFilter.Status -eq 'To Accept') {
                        Set-PnPFolderPermission -List $PONumber -Identity $folderRelPath -Group $vendorGroup.Id -RemoveRole 'MT Contributors - Vendor' -AddRole 'MT Readers' -SystemUpdate -Connection $subConn
                    }
                    elseif ($PFSFilter.Status -eq 'Commenting' -or $PFSFilter.Status -eq 'Comment Complete') {
                        Set-PnPFolderPermission -List $PONumber -Identity $folderRelPath -Group $vendorGroup.Id -RemoveRole 'MT Contributors - Vendor' -SystemUpdate -Connection $subConn
                    }

                    Write-Host "[SUCCESS] - List: $($PONumber) - Name: $($folder.Name) - Status: $($PFSFilter.Status) - UPDATED" -ForegroundColor Green
                }
                catch { Write-Host "[ERROR] - List: $($PONumber) - Name: $($folder.Name) - FAILED - $($_)" -ForegroundColor Red }
            }
        }
    }

}
catch { Throw }