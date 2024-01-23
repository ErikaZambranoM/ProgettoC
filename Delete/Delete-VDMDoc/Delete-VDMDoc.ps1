<#
    Script che rimuove i documenti da VDM.
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param (
    [Parameter(Mandatory = $true)][String]$ProjectCode
)

# Funzione che elimina un elemento da una lista
function Remove-ListItem {
    param (
        [parameter(Mandatory = $true)][String]$List,
        [parameter(Mandatory = $true)][String]$ID
    )

    try {
        Remove-PnPListItem -List $List -Identity $ID -Recycle -Force | Out-Null
        Write-Host "[SUCCESS] - List: $($List) - ID: $($ID) - REMOVED" -ForegroundColor Green
    }
    catch { Write-Host "[ERROR] - List: $($List) - ID: $($ID) - $($_)" -ForegroundColor Red }
}

Try {
    $siteUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
    $PFS = 'Process Flow Status List'
    $VDL = 'Vendor Documents List'
    $CSR = 'Comment Status Report'

    # Caricamento CSV o Documento singolo
    $CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv $CSVPath -Delimiter ';'
        # Validazione colonne
        $validCols = @('TCM_DN', 'Rev')
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
        $rev = Read-Host -Prompt "Revision Number"
        $csv = [PSCustomObject]@{
            TCM_DN = $CSVPath
            Rev    = $rev
            Count  = 1
        }
    }

    # Bypass document status
    $bypass = Read-Host -Prompt "Bypass Doc Status (true/false)"

    # Connessione al sito
    Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue

    # Caricamento Process Flow
    Write-Host "Caricamento '$($PFS)'..." -ForegroundColor Cyan
    $PFSItems = Get-PnPListItem -List $PFS -PageSize 5000 | ForEach-Object {
        [pscustomobject] @{
            ID     = $_['ID']
            TCM_DN = $_['VD_DocumentNumber']
            Rev    = $_['VD_RevisionNumber']
            Index  = [Int]$_['VD_Index']
            VDL_ID = [Int]$_['VD_VDL_ID']
            CSR_ID = [Int]$_['VD_CommentsStatusReportID']
            Status = $_['VD_DocumentStatus']
        }
    }
    Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

    $rowCounter = 0
    Write-Host 'Inizio pulizia...' -ForegroundColor Cyan
    foreach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity "Pulizia" -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

        Write-Host "Doc: $($row.TCM_DN)/$($row.Rev)" -ForegroundColor Blue
        
        # Ricerca documenti
        $filtered = $PFSItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

        if ($null -eq $filtered) { Write-Host "[ERROR] - Documento $($row.TCM_DN)/$($row.Rev) non trovato." -ForegroundColor Red }
        #elseif ($filtered.Count -eq 1) { Write-Host "[SKIPPED] - Documento $($row.TCM_DN)/$($row.Rev) non duplicato." -ForegroundColor Yellow }
        else {
            foreach ($item in $filtered) {
                Write-Host "Index: $($item.Index) - ID: $($item.ID) - VDL_ID: $($item.VDL_ID) - Status: $($item.Status)" -ForegroundColor Cyan
            
                # Se sulla PFS è Deleted, lo cancella direttamente
                if ($item.Status -eq "Deleted") { Remove-ListItem -List $PFS -Id $item.ID }
                else {
                    # Se lo status non è Placeholder e viene dato Bypass, aggiorna lo status a Placeholder per proseguire
                    if ($item.Status -ne 'Placeholder' -and [System.Convert]::ToBoolean($bypass)) {
                        try {
                            Set-PnPListItem -List $PFS -Identity $item.ID -Values @{ VD_DocumentStatus = 'Placeholder' } -UpdateType SystemUpdate | Out-Null
                            $item.Status = 'Placeholder'
                            Write-Host "[WARNING] - List: $($PFS) - ID: $($item.ID) - UPDATED to Placeholder" -ForegroundColor Yellow  
                        }
                        catch {
                            Write-Host "[ERROR] - List: $($List) - ID: $($ID) - $($_)" -ForegroundColor Red
                            Throw
                        }

                        Remove-ListItem -List $CSR -Id $item.CSR_ID
                    }
                    # Se lo status è Placeholder, elimina il record dalla VDL e poi dalla PFS
                    if ($item.Status -eq 'Placeholder') {

                        Remove-ListItem -List $VDL -Id $item.VDL_ID

                        # Aspetta che il PA imposti lo status sulla PFS a Deleted prima di eliminare il record
                        while ($item.Status -ne "Deleted") {
                            Write-Host "Verifica eliminazione..." -ForegroundColor DarkGray
                            Start-Sleep -Seconds 15
                            $item.Status = (Get-PnPListItem -List $PFS -Id $item.ID).FieldValues.VD_DocumentStatus

                            if ($item.Status -eq "Deleted") { Remove-ListItem -List $PFS -Id $item.ID }
                        }
                    }
                    else { Write-Host "[WARNING] - List: $($PFS) - ID: $($item.ID) - NOT PLACEHOLDER" -ForegroundColor Yellow }
                }
            }
        }
        Write-Host ''
    }
    if ($csv.Count -gt 1) { Write-Progress -Activity "Pulizia" -Completed }
    Write-Host 'Pulizia completata.' -ForegroundColor Cyan
}
catch { Throw }