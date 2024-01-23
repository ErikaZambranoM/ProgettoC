<#
    Questo script concede i permessi a tutti gli utenti inseriti del campo Consolidator sul Review Task Panel.

    AGGIORNARE IL FILTRO PRIMA DI ESEGUIRE.
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param (
    [Parameter(Mandatory = $true)][String]$ProjectCode
)

try {
    $siteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
    $listName = 'Review Task Panel'
    $permEdit = 'MT Contributors'
    #$permRead = 'MT Readers'

    $change = Read-Host -Prompt "Cambio Consolidator (true/false)"
    if ([System.Convert]::ToBoolean($change)) {
        $newConsolidator = Read-Host -Prompt "Mail"
    }

    Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    Write-Host "Caricamento '$($listName)'..." -ForegroundColor Cyan
    $listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
        [System.Collections.Generic.List[System.String]]$userEmails = $_['Consolidator'].Email | ForEach-Object { If ($null -ne $_ -and $_ -ne '') { $_.ToLower() } }
        [PSCustomObject]@{
            ID           = $_['ID']
            ClientCode   = $_['ClientCode']
            Rev          = $_['IssueIndex']
            CDL_ID       = $_['IDClientDocumentList']
            ActionReview = $_['ActionReview']
            Consolidator = $userEmails
        }
    }
    Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

    ##### Filtro RTP #####

    $filtered = $listItems | Where-Object -FilterScript { $_.ActionReview -eq 'Consolidate' -and $_.Consolidator.Count -ge 1 }
    
    ######################

    $itemCounter = 0
    Write-Host 'Inizio aggiornamento...' -ForegroundColor Cyan
    foreach ($item in $filtered) {
        if ($filtered.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($itemCounter+1)/$($filtered.Count)" -PercentComplete (($itemCounter++ / $filtered.Count) * 100) }

        Write-Host "Doc: $($item.ClientCode)/$($item.Rev) - ID: $($item.CDL_ID)" -ForegroundColor Blue

        # Cambio campo Consolidator
        if ([System.Convert]::ToBoolean($change)) {
            try {
                Set-PnPListItem -List $listName -Identity $item.ID -Values @{
                    Consolidator = $newConsolidator
                } -UpdateType SystemUpdate | Out-Null
                Write-Host "[SUCCESS] - List: $($listName) - ID: $($item.ID) - UPDATED" -ForegroundColor Green
            }
            catch { Write-Log "[ERROR] - List: $($listName) - ID: $($item.ID) - $($_)" -ForegroundColor Red }
        }

        # Assegna i permessi con per ogni Consolidator
        foreach ($mail in $item.Consolidator) {
            try {
                Set-PnPListItemPermission -List $listName -Identity $item.ID -User $mail -AddRole $permEdit | Out-Null
                Write-Host "[SUCCESS] - List: $($listName) - ID: $($item.ID) - User: $($mail) - PERMISSION UPDATED" -ForegroundColor Green
            }
            catch { Write-Log "[ERROR] - List: $($listName) - ID: $($item.ID) - User: $($mail) - $($_)" -ForegroundColor Red }
        }
    }
    Write-Host 'Aggiornamento completato.' -ForegroundColor Cyan
}
catch { Throw }
finally { if ($filtered.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }