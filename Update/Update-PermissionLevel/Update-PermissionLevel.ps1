<#
    Questo script consente di aggiornare i permessi su una lista.
    Per funzionare necessita di un file CSV con i seguenti campi:
        - SiteURL: URL del Sito
        - List: Nome della lista
        - Group: Nome del gruppo
        - Level: Permission Level
#>

# Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message -ForegroundColor Cyan }
}

Function Connect-SPOSite {
    Param(
        # SharePoint Online Site URL
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Url
    )

    Try {
        # Create SPOConnection to specified Site if not already established
        $SPOConnection = ($Global:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $Url }).Connection
        If (-not $SPOConnection) {
            # Create SPOConnection to SiteURL provided in the PATicket
            $SPOConnection = Connect-PnPOnline -Url $Url -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

            # Add SPOConnection to the list of connections
            $Global:SPOConnections += [PSCustomObject]@{
                SiteUrl    = $Url
                Connection = $SPOConnection
            }
        }

        Return $SPOConnection
    }
    Catch {
        Throw
    }
}

Try {
    # Caricamento CSV / Sito
    $CSVPath = (Read-Host -Prompt 'CSV Path o Site URL').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
    else {
        $listName = Read-Host -Prompt 'List Name'
        $groupName = Read-Host -Prompt 'Group Name'
        $permLevel = Read-Host -Prompt 'Permission Level'
        $csv = [PSCustomObject] @{
            SiteURL = $CSVPath
            List    = $ListName
            Group   = $GroupName
            Level   = $permLevel
            Count   = 1
        }
    }

    $rowCounter = 0
    Write-Log 'Inizio aggiornamento...'
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count) - $($row.SiteURL)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }
        $SPOConnection = Connect-SPOSite -Url $row.SiteURL
        Write-Log "Sito: $($row.SiteURL.Split('/')[-1])"

        $listGroup = Get-PnPGroup -Identity $row.Group -Connection $SPOConnection

        if ($row.List -eq '') {
            try {
                $pLevels = Get-PnPGroupPermissions -Identity $listGroup -Connection $SPOConnection | Where-Object -FilterScript { $_.RoleTypeKind -ne 'Guest' }
                ForEach ($pLevel in $pLevels) {
                    try {
                        Set-PnPGroupPermissions -Identity $listGroup -RemoveRole $pLevel.Name -Connection $SPOConnection | Out-Null
                    }
                    catch { Write-Log "[ERROR] - Group: $($row.Group) - Level: $($pLevel.Name) - FAILED - $($_)" }
                }
            }
            catch { Write-Log "[WARNING] - Group: $($row.Group) - NEW PERMISSION" }

            try {
                Set-PnPGroupPermissions -Identity $listGroup -AddRole $row.Level -Connection $SPOConnection | Out-Null
                Write-Log "[SUCCESS] - Group: $($row.Group) - Permission: $($row.Level) - UPDATED"
            }
            catch {
                Write-Log "[ERROR] - Group: $($row.Group) - Permission: $($row.Level) - FAILED - $($_)"
                Exit
            }
        }
        else {
            try {
                $pLevels = Get-PnPListPermissions -Identity $row.List -PrincipalId $listGroup.Id -Connection $SPOConnection | Where-Object -FilterScript { $_.RoleTypeKind -ne 'Guest' }
                ForEach ($pLevel in $pLevels) {
                    try {
                        Set-PnPListPermission -Identity $row.List -Group $row.Group -RemoveRole $pLevel.Name -Connection $SPOConnection | Out-Null
                    }
                    catch {
                        Set-PnPList -Identity $row.List -BreakRoleInheritance -CopyRoleAssignments -Connection $SPOConnection | Out-Null
                        Write-Log "[WARNING] - $($row.List) - INHERITANCE DISABLED"
                        Set-PnPListPermission -Identity $row.List -Group $row.Group -RemoveRole $pLevel.Name -Connection $SPOConnection | Out-Null
                    }
                }
            }
            catch { Write-Log "[WARNING] - List: $($row.List) - Group: $($row.Group) - NEW PERMISSION" }

            try {
                Set-PnPListPermission -Identity $row.List -Group $row.Group -AddRole $row.Level -Connection $SPOConnection | Out-Null
                Write-Log "[SUCCESS] - List: $($row.List) - Group: $($row.Group) - Permission: $($row.Level) - UPDATED"
            }
            catch {
                Write-Log "[ERROR] - List: $($row.List) - Group: $($row.Group) - Permission: $($row.Level) - FAILED"
                Exit
            }
        }
    }
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed }
    Write-Log 'Operazione completata.'
}
Catch {
    Throw
}
