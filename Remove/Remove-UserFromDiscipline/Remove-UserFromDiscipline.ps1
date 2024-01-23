param(
    [Parameter(Mandatory = $true)][String]$SiteURL,
    [Parameter(Mandatory = $true)][String]$UserEmailToRemove,
    [Switch]$Replace
)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    if ($Message.Contains("[SUCCESS]")) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains("[ERROR]")) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains("[WARNING]")) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message -ForegroundColor Cyan }
}

$listName = "Disciplines"
$fieldName = "VD_PersonID"
$UserEmailToRemove = $UserEmailToRemove.ToLower()

if ($Replace) { $userEmailToAdd = (Read-Host -Prompt "UserEmailToAdd").ToLower() }

Connect-PnPOnline -Url $SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

# Caricamento lista
Write-Log "Caricamento '$($listName)'..."
$listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
    [System.Collections.Generic.List[System.String]]$userEmails = $_[$fieldName].Email | Foreach-Object { If ($null -ne $_ -and $_ -ne '') { $_.ToLower() } }
    [PSCustomObject]@{
        ID     = $_["ID"]
        Name   = $_["Title"]
        People = $userEmails
    }
}
Write-Log "Caricamento completato."

# Filtro elementi della lista contenenti la mail
$userDisciplines = $listItems | Where-Object { $_.People -contains $UserEmailToRemove }

# Controllo presenza
If ( $userDisciplines.Count -eq 0 ) { Write-Log "[WARNING] - List: $($listName) - User: $($UserEmailToRemove) - NOT FOUND" }
Else
{
    $rowCounter = 0
    Write-Log "Inizio pulizia..."
    ForEach ($discipline in $userDisciplines) {
        if ($userDisciplines.Length -gt 1) { Write-Progress -Activity "Aggiornamento" -Status "$($rowCounter+1)/$($userDisciplines.Length)" -PercentComplete (($rowCounter++ / $userDisciplines.Length) * 100) }
        try {
            # Rimuove la mail della colonna
            $discipline.People.Remove($UserEmailToRemove) | Out-Null
            
            # Se $Replace, aggiunge la nuova mail alla colonna
            if ($Replace) { $discipline.People.Add($userEmailToAdd) | Out-Null }

            # Aggiorna la lista
            Set-PnPListItem -List $listName -Identity $discipline.ID -Values @{ $fieldName = [Array]$discipline.People } -UpdateType SystemUpdate | Out-Null
            $msg = "[SUCCESS] - List: $($listName) - $($discipline.Name) - User: $($UserEmailToRemove) - UPDATED"
        }
        catch { $msg = "[ERROR] - List: $($listName) - $($discipline.Name) - User: $($UserEmailToRemove) - FAILED - $($_)" }
        Write-Log $msg
    }
    if ($userDisciplines.Length -gt 1) { Write-Progress -Activity "Aggiornamento" -Completed }
    Write-Log "Pulizia completata."
}