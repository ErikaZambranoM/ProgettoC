# Questo script aggiunge un permission group a un sito

function Write-Log
{
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
    )

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else { Write-Host $Message -ForegroundColor Cyan }
}

try
{
    $GroupName = Read-Host -Prompt 'New Permission Group name'

    # Caricamento CSV / Sito
    $CSVPath = Read-Host -Prompt 'CSV Path o Site URL'
    if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
    else
    {
        $csv = [PSCustomObject] @{
            SiteURL = $CSVPath
            Count   = 1
        }
    }

    $rowCounter = 0
    Write-Log 'Inizio operazione...'
    ForEach ($row in $csv)
    {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($row.SiteURL)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }
        Connect-PnPOnline -Url $row.SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
        Write-Log "Sito: $($row.SiteURL.Split('/')[-1])"
        $Groups = Get-PnPGroup

        $found = $Groups | Where-Object -FilterScript { $_.Title -eq $GroupName }

        if ($null -ne $found) { Write-Log "[WARNING] - Permission Group: $($GroupName) - ALREADY EXISTS" }
        else
        {
            $SiteOwnerGroup = Get-PnPGroup -AssociatedOwnerGroup
            try
            {
                New-PnPGroup -Title $GroupName -Owner $SiteOwnerGroup.Title -WarningAction Stop -ErrorAction Stop | Out-Null
                Write-Log "[SUCCESS] - Permission Group: $($GroupName) - CREATED"
            }
            catch { Write-Log "[ERROR] - Permission Group: $($GroupName) - FAILED - $($_)" }
        }
    }
    Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }