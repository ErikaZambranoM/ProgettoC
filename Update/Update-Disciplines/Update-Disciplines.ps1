
# PowerShell 7 Script
# Required Modules: PnP.PowerShell
# Import required modules


Param(
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint Site URL")]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Path where to save CSV file")]
    [string]$CsvPath

)

Import-Module PnP.PowerShell
# Importa CSV da file locale
$ContenutoCsvPath = Import-Csv -Path $CsvPath -Delimiter ";"
Function Export-SCListaUtentiToCSV {

    Try {

        # Connessione al sito
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue

        $ContenutoCsv = $ContenutoCsvPath | Where-Object -FilterScript { $_.Discipline } | Group-Object -Property "Discipline" | ForEach-Object {
            [PSCustomObject]@{
                Discipline = $_.Name
                UserEmail  = $_.Group."User E-mail".ToLower()
                #UserEmail  = ($_.Group."User E-mail" -join ";")
            }
        }

        # Custom Object Lista Disciplines SharePoint
        $listaSharePoint = Get-PnPListItem -List "Disciplines" -PageSize 5000 | ForEach-Object {
            [System.Collections.Generic.List[System.String]]$userEmails = $_["VD_PersonID"].Email | Foreach-Object { If ($null -ne $_ -and $_ -ne '') { $_.ToLower() } }
            [PSCustomObject]@{
                ID         = $_["ID"]
                Discipline = $_["Title"]
                UserEmail  = $userEmails
                #UserEmail  = ($_["VD_PersonID"].Email -join ";")
            }
        }


        foreach ($discipline in $listaSharePoint) {
            [System.Collections.Generic.List[System.String]]$mail = $null
            [Array]$DisciplineCSV = $ContenutoCsv | Where-Object -FilterScript { $_.Discipline -eq $discipline.Discipline }
            if (-not $DisciplineCSV) {
                Write-Host "Discipline $($Discipline.Discipline) not found in the csv" -ForegroundColor Yellow
                Continue
            }
            [Array]$mailCsv = ($ContenutoCsv | Where-Object -FilterScript { $_.Discipline -eq $discipline.Discipline } | Select-Object -Property UserEmail).UserEmail
            $mail = $discipline.UserEmail
            $compare = Compare-Object -ReferenceObject $mailCsv -DifferenceObject $mail | Where-Object -FilterScript { $_.SideIndicator -eq "<=" }     #restituisce solo le mail non presenti

            ForEach ( $item in $compare.InputObject) {
                $mail.Add($item) | Out-Null #aggiorna l'array con le mail mancanti
            }
            try {
                Set-PnPListItem -List "Disciplines" -Identity $discipline.ID -Values @{ "VD_PersonID" = [Array]$mail } | Out-Null
                Write-Host "Set the discipline $($Discipline.discipline) with the follow mails:`n" -BackgroundColor White
                Write-Host $($mail -join "`n") -ForegroundColor Green

            }
            catch {
                Write-Host "Failed to add following user to Discipline $($Discipline.Discipline):`n$($mailCsv -join "`n")" -ForegroundColor Red
                Continue
            }

        }

    }
    Catch {
        Throw
    }
}

# Invoke the function to Update the mail of the disciplines List
Export-SCListaUtentiToCSV

# C:\Users\erika.zambrano_grupp\Downloads\Disciplines.csv
# path di Test https://tecnimont.sharepoint.com/sites/poc_vdm