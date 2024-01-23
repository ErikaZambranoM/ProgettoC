param (
    [parameter(Mandatory = $true)][String]$ProjectCode, # Codice del progetto
    [parameter(Mandatory = $true)][String]$GroupName, # Nome del gruppo SharePoint
    [parameter(Mandatory = $true)][String]$Mail # Mail (1 o più) da aggiungere
)

$mainUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
$clientUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"

$mailArray = $Mail.Split(';')

$mainConn = Connect-PnPOnline -Url $mainUrl -UseWebLogin -ValidateConnection -ReturnConnection
$clientConn = Connect-PnPOnline -Url $clientUrl -UseWebLogin -ValidateConnection -ReturnConnection

$count = 0
Write-Host 'Inizio aggiornamento...' -ForegroundColor Cyan
ForEach ($item in $mailArray) {
    if ($mailArray.Length -gt 1) { Write-Progress -Activity 'Aggiunta' -Status "$($count+1)/$($mailArray.Length) - $($item)" -PercentComplete (($count++ / $mailArray.Length) * 100) }
    try {
        Add-PnPGroupMember -LoginName $item -Group $GroupName -Connection $mainConn
        Write-Host "[SUCCESS] - Group: $($GroupName) - Mail: $($item) - ADDED MAIN" -ForegroundColor Green
        Add-PnPGroupMember -LoginName $item -Group $GroupName -Connection $clientConn
        Write-Host "[SUCCESS] - Group: $($GroupName) - Mail: $($item) - ADDED CLIENT" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] - Group: $($GroupName) - Mail: $($item) - FAILED - $($_)" -ForegroundColor Red
        Exit
    }
}
Write-Progress -Activity 'Aggiunta' -Completed
Write-Host 'Aggiornamento completato.' -ForegroundColor Cyan
