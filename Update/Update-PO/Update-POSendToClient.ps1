<#Questo Scipt va a settare il VD_SendToClient a true nel caso non sia stato valorizzato
Nel caso sia necessario un controllo sul documento si può aggiungere alla riga 51
#>

Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$SiteCode,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the PO')]
    [string]$PO
)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $codiceSito
    )
    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";
    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;Level;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Document: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}
$sito = "https://tecnimont.sharepoint.com/sites/vdm_$($SiteCode)"
$SPOConnection = Connect-PnPOnline -Url $sito -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction Continue
$listName = "Vendor Documents List"
Write-Log "Caricamento lista $($listName)"
$PODocuments = Get-PnPListItem -List $listName -Connection $SPOConnection -PageSize 5000 | Where-Object -FilterScript { $_['VD_PONumber'] -eq $($PO) }
$PODocumentsToModify = $PODocuments | ForEach-Object {
    [PSCustomObject]@{
        ID              = $_['ID']
        Name            = $_['VD_DocumentNumber']
        VD_SendToClient = $_['VD_SendToClient']
    }
}
Write-Log "Lista $($listName) caricata"

try {
    foreach ($item in $PODocumentsToModify) {
        Set-PnPListItem -List $listName -Identity $item.ID -Values @{ VD_SendToClient = "True" } -UpdateType SystemUpdate | Out-Null
        Write-Log "[WARNING] - List: $($listName) - Document: $($item.Name) - UPDATED Vendor Group"
    }
    break
}
catch {
    throw
}