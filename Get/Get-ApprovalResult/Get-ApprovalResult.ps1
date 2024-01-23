param(
    [parameter(Mandatory = $true)][string]$SiteUrl, #URL del sito
    [switch]$CountOnly #Only returns number of items on list without producing export
)

if ($SiteUrl.ToLower().Contains('digitaldocumentsc')) {
    $listType = 'CD'
    $listName = 'Client Document List'
}
elseif ($SiteUrl.ToLower().Contains('digitaldocuments')) {
    $listType = 'DD'
    $listName = 'DocumentList'
}
elseif ($SiteUrl.ToLower().Contains('vdm_')) {
    $listType = 'VD'
    $listName = 'Vendor Documents List'
}

$CSVPath = Read-Host -Prompt 'CSV Path o TCM Document Number'
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv $CSVPath -Delimiter ';' }
else {
    $rev = Read-Host -Prompt 'Issue Index'
    $csv = New-Object [PSCustomObject]@ {
        TCM_DN = $CSVPath
        Rev = $rev
        Count = 1
    }
}

# Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop

$siteCode = (Get-PnPWeb).Title.Split(' ')[0]
$ExecutionDate = Get-Date -Format 'yyyy_MM_dd-HH'
$filePath = "$($PSScriptRoot)\export\$($siteCode)-$($listType)-$($ExecutionDate).csv";

# Get all list items
Write-Host "Caricamento '$($listName)'..." -ForegroundColor Cyan
$listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
    if ($listType -eq 'DD' -or $listType -eq 'CD') {
        $item = New-Object -TypeName PSCustomObject -Property @{
            ID                = $_['ID']
            TCM_DN            = $_['Title']
            Rev               = $_['IssueIndex']
            IsCurrent         = $_['IsCurrent']
            CommentRequest    = $_['CommentRequest']
            ApprovalResult    = $_['ApprovalResult']
            PMApprovalRequest = $_['PMApprovalRequest']
            PMApprovalAction  = $_['PMApprovalAction']
        }
    }
    elseif ($listType -eq 'VD') {
        $item = New-Object -TypeName PSCustomObject -Property @{
            ID         = $($_['ID'])
            TCM_DN     = $($_['VD_DocumentNumber'])
            Rev        = $($_['VD_RevisionNumber'])
            ClientCode = $($_['VD_ClientDocumentNumber'])
        }
    }
    $item
}
Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

ForEach ($row in $csv) {
    $item = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

    if (!(Test-Path -Path $filePath)) { New-Item $filePath -Force -ItemType File | Out-Null }
    $item | Export-Csv -Path $filePath -Delimiter ';' -NoTypeInformation -Append
}
Write-Host "[SUCCESS] Export generato nel percorso $($filePath)" -ForegroundColor Green