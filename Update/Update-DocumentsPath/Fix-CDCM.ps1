param(
    [parameter(Mandatory = $true)][string]$ProjectCode
)

#Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $siteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID;TCM_DN;Rev;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) {
        Write-Host $Message -ForegroundColor Green
    }
    elseif ($Message.Contains('[ERROR]')) {
        Write-Host $Message -ForegroundColor Red
    }
    elseif ($Message.Contains('[WARNING]')) {
        Write-Host $Message -ForegroundColor Yellow
    }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Doc: ', ';').Replace(' - ID: ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

# Funzione che carica la Client Department Code Mapping
function Get-CDCM {
    param (
        [Parameter(Mandatory = $true)]$SiteConn
    )

    Write-Log "Caricamento '$($CDCM)'..."
    $items = Get-PnPListItem -List $CDCM -PageSize 5000 -Connection $SiteConn | ForEach-Object {
        [PSCustomObject] @{
            ID       = $_['ID']
            Title    = $_['Title']
            Value    = $_['Value']
            ListPath = $_['ListPath']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Return $items
}
# Funzione che verifica la correttezza del Client Discipline
function Find-ClientDiscipline {
    param (
        [Parameter(Mandatory = $true)][String]$ClientCode,
        [Parameter(Mandatory = $true)][Array]$List
    )

    $ccSplit = $ClientCode.Split('-')
    if ($ClientCode.toUpper().StartsWith('DS')) {
        if ($ccSplit.Length -ge 4) {
            $tempCode = $ccSplit[2] + '-' + $ccSplit[3]
            [Array]$found = $List | Where-Object -FilterScript { $_.Title -eq $tempCode }
        }
    }
    else {
        if ($ccSplit.Length -ge 2) {
            $tempCode = $ccSplit[1]
            if ($tempCode -eq 'CR') { [Array]$found = $List | Where-Object -FilterScript { $_.ListPath -eq 'CSR' } }
            else { [Array]$found = $List | Where-Object -FilterScript { $_.ListPath -eq $tempCode } }
        }
    }
    if ($found.Count -eq 0) { [Array]$found = $List | Where-Object -FilterScript { $_.ListPath -eq 'NA' } }
    return $found[0]
}

$tcmUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
$CDCM = 'ClientDepartmentCodeMapping'

$listName = 'DocumentList'

$tcmConn = Connect-PnPOnline -Url $tcmUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
$mainConn = $tcmConn

$siteCode = (Get-PnPWeb -Connection $mainConn).Title.Split(' ')[0]

# Legge tutta la document library
Write-Log "Caricamento '$($listName)'..."
$listItems = Get-PnPListItem -List $listName -PageSize 5000 -Connection $mainConn | ForEach-Object {
    [PSCustomObject]@{
        ID               = $_['ID']
        TCM_DN           = $_['Title']
        ClientCode       = $_['ClientCode']
        Rev              = $_['IssueIndex']
        Status           = $_['DocumentStatus']
        ClientDiscipline = $_['ClientDiscipline_Calculated']
        DocPath          = $_['DocumentsPath']
    }
}
Write-Log 'Caricamento lista completato.'

# Caricamento Client Department Code Mapping
$CDCMItems = Get-CDCM -SiteConn $tcmConn

$csv = $listItems | Where-Object -FilterScript { $_.ClientCode -ne $null }

$rowCounter = 0
Write-Log 'Inizio controllo...'
ForEach ($row in $csv) {
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Verifica' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)/$($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

    Write-Host "Documento: $($row.TCM_DN)/$($row.Rev)" -ForegroundColor Blue
    $ccSplit = $row.ClientCode.Split('-')

    $clientDiscipline_Calc = Find-ClientDiscipline -ClientCode $row.ClientCode -List $CDCMItems

    if ($row.ClientDiscipline -ne $clientDiscipline_Calc.Value) {
        Write-Host $row.ClientDiscipline
        Write-Host "$($clientDiscipline_Calc.Value) | $($clientDiscipline_Calc.Title) | $($clientDiscipline_Calc.ID)"
        try {
            Set-PnPListItem -List $listName -Identity $row.ID -Values @{
                ClientDiscipline_Calculated = $clientDiscipline_Calc.Value
            } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
            Write-Log "[SUCCESS] - List: DocumentList - Doc: $($row.TCM_DN)/$($row.Rev) - UPDATED"
        }
        catch { Write-Log "[ERROR] - List: DocumentList - Doc: $($row.TCM_DN)/$($row.Rev) - FAILED" }
    }
}
Write-Progress -Activity 'Verifica' -Completed
Write-Log 'Operazione completata.'