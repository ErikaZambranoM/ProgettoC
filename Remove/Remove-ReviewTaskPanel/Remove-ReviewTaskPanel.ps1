#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param (
    [Parameter(Mandatory = $true)][String]$ProjectCode
)

#Funzione di log to CSV
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $ProjectCode
    )

    $Path = "$($PSScriptRoot)\logs\$($Code)-$(Get-Date -Format 'yyyy_MM_dd').csv";

    if (!(Test-Path -Path $Path)) {
        $newLog = New-Item $Path -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;TCM_DN;Rev;Action;Value'
    }

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) {	Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }

    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Message = $Message.Replace(' - List: ', ';').Replace(' - TCM_DN: ', ';').Replace(' - Rev: ', ';').Replace(' - ID: ', ';').Replace(' - Folder: ', ';')
    Add-Content $Path "$FormattedDate;$Message"
}

try {
    $siteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
    $CDL = 'Client Document List'
    $RTP = 'Review Task Panel'

    $CSVPath = (Read-Host -Prompt 'CSV Path o ClientCode').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
    else {
        $Rev = Read-Host -Prompt 'Issue Index'
        $csv = [PSCustomObject] @{
            ClientCode = $CSVPath
            Rev        = $Rev
            Count      = 1
        }
    }

    Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    Write-Log "Caricamento '$($CDL)'..."
    $CDLItems = Get-PnPListItem -List $CDL -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            Rev        = $_['IssueIndex']
            ClientCode = $_['ClientCode']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Write-Log "Caricamento '$($RTP)'..."
    $RPTItems = Get-PnPListItem -List $RTP -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            ClientCode = $_['ClientCode']
            Rev        = $_['IssueIndex']
            Assignee   = $_['Assignee'].LookupValue
            CDL_ID     = $_['IDClientDocumentList']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Write-Log 'Inizio pulizia...'
    $rowCounter = 0
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

        $foundCDL = $CDLItems | Where-Object -FilterScript { $_.ClientCode -eq $row.ClientCode -and $_.Rev -eq $row.Rev }

        if ($null -eq $foundCDL) { Write-Log "[ERROR] - List: $($CDL) - Doc: $($row.ClientCode)/$($row.Rev) - NOT FOUND" }
        elseif ($foundCDL.Length -gt 1) { Write-Log "[WARNING] - List: $($CDL) - Doc: $($row.ClientCode)/$($row.Rev) - DUPLICATED" }
        else {
            Write-Host "Doc: $($foundCDL.ClientCode)/$($foundCDL.Rev) - ID: $($foundCDL.ID)" -ForegroundColor Blue

            [Array]$foundRTP = $RPTItems | Where-Object -FilterScript { $_.ClientCode -eq $row.ClientCode -and $_.Rev -eq $row.Rev }

            if ($null -eq $foundRTP) { Write-Log "[ERROR] - List: $($RTP) - Doc: $($row.ClientCode)/$($row.Rev) - NOT FOUND" }
            else {
                ForEach ($item in $foundRTP) {
                    if ($item.CDL_ID -ne $foundCDL.ID) {
                        try {
                            Remove-PnPListItem -List $RTP -Identity $item.ID -Recycle -Force | Out-Null
                            Write-Log "[SUCCESS] - List: $($RTP) - Assignee: $($item.Assignee) - DELETED ID: $($item.CDL_ID)"
                        }
                        catch { Write-Log "[ERROR] - List: $($RTP) - Assignee: $($item.Assignee) - $($_)" }
                    }
                }
            }
        }
    }
    Write-Log 'Pulizia completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Completed } }