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
    $RTA = 'Review Task Archive'

    Connect-PnPOnline -Url $siteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    $userList = Get-PnPUser

    Write-Log "Caricamento '$($CDL)'..."
    $CDLItems = Get-PnPListItem -List $CDL -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            Rev        = $_['IssueIndex']
            ClientCode = $_['ClientCode']
            IsCurrent  = $_['IsCurrent']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Write-Log "Caricamento '$($RTP)'..."
    $RTPItems = Get-PnPListItem -List $RTP -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID                   = $_['ID']
            Title                = $_['Title']
            ClientCode           = $_['ClientCode']
            ReasonForIssue       = $_['ReasonForIssue']
            IssueIndex           = $_['IssueIndex']
            DocumentTitle        = $_['DocumentTitle']
            ClientDiscipline     = $_['ClientDiscipline']
            IDClientDocumentList = $_['IDClientDocumentList']
            CommentDueDate       = $_['CommentDueDate']
            PONumber             = $_['PONumber']
            ActionReview         = $_['ActionReview']
            ClientCommentStatus  = $_['ClientCommentStatus']
            DocumentClass        = $_['DocumentClass']
            TaskPriority         = $_['TaskPriority']
            Assignee             = [Array]$_['Assignee'].Email
            Consolidator         = [Array]$_['Consolidator'].Email
            StatusChangeDate     = $_['StatusChangeDate']
            StatusChangeUser     = $_['StatusChangeUser'].Email
        }
    }
    Write-Log 'Caricamento lista completato.'

    Write-Log 'Inizio pulizia...'
    $rowCounter = 0
    ForEach ($item in $RTPItems) {
        if ($RTPItems.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Status "$($rowCounter+1)/$($RTPItems.Count)" -PercentComplete (($rowCounter++ / $RTPItems.Count) * 100) }

        $foundCDL = $CDLItems | Where-Object -FilterScript { $_.ID -eq $item.IDClientDocumentList }

        if ($null -eq $foundCDL) { 
            try {
                Remove-PnPListItem -List $RTP -Identity $item.ID -Force | Out-Null
                Write-Log "[WARNING] - List: $($RTP) - Doc: $($item.ClientCode)/$($item.IssueIndex) - DELETED (NOT FOUND IN CDL)"
            }
            catch { Write-Log "[ERROR] - List: $($RTP) - Doc: $($item.ClientCode)/$($item.IssueIndex) - $($_)" }
        }
        elseif ($foundCDL.Length -gt 1) { Write-Log "[WARNING] - List: $($CDL) - Doc: $($item.ClientCode)/$($item.IssueIndex) - DUPLICATED" }
        elseif ($foundCDL.IsCurrent) { Continue }
        else {
            Write-Host "Doc: $($foundCDL.ClientCode)/$($foundCDL.Rev) - ID: $($foundCDL.ID)" -ForegroundColor Blue

            $usersToConvert = @()
            ForEach ($mail in $item.Assignee) { $usersToConvert+= $userList | Where-Object -FilterScript { $_.Email -eq $mail } }
            $item.Assignee = $usersToConvert.LoginName

            $usersToConvert = @()
            ForEach ($mail in $item.Consolidator) { $usersToConvert+= $userList | Where-Object -FilterScript { $_.Email -eq $mail } }
            $item.Consolidator = $usersToConvert.LoginName

            # Crea HashTable da passare all'RTA
            $hashTable = @{}
            $item.psobject.Properties | ForEach-Object { 
                if ($_.Value -ne $null) {
                    $hashTable[$_.Name] = $_.Value
                }
            }
            $hashTable.Remove('ID')

            # Aggiunge HashTable a RTA
            try {
                Add-PnPListItem -List $RTA -Values $hashTable | Out-Null
                Write-Log "[SUCCESS] - List: $($RTA) - Doc: $($item.ClientCode)/$($item.IssueIndex) - ADDED"
            }
            catch { Write-Log "[ERROR] - List: $($RTA) - Doc: $($item.ClientCode)/$($item.IssueIndex) - $($_)" }

            # Rimuove l'item da RTP
            try {
                Remove-PnPListItem -List $RTP -Identity $item.ID -Force -Recycle | Out-Null
                Write-Log "[SUCCESS] - List: $($RTP) - Doc: $($item.ClientCode)/$($item.IssueIndex) - DELETED"
            }
            catch { Write-Log "[ERROR] - List: $($RTP) - Doc: $($item.ClientCode)/$($item.IssueIndex) - $($_)" }
        }
    }
}
catch { Throw }
finally { if ($RTPItems.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Completed } }