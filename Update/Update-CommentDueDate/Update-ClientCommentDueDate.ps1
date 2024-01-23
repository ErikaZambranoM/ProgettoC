<#
    Questo script consente di aggiornare la Comment Due Date su Client Document List, Review Task Panel e cartella del documento

    ToDo:
        - Add start and end time to the output
        - Add csv log
#>


param(
    [parameter(Mandatory = $true)][String]$SiteUrl,
    [parameter(Mandatory = $true)][Array]$TransmittalNumbers,
    [parameter(Mandatory = $true)][int]$DaysToAdd
)

# Funzione di log to CSVfunction Write-Log
{
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $siteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID/Doc;Action;Key;Value;OldValue'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Previous: ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}


try {
    $siteUrl = $siteUrl.TrimEnd('/')
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    # Retrieve items from the "Client Document List" where "Last Transmittal" equals $TransmittalNumber
    Write-Log "Getting 'Client Document List'..."
    $listItems = Get-PnPListItem -List 'Client Document List' -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            Rev        = $_['IssueIndex']
            ComDueDate = $_['CommentDueDate']
            TRN        = $_['LastTransmittal']
            TRNDate    = $_['LastTransmittalDate']
            Path       = $_['DocumentsPath']
        }
    }
    Write-Log 'List loaded.'
    $listFiltered = $listItems | Where-Object -FilterScript { $_.TRN -in $TransmittalNumbers } | Sort-Object -Property TRN

    Write-Log "Getting 'Review Task Panel'..."
    $listItemsReview = Get-PnPListItem -List 'Review Task Panel' -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID                   = $_['ID']
            IDClientDocumentList = $_['IDClientDocumentList']
            ComDueDate           = $_['CommentDueDate']
        }
    }
    Write-Log "List loaded.`n"

    foreach ($item in $listFiltered) {
        Write-Progress -Activity 'Updating Comment Due Date...' -Status "Item: $($listFiltered.IndexOf($item) + 1)/$($listFiltered.Count)" -PercentComplete (($listFiltered.IndexOf($item) + 1) / $listFiltered.Count * 100)

        Write-Log "Processing document $($item.TCM_DN) - Transmittal Number: $($item.TRN)"

        # Calculate the new "Comment Due Date" by adding $DaysToAdd to the "Transmittal Date"
        $newCommentDueDate = $item.TRNDate.AddDays($DaysToAdd)

        # Update the item's "Comment Due Date"
        Set-PnPListItem -List 'Client Document List' -Identity $item.Id -Values @{'CommentDueDate' = $newCommentDueDate } | Out-Null

        Write-Log "[SUCCESS] Updated Comment Due Date for Item ID $($item.Id) - New Comment Due Date: $newCommentDueDate on Client Document List"

        $relativePath = $item.Path -replace $SiteUrl, ''

        $FilesToUpdate = Get-PnPFolderItem -FolderSiteRelativeUrl $relativePath -Recursive -ItemType File | ForEach-Object {

            Get-PnPFile -AsListItem -Url $_.ServerRelativeUrl | Where-Object -FilterScript { $_.FieldValues.TransmittalDate -ne $null }

        }

        $FilesToUpdate | Set-PnPListItem -Values @{'CommentDueDate' = $newCommentDueDate } | Out-Null
        Write-Log "[SUCCESS] Updated Comment Due Date for Item ID $($item.Id) - New Comment Due Date: $newCommentDueDate in the folder of the document"

        $listReviewFiltered = $listItemsReview | Where-Object -FilterScript { $_.IDClientDocumentList -eq $item.ID }
        foreach ($elemento in $listReviewFiltered) {
            Set-PnPListItem -List 'Review Task Panel' -Identity $elemento.ID -Values @{'CommentDueDate' = $newCommentDueDate } | Out-Null
            Write-Log "[SUCCESS] Updated Comment Due Date for Item ID $($item.Id) - New Comment Due Date: $newCommentDueDate on Review Task Panel"
        }
        Write-Host ''
    }
}
Catch {
    Throw
}
Finally {
    Write-Progress -Activity 'Updating Comment Due Date...' -Completed
    Write-Log 'Script completed.'
}