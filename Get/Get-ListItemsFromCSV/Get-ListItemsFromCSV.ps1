# Created to export a CSV form List based on another source CSV
# Adapt List Fields accordingly
Param (
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [String]$CSVPath
)

$ListName = 'Client Document List'

# Import the CSV with the list of documents to be checked and exit if it is empty
$DocumentsToBeChecked = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8
If ($DocumentsToBeChecked.Count -eq 0) {
    Write-Host 'No documents found on CSV' -ForegroundColor Yellow
    Exit
}

Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection

# Create csv file to log the results
$LogFilePath = $CSVPath.Replace('.csv', "_$(Get-Date -Format 'dd-MM-yyyy_HH-mm-ss')_Log.csv")

# Get all item in Vendor Documents List
$AllListItems = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
    $Item = New-Object -TypeName PSCustomObject -Property @{
        ID             = $_['ID']
        TCM_DN         = $_['Title']
        Rev            = $_['IssueIndex']
        ApprovalResult = $_['ApprovalResult']
        #TCM_DN  = $_["VD_DocumentNumber"]
        #Rev     = $_["VD_RevisionNumber"]
        #DocPath = $_["VD_DocumentPath"]
    }
    $Item
}

# Loop through the list of documents to be checked and check if they are in the list
ForEach ($Document in $DocumentsToBeChecked) {
    Try {
        # Progress bar to show the progress of the script
        Write-Progress -Activity ("Checking documents on List '{0}' of Site {1}" -f $ListName, $SiteUrl) -Status "Checking document $($Document.TCM_DN) - Rev $($Document.Rev)" -PercentComplete (($DocumentsToBeChecked.IndexOf($Document) + 1) / $DocumentsToBeChecked.Count * 100)

        # Remove Rev from filterscript to search only for TCM_DN (also change count condition)
        [Array]$DocumentsOnList = $AllListItems | Where-Object -FilterScript { $_.TCM_DN -eq $Document.TCM_DN -and $_.Rev -eq $Document.Rev }
        If ($DocumentsOnList.Count -eq 1) {
            $DocumentsOnList | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Found' -Force
        }
        ElseIf ($DocumentsOnList.Count -gt 1) {
            $DocumentsOnList | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Duplicated' -Force
        }
        Else {
            $DocumentsOnList | Add-Member -MemberType NoteProperty -Name 'Missing' -Force
        }

        $DocumentsOnList | Export-Csv -Path $LogFilePath -NoTypeInformation -Append -Delimiter ';'
    }
    Catch {
        Write-Host "Error checking document $($Document.TCM_DN) - $($_)" -ForegroundColor Red
    }
}
Write-Progress -Activity "Checking documents on $ListName" -Completed
