<#
    CSV file must contain TCM_DN and, optionally, Rev columns

    ToDo:
        - Add parameter for single Document
        - Add parameter -Force to delete documents not in Placeholder status
#>

Param (
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [String]$CSVPath
)

# Import the CSV with the list of documents to be deleted and exit if it is empty
[array]$DocumentsTobeDeleted = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8
If ($DocumentsToBeDeleted.Count -eq 0) {
    Write-Host "`nNo documents found on CSV" -ForegroundColor Yellow
    Exit
}

# Create csv file to log the results
$Date = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$DeletationLogFilePath = $CSVPath.Replace('.csv', "_$($Date)_DeletationLog.csv")
$ProcessFlowCheckLogFilePath = $CSVPath.Replace('.csv', "_$($Date)_ProcessFlowCheckLog.csv")

# Connect to the site
$VDMSiteConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection

# Get all item in Vendor Documents List
Write-Host "`nLoading Vendor Documents List" -ForegroundColor Yellow
$AllVDLListItems = Get-PnPListItem -List 'Vendor Documents List' -PageSize 5000 -Connection $VDMSiteConnection | ForEach-Object {
    $Item = New-Object -TypeName PSCustomObject -Property @{
        ID     = $_['ID']
        TCM_DN = $_['VD_DocumentNumber']
        Rev    = $_['VD_RevisionNumber']
    }
    $Item
}

# Get all item in Process Flow Status List
Write-Host "`nLoading Process Flow Status List" -ForegroundColor Yellow
$AllProcessFlowListItems = Get-PnPListItem -List 'Process Flow Status List' -PageSize 5000 -Connection $VDMSiteConnection | ForEach-Object {
    $Item = New-Object -TypeName PSCustomObject -Property @{
        ID           = $_['ID']
        TCM_DN       = $_['VD_DocumentNumber']
        Rev          = $_['VD_RevisionNumber']
        Index        = $_['VD_Index']
        Status       = $_['VD_DocumentStatus']
        SubSitePOURL = $($_['VD_PONumberUrl'].Url)
    }
    $Item
}

$CSVColumns = ($DocumentsTobeDeleted | Get-Member -MemberType NoteProperty).Name
If ('Rev' -in $CSVColumns) {
    $filterScriptText = '(&{$_.TCM_DN -eq $Document.TCM_DN -and $_.Rev -eq $Document.Rev})'
    $filterScript = [scriptblock]::Create($filterScriptText)
}
Else {
    $filterScriptText = '(&{$_.TCM_DN -eq $Document.TCM_DN})'
    $filterScript = [scriptblock]::Create($filterScriptText)
}

# Loop through the list of documents to be deleted and delete them if they are in Placeholder status
$Progress = 0
ForEach ($Document in $DocumentsToBeDeleted) {
    # Progress bar to show the progress of the script
    $Progress = [int]($Progress + 1)
    $ProgressPercentage = [int](($Progress / $DocumentsToBeDeleted.Count) * 100)
    Write-Progress -Activity ('Checking and deleting documents {0}/{1}' -f $Progress, $DocumentsToBeDeleted.Count) -Status "Deleting document $($Document.TCM_DN)" -PercentComplete $ProgressPercentage

    Try {
        # Check if document is in Placeholder status
        [array]$ProcessFlowItem = $AllProcessFlowListItems | Where-Object -FilterScript $filterScript

        If ($ProcessFlowItem.Count -gt 1) {
            $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Document skipped, more then 1 revision found on Process Flow Status List' -Force
            $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
            Write-Host "`nDocument $($ProcessFlowItem.TCM_DN) skipped, more then 1 revision found Process Flow Status List" -ForegroundColor Yellow
            Continue
        }

        If ($ProcessFlowItem.Count -eq 0) {
            [array]$VDLListItem = $AllVDLListItems | Where-Object -FilterScript $filterScript
            If ($VDLListItem.Count -ge 1) {
                $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Document skipped, not found on Process Flow Status List but found $($VDLListItem.Count) occurrences on VDL" -Force
                $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
                Write-Host "`nDocument $($ProcessFlowItem.TCM_DN) skipped, not found on Process Flow Status List" -ForegroundColor Yellow
            }
            Else {
                $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Document skipped, not found on Process Flow Status List and VDL' -Force
                $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
                Write-Host "`nDocument $($ProcessFlowItem.TCM_DN) skipped, not found on Process Flow Status List and VDL" -ForegroundColor Yellow
            }
            Continue
        }

        If ('Rev' -in $CSVColumns -and $null -eq $Document.Rev) {
            $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Document skipped, no revision found on CSV' -Force
            $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
            Write-Host "`nDocument $($ProcessFlowItem.TCM_DN) skipped, no revision found on CSV" -ForegroundColor Yellow
            Continue
        }

        $SubSiteURL = $($ProcessFlowItem.SubSitePOURL.Substring(0, $ProcessFlowItem.SubSitePOURL.LastIndexOf('/')))

        # Delete the document if it is in Placeholder status, otherwise skip it
        If ($ProcessFlowItem.Status -eq 'Placeholder') {
            # Delete document from Vendor Documents List
            [array]$VDLListItem = $AllVDLListItems | Where-Object -FilterScript $filterScript
            If ($VDLListItem.Count -gt 1) {
                $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Document skipped, more then 1 revision found on VDL' -Force
                $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
                Write-Host "`nDocument $($VDLListItem.TCM_DN) skipped, more then 1 revision found on VDL" -ForegroundColor Yellow
                Continue
            }

            # Compare Process Flow Status List Item with Vendor Documents List Item
            If (($ProcessFlowItem.TCM_DN -ne $VDLListItem.TCM_DN) -or ($ProcessFlowItem.Rev -ne $VDLListItem.Rev) -and ($VDLListItem.Count -ne 0)) {
                $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Document skipped, Process Flow Status List Item and Vendor Documents List Item are different' -Force
                $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
                Write-Host "`nDocument $($VDLListItem.TCM_DN) skipped, Process Flow Status List Item and Vendor Documents List Item are different" -ForegroundColor Yellow
                Continue
            }

            # Connect to the sub site
            $VDMSubSiteConnection = Connect-PnPOnline -Url $SubSiteURL -UseWebLogin -ValidateConnection -ReturnConnection

            $VDFolderRev = "$($ProcessFlowItem.TCM_DN)-$($ProcessFlowItem.Index.ToString().PadLeft(3, '0'))"
            $VDDocumentPath = "$($ProcessFlowItem.SubSitePOURL)/$($VDFolderRev)"

            # Check if SharePoint Online folder exists for the document
            $VDDocFolder = Get-PnPFolder -Url $VDDocumentPath -ErrorAction SilentlyContinue -Connection $VDMSubSiteConnection

            # If only 1 item is found and it is in Placeholder status, delete it
            $DelProcessArray = @()
            If ($VDLListItem.Count -ne 0) {
                Remove-PnPListItem -List 'Vendor Documents List' -Identity $VDLListItem.ID -Recycle -Force -Connection $VDMSiteConnection | Out-Null
                Write-Host "`nDocument $($VDLListItem.TCM_DN) - Rev: $($VDLListItem.Rev) deleted from Vendor Documents List" -ForegroundColor Green
                $DelProcessArray += 'VDL'
            }

            # If folder does not exist or if Document is not in VDL, delete orphaned objeccts (Process Flow Status List Item, Revision Folder Dashboard Item and maybe Document Folder)
            If ($null -eq $VDDocFolder -or $VDLListItem.Count -eq 0) {
                If ($null -eq $VDDocFolder) {
                    Write-Host "`nDocument $($VDLListItem.TCM_DN) - Rev $($VDLListItem.Rev) Folder does not exists!" -ForegroundColor Yellow
                }

                # Set document status to Deleted on Process Flow Status List and then delete the item
                Set-PnPListItem -List 'Process Flow Status List' -Identity $ProcessFlowItem.ID -Values @{'VD_DocumentStatus' = 'Deleted' } -Force -Connection $VDMSiteConnection | Out-Null
                Remove-PnPListItem -List 'Process Flow Status List' -Identity $ProcessFlowItem.ID -Recycle -Force -Connection $VDMSiteConnection | Out-Null
                Write-Host "`nDocument $($ProcessFlowItem.TCM_DN) - Rev: $($ProcessFlowItem.Rev) set on status Deleted and then deleted" -ForegroundColor Yellow
                $DelProcessArray += 'ProcessFlow'

                # Get and delete the item from Revision Folder Dashboard
                $AllRevisionFolderDashboardListItems = Get-PnPListItem -List 'Revision Folder Dashboard' -PageSize 5000 -Connection $VDMSubSiteConnection | ForEach-Object {
                    $Item = New-Object -TypeName PSCustomObject -Property @{
                        ID     = $_['ID']
                        TCM_DN = $_['VD_DocumentNumber']
                        Rev    = $_['VD_RevisionNumber']
                        Status = $_['VD_DocumentSubmissionStatus']
                    }
                    $Item
                }

                [array]$RFDListItem = $AllRevisionFolderDashboardListItems | Where-Object -FilterScript $filterScript
                If ($RFDListItem.Count -eq 0) {
                    Write-Host "`nDocument $($VDLListItem.TCM_DN) - Rev: $($VDLListItem.Rev) not found on Revision Folder Dashboard" -ForegroundColor Yellow
                }
                Else
                { # missing condition to check if more then 1 item is found
                    Remove-PnPListItem -List 'Revision Folder Dashboard' -Identity $RFDListItem.ID -Recycle -Force -Connection $VDMSubSiteConnection | Out-Null
                    Write-Host "`nDocument $($RFDListItem.TCM_DN) - Rev: $($RFDListItem.Rev) deleted from Revision Folder Dashboard" -ForegroundColor Yellow
                    $DelProcessArray += 'RevFDashBoard'
                }

                If ($VDLListItem.Count -eq 0 -and $null -ne $VDDocFolder) {
                    # Orphaned document, remove the document from the SharePoint Online folder
                    $VDDocPathSplit = $VDDocumentPath.Split('/')
                    $VDDocParentFolderPath = ($VDDocPathSplit[6..($VDDocPathSplit.Length - 2)] -join '/')
                    Remove-PnPFolder -Name "$($VDFolderRev)" -Folder "$VDDocParentFolderPath" -Recycle -Force -Connection $VDMSubSiteConnection | Out-Null
                    $DelProcessArray += 'Folder'
                    Write-Host "`nDocument $($RFDListItem.TCM_DN) - Rev: $($RFDListItem.Rev). Folder $($VDFolderRev) deleted in $VDDocParentFolderPath" -ForegroundColor Yellow

                }
            }



            Start-Sleep -Milliseconds 500

            # Log the result in the csv file and write it to the console
            $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value "Document deleted from $($DelProcessArray -join ', ')" -Force
            $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
        }
        Else {
            # Log the result in the csv file and write it to the console
            $Document | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Document skipped, not in Placeholder status' -Force
            $Document | Export-Csv -Path $DeletationLogFilePath -Delimiter ';' -NoTypeInformation -Append
            Write-Host "`nDocument $($VDLListItem.TCM_DN) - Rev: $($VDLListItem.Rev) is not in Placeholder status" -ForegroundColor Yellow
        }
    }
    Catch {
        Write-Host "`nError deleting document $($Document.TCM_DN) - Rev: $($Document.Rev) - $($Error[0] | Out-String)" -ForegroundColor Red
    }

    # Uncomment to test only on first CSV row
    #Exit
}
Write-Progress -Activity 'Activity Completed' -Completed

# Disconnect from the site
# Disconnect-PnPOnline