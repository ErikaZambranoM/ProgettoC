# Script to move Documents in Reserved status to the proper folder

Param
(
    [Parameter(Mandatory = $true, HelpMessage = "The SharePoint site URL.")]
    [String]
    $SiteUrl,

    [Parameter(Mandatory = $false, HelpMessage = "To only export the list of items in Reserved status to CSV.")]
    [Switch]
    $ReportOnly
)


# Connect to the SharePoint Site
Try
{
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue
    $CSVPath = "$($PSScriptRoot)\$($SiteUrl.Split('/')[-1])ReservedDocumentsPath.csv"

    # Get all items where Reserved is 'Yes'
    $AllItems = Get-PnPListItem -List "DocumentList" -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID            = $_["ID"]
            TCM_DN        = $_["Title"]
            Rev           = $_["IssueIndex"]
            Reserved      = $_["Reserved"]
            DocumentsPath = $_["DocumentsPath"]
        }
    }

    [Array]$ReservedItems = $AllItems | Where-Object { $_.Reserved -eq $true }

    # Initialize the progress bar
    $totalItems = $ReservedItems.Count
    $currentIndex = 0

    # Loop through each item
    ForEach ($item in $ReservedItems)
    {
        $currentIndex++
        Write-Progress -PercentComplete (($currentIndex / $totalItems) * 100) -Activity 'Moving Documents' -Status "Processing" -CurrentOperation "$currentIndex out of $totalItems"

        # If the ReportOnly switch is set, just export the list of items to CSV
        If ($ReportOnly)
        {
            If ($item.DocumentsPath -like "*Reserved*")
            {
                $item | Add-Member -MemberType NoteProperty -Name IsCorrectPath -Value 'TRUE'
                $item | Export-Csv -Path $CSVPath -Append -NoTypeInformation -Delimiter ";"
            }
            Else
            {
                $item | Add-Member -MemberType NoteProperty -Name IsCorrectPath -Value 'FALSE'
                $item | Export-Csv -Path $CSVPath -Append -NoTypeInformation -Delimiter ";"
            }
        }
        Else
        # Move the documents to the Reserved folder if needed
        {
            if ($item.DocumentsPath -like "*Reserved*")
            {
                Write-Host "'$($item.TCM_DN) - $($item.Rev)' already in Reserved folder"
                $item | Add-Member -MemberType NoteProperty -Name ReservedPath -Value $item.DocumentsPath
                $item | Add-Member -MemberType NoteProperty -Name Result -Value 'Already Reserved'
                $item | Export-Csv -Path $CSVPath -Append -NoTypeInformation -Delimiter ";"
                Continue
            }

            # Compose the destination URL
            $NewDocumentsPath = "$($SiteUrl)/Reserved$($item.DocumentsPath.Replace($SiteUrl, ''))"
            $DestinationRelativeUrlSplit = $NewDocumentsPath.Split('/')
            $DestinationRelativeUrl = "$($DestinationRelativeUrlSplit[0..($DestinationRelativeUrlSplit.Count - 2)] -Join '/')" -replace "$($SiteUrl)/", ''

            $item | Add-Member -MemberType NoteProperty -Name ReservedPath -Value $NewDocumentsPath
            Resolve-PnPFolder -SiteRelativePath $DestinationRelativeUrl | Out-Null

            # Move document to the new location
            Write-Host "Moving '$($item.TCM_DN) - $($item.Rev)' to $NewDocumentsPath"
            Move-PnPFolder -Folder $($item.DocumentsPath.Replace($SiteUrl, '').Trim('/')) -TargetFolder  $DestinationRelativeUrl | Out-Null

            # Update the DocumentsPath field
            Set-PnPListItem -List "DocumentList" -Identity $item.ID -Values @{"DocumentsPath" = $NewDocumentsPath } | Out-Null

            $item | Add-Member -MemberType NoteProperty -Name Result -Value 'Success'
            $item | Export-Csv -Path $CSVPath -Append -NoTypeInformation -Delimiter ";"
        }
    }

    Write-Progress -Activity 'Moving Documents' -Completed
}
Catch
{
    Write-Progress -Activity 'Moving Documents' -Completed
    $item | Add-Member -MemberType NoteProperty -Name Result -Value 'Failed'
    $item | Export-Csv -Path $CSVPath -Append -NoTypeInformation -Delimiter ";"
    Throw
}
