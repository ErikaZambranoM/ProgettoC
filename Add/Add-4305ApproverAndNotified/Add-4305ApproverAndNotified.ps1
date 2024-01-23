Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Approver, #=true # Set a default value only for testing purposes

    [Parameter(Mandatory = $false)]
    [switch]$Notified = $true # Set a default value only for testing purposes
)

try
{

    Write-Host ''
    if (-not $Approver -and -not $Notified)
    {
        Write-Error 'At least one of the switches (-Approver or -Notified) must be used.' -ErrorAction Stop
    }

    $TCMSiteConnection = Connect-PnPOnline -Url 'https://tecnimont.sharepoint.com/sites/4305DigitalDocuments' -ValidateConnection -UseWebLogin -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue
    $ClientSiteConnection = Connect-PnPOnline -Url 'https://tecnimont.sharepoint.com/sites/4305DigitalDocumentsc' -ValidateConnection -UseWebLogin -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue

    # Start Transcript
    if (-not (Test-Path -Path "$PSScriptRoot\Logs" -PathType Container))
    {
        New-Item -Path "$PSScriptRoot\Logs" -ItemType Directory -WhatIf:$false | Out-Null
    }
    $ScriptName = (Get-Item -Path $MyInvocation.MyCommand.Path).BaseName
    $ScriptRunDateTime = Get-Date -Format 'dd-MM-yyyy_HH-mm-ss'
    Start-Transcript -Path "$PSScriptRoot\Logs\$($ScriptName)_$($ScriptRunDateTime).log" -IncludeInvocationHeader
    Write-Host ''

    # Get all list items from 'Client Document List'
    Write-Host 'Loading Client Document List...' -ForegroundColor Cyan
    $CDL = Get-PnPListItem -List 'Client Document List' -Connection $ClientSiteConnection -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID             = $_['ID']
            IDDocumentList = $_['IDDocumentList']
            TCM_DN         = $_['Title']
            Rev            = $_['IssueIndex']
            Trn            = $_['LastTransmittal']
            Approver       = [Array]$_['InvolvedUsers'].Email -join "`n"
            Notified       = [Array]$_['InvolvedForNotify'].Email -join "`n"
        }
    }

    # Filter CDL items with a Trn value and without an Approver or Notified value
    $CDL = $CDL | Where-Object -FilterScript { $_.Trn -and (-not $_.Approver -or -not $_.Notified) }

    # Get all list items from list 'DocumentList'
    Write-Host 'Loading Document List...' -ForegroundColor Cyan
    $DL = Get-PnPListItem -List 'DocumentList' -Connection $TCMSiteConnection -PageSize 5000 | ForEach-Object {
        [PSCustomObject]@{
            ID                       = $_['ID']
            TCM_DN                   = $_['Title']
            Rev                      = $_['IssueIndex']
            DepartmentForNotify_Calc = $_['DepartmentForNotify_Calc']
        }
    }

    # Get all items from list 'Disciplines' from TCM site
    Write-Host 'Loading Disciplines...' -ForegroundColor Cyan
    $Disciplines = Get-PnPListItem -List 'Disciplines' -Connection $TCMSiteConnection | ForEach-Object {
        [PSCustomObject]@{
            ID             = $_['ID']
            DepartmentCode = $_['DepartmentCode']
            Recipients     = [Array]($_['Recipients']?.Split(';') | Where-Object -FilterScript { $_ })
            CCRecipients   = [Array]($_['CCRecipients']?.Split(';') | Where-Object -FilterScript { $_ })
        }
    }

    # Get all users from TCM site
    Write-Host 'Loading TCM Site Users...' -ForegroundColor Cyan
    $TCMSiteUsers = Get-PnPUser -Connection $TCMSiteConnection

    # Get all users from Client site
    Write-Host 'Loading Client Site Users...' -ForegroundColor Cyan
    $ClientSiteUsers = Get-PnPUser -Connection $ClientSiteConnection

    # Combine the users from TCM and Client site
    $TCM_and_Client_Users = $TCMSiteUsers + $ClientSiteUsers

    # Iterate through each item in CDL
    $MissingItems = @()
    $Batch = New-PnPBatch -RetainRequests -Connection $ClientSiteConnection
    Write-Host "`nLooping through all valid $($CDL.Count) items on Client Document List..." -ForegroundColor Cyan
    foreach ($CDL_Item in $CDL)
    {
        # Progress bar
        $Parameters = @{
            Activity        = 'Processing CDL Items'
            Status          = "TCM_DN: $($CDL_Item.TCM_DN), Rev: $($CDL_Item.Rev)"
            PercentComplete = (($CDL.IndexOf($CDL_Item) + 1) / $CDL.Count * 100)
        }
        Write-Progress @Parameters

        # Find the corresponding item in DL using IDDocumentList as primary key
        $Matching_DL_Item = $DL | Where-Object -FilterScript { $_.ID -eq $CDL_Item.IDDocumentList }
        if ($Matching_DL_Item)
        {
            # Get the value of Recipients and CCRecipients from 'Discplines' for the current item
            $MatchingDisciplines = $Disciplines | Where-Object -FilterScript { $_.DepartmentCode -eq $Matching_DL_Item.DepartmentForNotify_Calc }
            Write-Host ('[{0}/{1}] TCM_DN: {2}, Rev: {3}, DepartmentCode: {4}' -f
                ($CDL.IndexOf($CDL_Item) + 1),
                $CDL.Count,
                $($CDL_Item.TCM_DN),
                $($CDL_Item.Rev),
                $MatchingDisciplines.DepartmentCode
            ) -ForegroundColor Cyan

            # Create a hashtable to store the new values
            $NewValues = @{}

            # If Notified is not set on CDL_Item, add the value of CCRecipients to the hashtable
            if (-not $CDL_Item.Approver -and $Approver -and $MatchingDisciplines.Recipients)
            {
                $UsersToAdd = @()
                foreach ($Mail in $MatchingDisciplines.Recipients)
                {
                    $UsersToAdd += $TCM_and_Client_Users | Where-Object -FilterScript { $_.Email -eq $Mail }
                }
                Write-Host 'Adding Notified...' -ForegroundColor DarkCyan
                $NewValues.Add('InvolvedUsers', $UsersToAdd.LoginName)
                $CDL_Item.Approver = [Array]$UsersToAdd.Email -join "`n"
            }

            # If Notified is not set on CDL_Item, add the value of CCRecipients to the hashtable
            if (-not $CDL_Item.Notified -and $Notified -and $MatchingDisciplines.CCRecipients)
            {
                $UsersToAdd = @()
                foreach ($Mail in $MatchingDisciplines.CCRecipients)
                {
                    $UsersToAdd += $TCM_and_Client_Users | Where-Object -FilterScript { $_.Email -eq $Mail }
                }
                Write-Host 'Adding Notified...' -ForegroundColor DarkCyan
                $NewValues.Add('InvolvedForNotify', $UsersToAdd.LoginName)
                $CDL_Item.Notified = [Array]$UsersToAdd.Email -join "`n"
            }

            # Update the item in CDL
            if ($NewValues.Count -gt 0)
            {
                Set-PnPListItem -List 'Client Document List' -Identity $CDL_Item.ID -Values $NewValues -Batch $Batch -Connection $ClientSiteConnection | Out-Null
                Write-Host "Updated:`n$(($CDL_Item | Format-List | Out-String).Trim())`n" -ForegroundColor Green
            }
            else
            {
                Write-Host "No changes required.`n" -ForegroundColor Gray
            }
        }
        else
        {
            Write-Host "No matching item found for IDDocumentList $($CDL_Item.IDDocumentList)`n" -ForegroundColor Red
            $MissingItems += $CDL_Item
        }
    }

    if ($Batch.RequestCount -gt 0)
    {
        $BatchDetails = ((Invoke-PnPBatch -Batch $Batch -Connection $ClientSiteConnection -Details).ResponseJson | ConvertFrom-Json -Depth 100).Value
    }
    Remove-Variable -Name 'Batch' -Force
    New-Variable -Name 'BatchDetails' -Value $BatchDetails -Option ReadOnly -Scope Global -Force
    New-Variable -Name 'MissingItems' -Value $MissingItems -Option ReadOnly -Scope Global -Force
    Write-Host "Finished processing all items.`nCheck variable 'BatchDetails' and 'MissingItems' for more details." -ForegroundColor Green
}
catch
{
    throw
}
finally
{
    # End Progress Bar and Transcript
    Write-Progress -Activity 'Processing CDL Items' -Completed
    try { Stop-Transcript }catch {}
}