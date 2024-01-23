<# Notes:
    This script load files from Document Library 'TransmittalFromClient_Archive' on TCM site and items from List 'ClientTransmittalQueue_Registry' on Client site.
    Then it checks for every Tranmittal if file in 'TransmittalFromClient_Archive' is different then List item's attachment in 'ClientTransmittalQueue_Registry' and, if so, it replace the file on Client site.

    If a single TransmittalNumber is passed, only that one will be checked.
    If a CSV with TransmittalNumber is passed, only those ones will be checked.
    If no TransmittalNumber is passed, all Transmittals will be checked.

    ToDo:
        - Add SystemUpate param
        - Add choice with ProcessPreview and keys to Cancel, Simulate or confim Import
        - Change logic to permit to run the script on both directions (add DDCtoDD)
            - Generalize ForEach loop in a function that gets Source and Destination as parameters
#>
Param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            If ($_.ToLower().EndsWith('c'))
            {
                Throw ("`n `nInput error: provided URL ends with '{0}', so it points to a Client Site.`nPlease provide the URL of TCM Site instead, Client Area will be automatically set.`n " -f $_[-1])
            }
            Else
            {
                Return $True
            }
        })]
    [string]$TCMSiteURL
)

<#
    Input Validation
    Connect to SharePoint Online sites (both TCM and Client Area)
    Then check if Document Library 'TransmittalFromClient_Archive' and ClientTransmittalQueue_Registry exist
#>
Try
{

    # Reset variables
    $CSVPrimaryKeyColumn = $null
    $CSVPrimaryKeyColumnSplatting = @{CSVPrimaryKeyColumn = $Null }

    # Set target Documents providing a single TCM_Document_Number, a CSV with more TCM_Document_Number listed or leave empty to target the whole List.
    $TrnToFixInput = Read-Host -Prompt 'CSV Path or Transmittal Number (leave empty to target all Transmittals)'
    If ($TrnToFixInput -eq "")
    {
        $TrnToFix = 'All'
    }
    ElseIf ($TrnToFixInput.ToLower().EndsWith(".csv"))
    {
        # Check if CSV contains at least one Row
        [Array]$TrnToFix = Import-Csv -Path $TrnToFixInput -Delimiter ";"
        If ($TrnToFix.Count -eq 0)
        {
            Write-Host "No valid TransmittalNumber found in CSV" -ForegroundColor Red
            Exit
        }

        # Ask for the column in the CSV used to match List items
        Do
        {
            $CSVPrimaryKeyColumn = Read-Host -Prompt 'Column in the CSV used to match List items'
            If
            (
                $CSVPrimaryKeyColumn -notin ($TrnToFix | Get-Member -MemberType NoteProperty).Name -or
                $null -eq $CSVPrimaryKeyColumn -or '' -eq $CSVPrimaryKeyColumn
            )
            {
                Write-Host "Column '$CSVPrimaryKeyColumn' does not exist in CSV" -ForegroundColor Yellow
            }
        }
        Until
        # CSVPrimaryKeyColumn exists in CSV
        (
            ($CSVPrimaryKeyColumn -In ($TrnToFix | Get-Member -MemberType NoteProperty).Name) -and
            ($null -ne $CSVPrimaryKeyColumn -and '' -ne $CSVPrimaryKeyColumn)
        )
        $CSVPrimaryKeyColumnSplatting = @{CSVPrimaryKeyColumn = $CSVPrimaryKeyColumn }
    }
    Else
    {
        [String]$TrnToFix = $TrnToFixInput
    }

    # Connect to sites and check for List and Library existence
    $SourceSiteConnection = Connect-PnPOnline -Url $TCMSiteURL -UseWebLogin -ValidateConnection -ReturnConnection
    $ClientSiteURL = $TCMSiteURL + 'C'
    $DestinationSiteConnection = Connect-PnPOnline -Url $ClientSiteURL -UseWebLogin -ValidateConnection -ReturnConnection

    If ($null -eq (Get-PnPFolder -Url "TransmittalFromClient_Archive" -ErrorAction SilentlyContinue -Connection $SourceSiteConnection))
    {
        Write-Host "Document Library 'TransmittalFromClient_Archive' does not exist on site $($TCMSiteURL)" -ForegroundColor Red
        Exit
    }
    If ($null -eq (Get-PnPList -Identity "ClientTransmittalQueue_Registry" -ErrorAction SilentlyContinue -Connection $DestinationSiteConnection))
    {
        Write-Host ("List 'ClientTransmittalQueue_Registry' does not exist on site '{0}'" -f $($ClientSiteURL)) -ForegroundColor Red
        Exit
    }
}
Catch
{
    Write-Host ($_ | Out-String) -ForegroundColor Red
    Exit
}

<#
    Function to get all items in List ClientTransmittalQueue_Registry on Client site including attachments properties (Name, Path, Size, SHA-256)
#>
Function Get-ListFilesProperties
{
    Param (

        [Parameter(Mandatory = $true)]
        [String]$ListName,

        # Connection parameters which reconize if the URL is a TCM or Client site
        [Parameter(Mandatory = $true)]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $Connection,

        # Internal or Display names of the columns to get from the List
        [Parameter(Mandatory = $false)]
        [Array]$ListColumns,

        # Column in the CSV used to match List items
        [Parameter(Mandatory = $false)]
        [String]$CSVPrimaryKeyColumn,

        # Filter to apply to the List (applied on Title column for Lists or FileLeafRef for Document Libraries)
        [Parameter(Mandatory = $false)]
        $Filter = 'All'
    )

    # Get all items in List ClientTransmittalQueue_Registry on Client site
    Try
    {
        # Validate input parameters for CSV
        If
        # CSV has been provided but without CSVPrimaryKeyColumn
        (
            ($Filter -is [Array] -and $Filter[0].GetType().Name -eq 'PSCustomObject') -and
            ($Null -eq $CSVPrimaryKeyColumn -or '' -eq $CSVPrimaryKeyColumn)
        )
        {
            Throw "Please specify a CSVPrimaryKeyColumn when using a CSV file as input"
            Exit
        }
        # CSVPrimaryKeyColumn has been provided but without CSV
        ElseIf
        (
            ($Filter -isnot [Array] -and $Filter[0].GetType().Name -ne 'PSCustomObject') -and
            ($Null -ne $CSVPrimaryKeyColumn -and '' -ne $CSVPrimaryKeyColumn)
        )
        {
            Throw "CSVPrimaryKeyColumn parameter has been provided but input Filter is not a CSV file"
            Exit
        }
        # Both CSV and CSVPrimaryKeyColumn have been provided
        ElseIf
        (
            ($Filter -is [Array] -and $Filter[0].GetType().Name -eq 'PSCustomObject') -and
            ($Null -ne $CSVPrimaryKeyColumn -and '' -ne $CSVPrimaryKeyColumn)
        )
        {
            $CSVRows = $Filter
            $Filter = 'CSV'
        }

        # Get SiteURL from passed Connection
        $ConnectionURL = (Get-PnPWeb -Connection $Connection).Url

        # Get the List
        $SPOList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue -Connection $Connection

        # Check if List exists
        If ($Null -eq $($SPOList))
        {
            Write-Host ("List '{0}' does not exist on site '{1}'" -f $ListName, $($ConnectionURL)) -ForegroundColor Red
            Exit
        }

        <#
            Get $MainColumnInternalName Column and all List Columns passed as $ListColumns parameter. Columns are searched both by Internal and DisplayName.
            If more than one column with the same name is found, only the one that has equal InternalName and DisplayName is returned.
            Consider changing this behaviour to return only Column for internal name if more then one is found.
        #>
        If ($SPOList.BaseType -ne "DocumentLibrary")
        {
            $MainColumnInternalName = 'Title'
            $FilterColumn = "['Title']"
        }
        Else
        {
            $MainColumnInternalName = 'FileLeafRef'
            $FilterColumn = ".FieldValues.FileLeafRef"
        }
        $ListFields = Get-PnPField -List $ListName -Connection $Connection | Where-Object -FilterScript {
            $_.InternalName -eq $MainColumnInternalName -or
            ($_.InternalName -in $ListColumns -or $_.Title -in $ListColumns)
        }  |
        Group-Object -Property Title |
        ForEach-Object {
            If ($_.Count -gt 1)
            {
                $_.Group | Where-Object { $_.InternalName -eq $_.Title }
            }
            Else
            {
                $_.Group
            }
        }

        # Compose filter based on input parameter
        Switch ($Filter)
        {
            # Whole List
            'All'
            {
                $FilterScriptText = "(& {`$True})"
                Break
            }

            # CSV file
            'CSV'
            {
                # Filter items where Title or FileLeafRef (based on List type) is in the CSV in the column with name equal to $MainColumnInternalName
                $FilterScriptText = "(& {`$([System.IO.Path]::GetFileNameWithoutExtension(`$_$($FilterColumn))) -in `$(`$CSVRows.$($CSVPrimaryKeyColumn))})"
                Break
            }

            # No value
            { ($Null -eq $Filter -or '' -eq $Filter) }
            {
                Throw "Empty filter value passed"
                Break
            }

            # Single Document
            Default
            {
                $FilterScriptText = "(& {`$([System.IO.Path]::GetFileNameWithoutExtension(`$_$($FilterColumn))) -eq '$($Filter)'})"
                Break
            }
        }
        $FilterScript = [ScriptBlock]::Create($FilterScriptText)

        # Get raw List items from SPO
        Write-Host ("Getting requested items on List '{0}' from site '{1}'." -f $ListName, $($ConnectionURL)) -ForegroundColor Cyan
        [Array]$SPO_ListItems = Get-PnPListItem -List $ListName -Connection $Connection -PageSize 5000 | Where-Object -FilterScript $FilterScript

        # Check if items were found
        If ($SPO_ListItems.Count -eq 0)
        {
            Write-Host ("Requested items not found on List '{0}' from site '{1}'." -f $ListName, $($ConnectionURL)) -ForegroundColor Yellow
            Exit
        }

        # Loop through all items and append them to structured PSObject
        $PSO_List = @()
        ForEach ($ListItem in $SPO_ListItems)
        {
            #Write a progress bar
            $CurrentItem = $SPO_ListItems.IndexOf($ListItem) + 1
            $PercentageComplete = (($CurrentItem / $SPO_ListItems.Count) * 100)
            $ProgressBarSplatting = @{
                Activity         = ("Processing {0} items on List '{1}'" -f $($SPO_ListItems.Count), $ListName)
                Status           = ("Progress: {0}% ({1}/{2})" -f
                    [Math]::Round($PercentageComplete),
                    $CurrentItem, $SPO_ListItems.Count)
                CurrentOperation = ("Processing item with ID '{0}' and {1} '{2}'" -f
                    $($ListItem['ID']),
                    $($ListFields.Where({ $_.InternalName -eq $MainColumnInternalName }).Title),
                    $($ListItem["$($MainColumnInternalName)"])
                )
                PercentComplete  = $PercentageComplete
            }
            Write-Progress @ProgressBarSplatting

            # Create a new temporary object with ID and Attachments properties
            $Item = New-Object PSObject -Property ([Ordered]@{
                    ID = $ListItem['ID']
                })

            # Add List Columns passed as $ListColumns parameter
            ForEach ($Column in $ListFields)
            {
                Switch ($Column.InternalName)
                {
                    'FileLeafRef'
                    {
                        $Item | Add-Member -MemberType NoteProperty -Name 'FileName' -Value $($ListItem[$Column.InternalName])
                        Break
                    }

                    Default
                    {
                        $Item | Add-Member -MemberType NoteProperty -Name $($Column.Title) -Value $($ListItem[$Column.InternalName])
                        Break
                    }
                }
            }

            # Get all attachments' properties if supported by List type
            If ($SPOList.BaseType -eq "DocumentLibrary")
            {
                # Set main file properties
                $FileRelativeURL = $ListItem["FileRef"]
                $FileSize = $ListItem["File_x0020_Size"]
                $FilePropertiesColumnName = 'File Properties'

                # Calculate the SHA-256 hash
                $FileAsString = Get-PnPFile -Url $($FileRelativeURL) -Connection $Connection -AsString
                $FileContentBytes = [System.Text.Encoding]::UTF8.GetBytes($FileAsString)
                $SHA256HashBytes = (New-Object System.Security.Cryptography.SHA256Managed).ComputeHash($FileContentBytes)
                $FileSHA256Digest = [System.BitConverter]::ToString($SHA256HashBytes).Replace("-", "").ToLower()

                $ItemFileProperties = New-Object PSObject -Property @{
                    RelativeURL  = $FileRelativeURL
                    Size         = $FileSize
                    SHA256Digest = $FileSHA256Digest
                }
            }
            Else
            {
                $FilePropertiesColumnName = 'Attachments'
                $ItemFileProperties = @((Get-PnPProperty -ClientObject $ListItem -Property "AttachmentFiles" -Connection $Connection) | ForEach-Object {
                        # Get attachment file size
                        $FileSize = (Get-PnPFile -Url $($_.ServerRelativeUrl) -AsFileObject -Connection $Connection).Length

                        # Calculate the SHA-256 hash
                        $FileAsString = Get-PnPFile -Url $($_.ServerRelativeUrl) -Connection $Connection -AsString
                        $FileContentBytes = [System.Text.Encoding]::UTF8.GetBytes($FileAsString)
                        $SHA256HashBytes = (New-Object System.Security.Cryptography.SHA256Managed).ComputeHash($FileContentBytes)
                        $FileSHA256Digest = [System.BitConverter]::ToString($SHA256HashBytes).Replace("-", "").ToLower()

                        # Create a new temporary object with attachment properties
                        $ItemAttachments = New-Object PSObject -Property @{
                            AttachmentRelativeUrl = $_.ServerRelativeUrl
                            AttachmentFileName    = $_.FileName
                            AttachmentSize        = $FileSize
                            AttachmentSHA256      = $FileSHA256Digest
                        }
                        $ItemAttachments
                    })
            }

            # Add file properties to temporary object
            $Item | Add-Member -MemberType NoteProperty -Name $FilePropertiesColumnName -Value $ItemFileProperties

            # Add the object to PSO_ClientTransmittalQueueRegistry array
            $PSO_List += $Item
        }

        # Complete the progress bar
        Write-Progress -Activity ("Processing {0} items on List '{1}'" -f $($SPO_ListItems.Count), $ListName) -Status 'Completed' -Completed
        Write-Host ("List '{0}' from site '{1}' loaded." -f $ListName, $($ConnectionURL)) -ForegroundColor Cyan
        Write-Host ''
        Return $PSO_List
    }
    Catch
    {
        Write-Progress -Activity ("Processing {0} items on List '{1}'" -f $($SPO_ListItems.Count), $ListName) -Status 'Error' -Completed
        Write-Host ($_ | Out-String) -ForegroundColor Red
        Exit
    }
}

# Get all items in List ClientTransmittalQueue_Registry and all files in the Document Library 'TransmittalFromClient_Archive'
Try
{
    $ClientTransmittalQueue_Registry = $null
    [Array]$ClientTransmittalQueue_Registry = Get-ListFilesProperties -ListName "ClientTransmittalQueue_Registry" -Filter $TrnToFix @CSVPrimaryKeyColumnSplatting -Connection $DestinationSiteConnection

    $TransmittalFromClient_Archive = $null
    [Array]$TransmittalFromClient_Archive = Get-ListFilesProperties "TransmittalFromClient_Archive" -ListColumns "Transmittal Number" -Filter $TrnToFix @CSVPrimaryKeyColumnSplatting -Connection $SourceSiteConnection
}
Catch
{
    Write-Host ($_ | Out-String) -ForegroundColor Red
    Exit
}

# Path to CSV export log
$ProjectCode = $TCMSiteURL.Split("/")[-1] -replace 'DigitalDocuments.*'
$CSVExportLogPath = "$($PSScriptRoot)\$($ProjectCode)_Import-FromClientCoverDDtoDDC_Log_$(Get-Date -Format 'dd-MM-yyyy_HH-mm-ss').csv"

# Cover comparison
<#
    Loop through every file in the Document Library 'TransmittalFromClient_Archive',
    compare both SHA-256 and size with attachments on List ClientTransmittalQueue_Registry.
    If differences are found, replace the List items' attachment with the file in TransmittalFromClient_Archive.
#>
<# TODO:
    - Create bidirectional function for following code
    - Duplicated $File.'Transmittal Number' ?
#>
ForEach ($File in $TransmittalFromClient_Archive)
{
    Try
    {
        # Create Log object to be exported to CSV
        $LogRow = [PSCustomObject]@{
            SourceTransmittalNumber  = $($File.'Transmittal Number')
            TransmittalCoverSourceID = $($File.ID)
        }

        #Write a progress bar
        $CurrentItem = $TransmittalFromClient_Archive.IndexOf($File) + 1
        $PercentageComplete = (($CurrentItem / $TransmittalFromClient_Archive.Count) * 100)
        $ProgressBarSplatting = @{
            Activity         = "Comparing cover files from 'TransmittalFromClient_Archive' with attachments in 'ClientTransmittalQueue_Registry'"
            Status           = ("Progress: {0}% ({1}/{2})" -f
                [Math]::Round($PercentageComplete),
                $CurrentItem, $TransmittalFromClient_Archive.Count)
            CurrentOperation = ("Currently processing '{0}'" -f $File.FileName)
            PercentComplete  = $PercentageComplete
        }
        Write-Progress @ProgressBarSplatting

        # Filter List ClientTransmittalQueue_Registry by Transmittal Number equal to itered file
        [Array]$List_TransmittalItemToCheck = $ClientTransmittalQueue_Registry | Where-Object -FilterScript { $_.Title -eq $File.'Transmittal Number' }

        # Check if there is only one item with the same Transmittal Number
        If ($List_TransmittalItemToCheck.Count -gt 1)
        {
            Write-Host ("Multiple Transmittal items found for '{0}' on List 'ClientTransmittalQueue_Registry'" -f $($List_TransmittalItemToCheck.Title)) -ForegroundColor Red
            Write-Host ("Please check the List and remove duplicates (ID: {0})" -f $($List_TransmittalItemToCheck.ID -Join ', ')) -ForegroundColor Red
            $LogRow | Add-Member -MemberType NoteProperty -Name TransmittalCoverTargetID -Value $($List_TransmittalItemToCheck.ID -Join ', ')
            $LogRow | Add-Member -MemberType NoteProperty -Name Result -Value "Skipped"
            $LogRow | Add-Member -MemberType NoteProperty -Name ResultDetails -Value "Multiple Transmittal items found"
        }
        ElseIf ($List_TransmittalItemToCheck.Count -eq 0)
        {
            Write-Host ("No Transmittal item found for '{0}' on List 'ClientTransmittalQueue_Registry'" -f $($List_TransmittalItemToCheck.Title)) -ForegroundColor Red
            Write-Host "Please check the List and add the missing Transmittal item" -ForegroundColor Red
            $LogRow | Add-Member -MemberType NoteProperty -Name TransmittalCoverTargetID -Value 'Not found'
            $LogRow | Add-Member -MemberType NoteProperty -Name Result -Value "Skipped"
            $LogRow | Add-Member -MemberType NoteProperty -Name ResultDetails -Value "Transmittal item not found"
        }
        Else
        {
            # Filter item attachment with same name of transmittal cover
            $List_TransmittalCoverToCheck = $List_TransmittalItemToCheck | Where-Object -FilterScript { $_.Attachments.AttachmentFileName -eq $($File.FileName) }

            # Compare SHA-256 and size of the file with the ones in the List item
            If (
                $($List_TransmittalCoverToCheck.Attachments.AttachmentSHA256) -ne $($File.'File Properties'.SHA256Digest) -or
                $($List_TransmittalCoverToCheck.Attachments.AttachmentSize) -ne $($File.'File Properties'.Size)
            )
            {
                # Check if no cover file is attached so it only needs to be added
                If ($Null -eq $List_TransmittalCoverToCheck)
                {
                    Write-Host ("Adding missing Transmittal Cover for '{0}' on List 'ClientTransmittalQueue_Registry'" -f $($List_TransmittalItemToCheck.Title)) -ForegroundColor Yellow
                    $Msg = "Transmittal Cover added"
                }
                # Cover file is attached but it is not the correct one so it need to be deleted.
                ELse
                {
                    # Delete provided cover attachment from the List Item
                    Write-Host ("Replacing Transmittal Cover for '{0}' on List 'ClientTransmittalQueue_Registry'" -f $($List_TransmittalItemToCheck.Title)) -ForegroundColor Yellow
                    Remove-PnPListItemAttachment -List 'ClientTransmittalQueue_Registry' -Identity $($List_TransmittalItemToCheck.ID) -FileName $($File.FileName) -Recycle -Force -Connection $DestinationSiteConnection
                    Write-Host "Wrong Transmittal Cover deleted" -ForegroundColor Green
                    $Msg = "Transmittal Cover replaced"
                }

                # Upload the file as attachment to the List item
                $SourceFileStream = Get-PnPFile -Url $($File.'File Properties'.RelativeURL) -Connection $SourceSiteConnection -AsMemoryStream
                <#
                    Error or non-exisiting attachments folder
                    $AttachmentPath = ('/sites/{0}/Lists/ClientTransmittalQueue_Registry/Attachments/{1}' -f $($TCMSiteURL.Split("/")[-1] + 'C'), $List_TransmittalItemToCheck.ID)
                    Add-PnPFile -FileName $($File.FileName) -Folder $AttachmentPath -Stream  $SourceFileStream -Connection $DestinationSiteConnection
                #>
                Add-PnPListItemAttachment -List 'ClientTransmittalQueue_Registry' -Identity $List_TransmittalItemToCheck.ID -FileName $($File.FileName) -Stream $SourceFileStream -Connection $DestinationSiteConnection | Out-Null
                Write-Host ($Msg + "(ID: {0})" -f $($List_TransmittalItemToCheck.ID)) -ForegroundColor Green
                $LogRow | Add-Member -MemberType NoteProperty -Name TransmittalCoverTargetID -Value $List_TransmittalItemToCheck.ID
                $LogRow | Add-Member -MemberType NoteProperty -Name Result -Value "Success"
                $LogRow | Add-Member -MemberType NoteProperty -Name ResultDetails -Value $Msg
            }
            # The cover file is already correct
            Else
            {
                Write-Host ("Transmittal Cover for '{0}' on List 'ClientTransmittalQueue_Registry' is already correct (ID: {1})" -f $($List_TransmittalItemToCheck.Title), $($List_TransmittalItemToCheck.ID)) -ForegroundColor Green
                $LogRow | Add-Member -MemberType NoteProperty -Name TransmittalCoverTargetID -Value $List_TransmittalItemToCheck.ID
                $LogRow | Add-Member -MemberType NoteProperty -Name Result -Value "Skipped"
                $LogRow | Add-Member -MemberType NoteProperty -Name ResultDetails -Value "Transmittal Cover already correct"
            }
        }
        Write-Host ''
        $LogRow | Export-Csv -Path $CSVExportLogPath -Append -Delimiter ';' -NoTypeInformation -Encoding UTF8
    }
    Catch
    {
        # Log the error and continue with the next file
        Write-Host ''
        Write-Host ("Error on Transmittal '{0}'" -f $($File.'Transmittal Number')) -ForegroundColor Red
        Write-Host ($_ | Out-String) -ForegroundColor Red
        Write-Host ''
        $LogRow | Add-Member -MemberType NoteProperty -Name Result -Value "Error"
        $LogRow | Add-Member -MemberType NoteProperty -Name ResultDetails -Value ($_ | Out-String)
        $LogRow | Export-Csv -Path $CSVExportLogPath -Append -Delimiter ';' -NoTypeInformation -Encoding UTF8
    }
}
# Terminate the progress bar
Write-Progress -Activity 'Processing completed' -Status 'Completed' -Completed