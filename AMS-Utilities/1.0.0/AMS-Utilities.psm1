<#
    ToDo:
        Add PnP Module requirements and references
        Add parameter validation
        Test $Connection object and GetItem method after disconnecting from SharePoint Online
        Check $IsValidSPOConnectionObject with fake connection string
        Add parameter to GetAllItems method to get attachments (See function Get-ListFilesProperties below)
        Add Client Site ValidateScript
            [ValidateScript({
                If (!($_.Url.ToLower().EndsWith('c')))
                {
                    Throw ("`n `nInput error: provided URL does not end with 'C' or 'c', so it points to a TCM Site.`nPlease provide the URL of a Client Site instead.`n " -f $_.Url[-1])
                }
                Else
                {
                    Return $True
                }
            })]

#>

<# ArgumentCompleter
    Function to provide tab autocomplete when using this module's functions
#>
Function ArgumentCompleter {
    Param (
        $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters
    )

    $PossibleValues = @{
        Columns = @('InternalName', 'DisplayName', '$ListColumnsMapping')
    }

    if ($PossibleValues.ContainsKey($ParameterName) -and $CommandAst.CommandElements.Count -ge 2) {
        $PossibleValues[$ParameterName] | Where-Object {
            $_ -like "$WordToComplete*"
        }
    }
}

<# Get-SPOList
    Function to load a SharePoint Online List
#>
Function Get-SPOList {
    [CmdletBinding()]
    Param(
        # Name of the SharePoint Online List to get
        [Parameter(Mandatory = $true)]
        [String]
        $DisplayName,

        # The name type (Internal or Display names) of the columns to get from the List. Also a custom mapping can be used in PSObject format
        [Parameter(Mandatory = $false)]
        [ArgumentCompleter({ ArgumentCompleter @Args })]
        [Array]
        $Columns,

        # The SharePoint Online Connection object to be used to load the List (obtaind via Connect-PnPOnline)
        [Parameter(Mandatory = $false)]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection
    )

    Try {
        # Create List Object
        $SPOListObject = [SPOList]::New($DisplayName, $SPOConnection)

        # Set default column name type
        If ($null -eq $Columns) {
            $Columns = 'Internal'
        }

        # Get the List
        $SPOList = $SPOListObject.GetAllItems($($Columns))

        Return $SPOList
    }
    Catch {
        Write-Host ('Error:{0}' -f ($_ | Out-String).TrimEnd()) -ForegroundColor Red
    }
}

<#

###########################################
################## NOTES ##################
###########################################


<#
    Function to get all items in List ClientTransmittalQueue_Registry on Client site including attachments properties (Name, Path, Size, SHA-256)
#><#
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
        [Array]$ListColumns
    )

    # Get all items in List ClientTransmittalQueue_Registry on Client site
    Try
    {
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

        # Get raw List items from SPO
        Write-Host ("Getting all items on List '{0}' from site '{1}'" -f $ListName, $($ConnectionURL)) -ForegroundColor Cyan
        $SPO_List = Get-PnPListItem -List $ListName -Connection $Connection -PageSize 5000

        <#
            Get $MainColumnInternalName Column and all List Columns passed as $ListColumns parameter. Columns are searched both by Internal and DisplayName.
            If more than one column with the same name is found, only the one that has equal InternalName and DisplayName is returned.
            Consider changing this behaviour to return only Column for internal name if more then one is found.
        #><#
        If ($SPOList.BaseType -ne "DocumentLibrary")
        {
            $MainColumnInternalName = 'Title'
        }
        Else
        {
            $MainColumnInternalName = 'FileLeafRef'
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

        # Loop through all items and append them to structured PSObject
        $PSO_List = @()
        ForEach ($ListItem in $SPO_List)
        {
            #Write a progress bar
            $CurrentItem = $SPO_List.IndexOf($ListItem) + 1
            $PercentageComplete = (($CurrentItem / $SPO_List.Count) * 100)
            $ProgressBarSplatting = @{
                Activity         = ("Processing $($SPO_List.Count) items on List '{0}'" -f $ListName)
                Status           = ("Progress: {0}% ({1}/{2})" -f
                    [Math]::Round($PercentageComplete),
                    $CurrentItem, $SPO_List.Count)
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
                        $Item | Add-Member -MemberType NoteProperty -Name $($Column.InternalName) -Value $($ListItem[$Column.InternalName])
                        Break
                    }
                }
            }

            # Get all attachments' properties if supported by List type
            If ($SPOList.BaseType -eq "DocumentLibrary")
            {
                # Get main file properties
                $FileRelativeURL = $ListItem["FileRef"]
                $FileSize = $ListItem["File_x0020_Size"]

                # Calculate the SHA-256 hash
                $FileAsString = Get-PnPFile -Url $($FileRelativeURL) -Connection $Connection -AsString
                $FileContentBytes = [System.Text.Encoding]::UTF8.GetBytes($FileAsString)
                $SHA256HashBytes = (New-Object System.Security.Cryptography.SHA256Managed).ComputeHash($FileContentBytes)
                $FileSHA256Digest = [System.BitConverter]::ToString($SHA256HashBytes).Replace("-", "").ToLower()

                $ItemFileProperties = New-Object PSObject -Property @{
                    FileRelativeURL  = $FileRelativeURL
                    FileSize         = $FileSize
                    FileSHA256Digest = $FileSHA256Digest
                }
            }
            Else
            {
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
            $Item | Add-Member -MemberType NoteProperty -Name 'Attachments' -Value $ItemFileProperties

            # Add the object to PSO_ClientTransmittalQueueRegistry array
            $PSO_List += $Item
        }

        # Complete the progress bar
        Write-Progress -Activity ("Processing {0} items on List '{1}'" -f $($SPO_List.Count), $ListName) -Status 'Completed' -Completed
        Write-Host ("List '{0}' from site '{1}' loaded" -f $ListName, $($ConnectionURL)) -ForegroundColor Cyan
        Write-Host ''
        Return $PSO_List
    }
    Catch
    {
        Write-Progress -Activity ("Processing {0} items on List '{1}'" -f $($SPO_List.Count), $ListName)-Status 'Error' -Completed
        Write-Host ($_ | Out-String) -ForegroundColor Red
        Exit
    }
}
#>