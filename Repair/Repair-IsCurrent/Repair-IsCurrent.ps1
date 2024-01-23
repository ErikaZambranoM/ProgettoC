<#
    Get all Documents in descending index order and set the first to IsCurrent=$True, the others to IsCurrent=$False
    If a single TCM_Document_Number is passed, only that Document will be updated.
    If a CSV with TCM_DN is passed, only those Documents will be updated.

    !Priority ToDo:
        !If more then 1 revision is found and all have no LastTransmittal, return warning
        !Simulate param

    ToDo:
        ? Consider wrong index order
        !flag is current in file on VDFolder
        ! add check for $item members inside function
        ! better CSV/SingCode/Whole list handling
#>
Param (
    [Parameter(Mandatory = $true)]
    [String]$SiteURL
)

# Function to connect to a SharePoint Online Site or Sub Site and keep track of the connections
Function Connect-SPOSite {
    <#
    .SYNOPSIS
        Connects to a SharePoint Online Site or Sub Site.

    .DESCRIPTION
        This function connects to a SharePoint Online Site or Sub Site and returns the connection object.
        If a connection to the specified Site already exists, the function returns the existing connection object.

    .PARAMETER SiteUrl
        Mandatory parameter. Specifies the URL of the SharePoint Online site or subsite.

    .EXAMPLE
        PS C:\> Connect-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/contoso"
        This example connects to the "https://contoso.sharepoint.com/sites/contoso" site.

    .OUTPUTS
        The function returns an object with the following properties:
            - SiteUrl: The URL of the SharePoint Online site or subsite.
            - Connection: The connection object to the SharePoint Online site or subsite as returned by the Connect-PnPOnline cmdlet.
#>

    Param(
        # SharePoint Online Site URL
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                # Match a SharePoint Main Site or Sub Site URL
                If ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$') {
                    $True
                }
                Else {
                    Throw "`n'$($_)' is not a valid SharePoint Online site or subsite URL."
                }
            })]
        [String]
        $SiteUrl
    )

    Try {

        # Initialize Global:SPOConnections array if not already initialized
        If (-not $Script:SPOConnections) {
            $Script:SPOConnections = @()
        }
        Else {
            # Check if SPOConnection to specified Site already exists
            $SPOConnection = ($Script:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $SiteUrl }).Connection
        }

        # Create SPOConnection to specified Site if not already established
        If (-not $SPOConnection) {
            # Create SPOConnection to SiteURL
            $SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

            # Add SPOConnection to the list of connections
            $Script:SPOConnections += [PSCustomObject]@{
                SiteUrl    = $SiteUrl
                Connection = $SPOConnection
            }
        }

        Return $SPOConnection
    }
    Catch {
        Throw
    }
}

# Function to set IsCurrent on a Document
Function Set-IsCurrentDocumentRevision {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl, # URL of the SharePoint site

        [Parameter(Mandatory = $true)]
        $Item, # SharePoint List Item ID or custom object to update

        [Parameter(Mandatory = $true)]
        [Bool]$IsCurrent, # IsCurrent value to set

        [Parameter(Mandatory = $false)]
        [Bool]$SystemUpdate = $false
    )

    Try {
        If ($SystemUpdate) {
            $SystemUpdateSetting = @{
                UpdateType = 'SystemUpdate'
            }
        }
        Else {
            $SystemUpdateSetting = @{
                UpdateType = 'Update'
            }
        }

        # Infer ListType and ListName from SiteUrl
        If ($SiteUrl[-1] -eq '/') {
            $SiteUrl = $SiteUrl.TrimEnd('/')
        }
        If ($SiteUrl.ToLower().EndsWith('digitaldocumentsc') -or $SiteUrl.ToLower().EndsWith('ddwave2c')) {
            $ListType = 'CDL'
            $ListName = 'Client Document List'
            $IsCurrentField = 'IsCurrent'
        }
        ElseIf ($SiteUrl.ToLower().EndsWith('digitaldocuments') -or $SiteUrl.ToLower().EndsWith('ddwave2')) {
            $ListType = 'DD'
            $ListName = 'DocumentList'
            $IsCurrentField = 'IsCurrent'
        }
        ElseIf ($SiteURL.Split('/')[-1].ToLower().Contains('vdm_') -or $SiteUrl.Split('/')[-1].ToLower().Contains('poc_vdm')) {
            $ListType = 'VD'
            $ListName = 'Vendor Documents List'
            $IsCurrentField = 'VD_isCurrent'
        }
        Else {
            Write-Host 'Invalid SiteUrl!' -ForegroundColor Red
            Exit
        }

        # Set log file path
        $SiteName = $SiteUrl.Split('/')[-1]
        # Create Logs folder if it doesn't exist
        If (!(Test-Path -Path ($PSScriptRoot + '\Logs'))) {
            New-Item -Path ($PSScriptRoot + '\Logs') -ItemType Directory | Out-Null
        }
        $LogFilePath = $PSScriptRoot + "\Logs\$($SiteName)_FixIsCurrentLog_$(Get-Date -Format 'dd-MM-yyyy').log"

        # In case a single ID is passed, get the List Item
        If ($Item.GetType().Name -eq 'Int32') {
            # Get the List Item
            $Item = Get-PnPListItem -List $ListName -Id $Item | ForEach-Object {
                If ($ListType -eq 'DD' -or $ListType -eq 'CDL') {
                    $Item = New-Object -TypeName PSCustomObject -Property @{
                        ID            = $($_['ID'])
                        TCM_DN        = $($_['Title'])
                        Rev           = $($_['IssueIndex'])
                        ClientCode    = $($_['ClientCode'])
                        Index         = $($_['Index'])
                        Created       = $($_['Created'])
                        IsCurrent     = $($_[$($IsCurrentField)])
                        DocumentsPath = $($_['DocumentsPath'])
                        LastTrn       = $($_['LastTransmittal'])
                    }
                }
                ElseIf ($ListType -eq 'VD') {
                    $Item = New-Object -TypeName PSCustomObject -Property @{
                        ID            = $($_['ID'])
                        TCM_DN        = $($_['VD_DocumentNumber'])
                        Rev           = $($_['VD_RevisionNumber'])
                        ClientCode    = $($_['VD_ClientDocumentNumber'])
                        Index         = $($_['VD_Index'])
                        DocumentsPath = $($_['VD_DocumentPath'])
                        LastTrn       = $($_['LastTransmittal'])
                    }
                }
                $Item
            }
        }

        # Check if Index is set
        If ($Null -eq $Item.Index) {
            $Msg = ("ERROR ON '{0}': Index not set" -f $($Item.TCM_DN))
            Write-Host $Msg -ForegroundColor Red
            $Msg | Out-File -FilePath $LogFilePath -Append
            Return
        }

        # For List item, set IsCurrent to the passed value if it's different from the current value
        If ($Item.IsCurrent -ne $IsCurrent) {
            # Set IsCurrent to passed value
            Set-PnPListItem -List $ListName -Identity $Item.ID -Values @{$($IsCurrentField) = $IsCurrent } @SystemUpdateSetting | Out-Null
            $Msg = ("'{0}' - '{1}': changing IsCurrent from {2} to {3}" -f $($Item.TCM_DN), $($Item.Index), $Item.IsCurrent.ToString().ToUpper(), $IsCurrent.ToString().ToUpper())
            Write-Host $Msg -ForegroundColor Green
            $Msg | Out-File -FilePath $LogFilePath -Append
        }
        Else {
            $Msg = ("'{0}' - '{1}': IsCurrent already {2}" -f $($Item.TCM_DN), $($Item.Index), $IsCurrent.ToString().ToUpper())
            Write-Host $Msg -ForegroundColor DarkGray
            $Msg | Out-File -FilePath $LogFilePath -Append
        }

        # If needed, set IsCurrent on the MainContent file (based on ListType)
        Switch ($ListType) {
            # MainContent file need to be coherent with the IsCurrent value of the List Item
            'CDL' {
                # Get the folder containing the item's documents
                If ($Null -ne $Item.DocumentsPath) {
                    $Folder = Get-PnPFolder -Url $Item.DocumentsPath -ErrorAction SilentlyContinue
                }
                Else {
                    $Folder = $Null
                }

                If ($Null -ne $Folder) {
                    # Compose MainContentFileNames
                    $MainContentFileNames = @()
                    $MainContentFileNames += "$($Item.TCM_DN).pdf"
                    $MainContentFileNames += "$($Item.ClientCode).pdf"
                    $MainContentFileNames += "$($Item.TCM_DN)_$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.ClientCode)_$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.TCM_DN)-$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.ClientCode)-$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.TCM_DN)_IS$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.ClientCode)_IS$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.TCM_DN)-IS$($Item.Rev).pdf"
                    $MainContentFileNames += "$($Item.ClientCode)-IS$($Item.Rev).pdf"

                    # Get the relative URL of the folder
                    $ServerSiteRelativeUrl = $Item.DocumentsPath -Replace $SiteUrl, ''

                    # Get all files in the folder
                    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $ServerSiteRelativeUrl -ItemType File -ErrorAction SilentlyContinue

                    # Set IsCurrent on MainContent file
                    $MainContentFile = $Files | Where-Object { $_.Name -in $MainContentFileNames }

                    # No MainContent file found
                    If ($Null -eq $MainContentFile) {
                        $Msg = ("ERROR ON '{0}' - '{1}':`nMainContent file not found in DocumentsPath: {2}" -f $($Item.TCM_DN), $($Item.Index), $($Item.DocumentsPath))
                        Write-Host $Msg -ForegroundColor Red
                        $Msg | Out-File -FilePath $LogFilePath -Append

                        Continue
                    }
                    # More then 1 MainContent file found
                    ElseIf ($MainContentFile.Count -gt 1) {
                        $Msg = ('More than 1 MainContent file found in: {0}' -f $($Item.DocumentsPath))
                        Write-Host $Msg -ForegroundColor Red
                        $Msg | Out-File -FilePath $LogFilePath -Append

                        Continue
                    }
                    # Correctly found 1 MainContent file
                    Else {
                        # Get the file
                        $FileItem = Get-PnPFile -Url $MainContentFile.ServerRelativeUrl -AsListItem -ErrorAction SilentlyContinue
                        $DLLIst = $MainContentFile.ServerRelativeUrl.Split('/')[3]

                        # Set IsCurrent if needed
                        If ($IsCurrent -ne $FileItem[$($IsCurrentField)]) {
                            Set-PnPListItem -List $DLLIst -Identity $FileItem.Id -Values @{$($IsCurrentField) = $IsCurrent } @SystemUpdateSetting | Out-Null
                            $Msg = ("'{0}' - '{1}': changing IsCurrent on file '{2}' from {3} to {4}" -f $($Item.TCM_DN), $($Item.Index), $($MainContentFile.Name), $FileItem[$($IsCurrentField)].ToString().ToUpper(), $IsCurrent.ToString().ToUpper())
                            $MsgFColor = 'Green'
                        }
                        Else {
                            $Msg = ("'{0}' - '{1}': IsCurrent on file '{2}' already {3}" -f $($Item.TCM_DN), $($Item.Index), $($MainContentFile.Name), $IsCurrent.ToString().ToUpper())
                            $MsgFColor = 'DarkGray'
                        }
                        Write-Host $Msg -ForegroundColor $MsgFColor
                        $Msg | Out-File -FilePath $LogFilePath -Append

                    }
                }
                Else {
                    $Msg = ("'{0}' - '{1}': Folder not found at following DocumentsPath:`n{2}" -f $($Item.TCM_DN), $($Item.Index), $($Item.DocumentsPath))
                    Write-Host $Msg -ForegroundColor Red
                    $Msg | Out-File -FilePath $LogFilePath -Append

                    Continue
                }
                Break
            }

            'VD' {
                # Get the folder containing the item's documents
                $SubsiteUrl = $Item.DocumentsPath.Split('/')[0..5] -Join '/'
                $POLibrary = $Item.DocumentsPath.Split('/')[6]
                $SubSiteConnection = Connect-SPOSite -SiteUrl $SubsiteUrl
                $Folder = Get-PnPFolder -Url $Item.DocumentsPath -Connection $SubSiteConnection
                $FolderListItemProperties = Get-PnPProperty -ClientObject $Folder -Property ListItemAllFields -Connection $SubSiteConnection
                $ListItem = Get-PnPListItem -List $POLibrary -Id $FolderListItemProperties.Id -Connection $SubSiteConnection

                # Set IsCurrent on Document Folder
                If ($IsCurrent -ne $ListItem[$($IsCurrentField)]) {
                    Set-PnPListItem -Identity $ListItem -Values @{$($IsCurrentField) = $IsCurrent } -Connection $SubSiteConnection | Out-Null
                    $Msg = ("'{0}' - '{1}': changing IsCurrent on folder to {2}" -f $($Item.TCM_DN), $($Item.Index), $IsCurrent)
                    $MsgFColor = 'Green'
                }
                Else {
                    $Msg = ("'{0}' - '{1}': IsCurrent on folder already {2}" -f $($Item.TCM_DN), $($Item.Index), $IsCurrent.ToString().ToUpper())
                    $MsgFColor = 'DarkGray'
                }
                Write-Host $Msg -ForegroundColor $MsgFColor
                break
            }
        }
    }
    Catch {
        $Msg = ("ERROR ON '{0}' - '{1}':`n{2}" -f $($Item.TCM_DN), $($Item.Index), ($_ | Out-String))
        Write-Host $Msg -ForegroundColor Red
        $Msg | Out-File -FilePath $LogFilePath -Append
    }
}

try {
    # Check SystemUpdate
    Do {
        $SystemUpdate = Read-Host -Prompt 'System Update (true or false)'
    }
    While
    (
    ($SystemUpdate.ToLower() -notin ('true', 'false'))
    )
    $SystemUpdateSetting = @{
        SystemUpdate = [System.Convert]::ToBoolean($SystemUpdate)
    }

    # Set target Documents providing a single TCM_Document_Number, a CSV with more TCM_Document_Number listed or leave empty to target the whole List.
    $DocsToFixInput = Read-Host -Prompt 'CSV Path or TCM Document Number (leave empty to target the whole List)'
    If ($DocsToFixInput -eq '') {
        $DocsToFix = 'List'
    }
    ElseIf ($DocsToFixInput.ToLower().EndsWith('.csv')) {
        [Array]$DocsToFix = Import-Csv -Path $DocsToFixInput -Delimiter ';'
        If ($DocsToFix.Count -eq 0) {
            Write-Host 'No valid TCM_DN found in CSV' -ForegroundColor Red
            Exit
        }
    }
    Else {
        [String]$DocsToFix = $DocsToFixInput
    }

    # Connect to the site and infer the ListType and ListName based on SiteURL
    Connect-PnPOnline -Url $SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    If ($SiteUrl[-1] -eq '/') {
        $SiteURL = $SiteUrl.TrimEnd('/')
    }
    If ($SiteURL.ToLower().EndsWith('digitaldocumentsc') -or $SiteUrl.ToLower().EndsWith('ddwave2c')) {
        $ListType = 'CDL'
        $ListName = 'Client Document List'
    }
    ElseIf ($SiteURL.ToLower().EndsWith('digitaldocuments') -or $SiteUrl.ToLower().EndsWith('ddwave2')) {
        $ListType = 'DD'
        $ListName = 'DocumentList'
    }
    ElseIf ($SiteURL.Split('/')[-1].ToLower().Contains('vdm_') -or $SiteUrl.Split('/')[-1].ToLower().Contains('poc_vdm')) {
        $ListType = 'VD'
        $ListName = 'Vendor Documents List'
    }
    Else {
        Write-Host 'Invalid SiteUrl!' -ForegroundColor Red
        Exit
    }

    # Create filter script based on input type
    Switch ($DocsToFix) {
        # CSV
        { ($DocsToFix -is [Array]) } {
            $FilterScriptText = "(& {`$_.TCM_DN -in `$(`$DocsToFix.TCM_DN)})"
        }

        # Whole List
        'List' {
            $FilterScriptText = "(& {`$True})"
            Break
        }

        # Single Document
        Default {
            $FilterScriptText = "(& {`$_.TCM_DN -eq '$($DocsToFix)'})"
            Break
        }
    }
    $FilterScript = [ScriptBlock]::Create($FilterScriptText)

    # Load List
    Write-Host ("Loading List '{0}' on Site '{1}'" -f $ListName, $SiteURL) -ForegroundColor Cyan
    [Array]$ListItems = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
        If ($ListType -eq 'DD' -or $ListType -eq 'CDL') {
            $Item = New-Object -TypeName PSCustomObject -Property @{
                ID            = $($_['ID'])
                TCM_DN        = $($_['Title'])
                Rev           = $($_['IssueIndex'])
                ClientCode    = $($_['ClientCode'])
                Index         = $($_['Index'])
                Created       = $($_['Created'])
                IsCurrent     = $($_['IsCurrent'])
                DocumentsPath = $($_['DocumentsPath'])
                LastTrn       = $($_['LastTransmittal'])
            }
        }
        ElseIf ($ListType -eq 'VD') {
            $Item = New-Object -TypeName PSCustomObject -Property @{
                ID            = $($_['ID'])
                TCM_DN        = $($_['VD_DocumentNumber'])
                Rev           = $($_['VD_RevisionNumber'])
                ClientCode    = $($_['VD_ClientDocumentNumber'])
                Index         = $($_['VD_Index'])
                DocumentsPath = $($_['VD_DocumentPath'])
                LastTrn       = $($_['LastTransmittal'])
                IsCurrent     = $($_['VD_isCurrent'])

            }
        }
        $Item
    } | Where-Object -FilterScript $FilterScript | Sort-Object -Descending TCM_DN, Index

    If ($ListItems.Count -eq 0) {
        Write-Host ("No items found on List '{0}' on Site '{1}' matching '{2}'" -f $ListName, $SiteURL, $DocsToFix) -ForegroundColor Red
        Exit
    }
}
catch { throw }

# Loop through all List Items and set IsCurrent where required
$Counter = 0
$DocumentiRevisionati = @()
$NewLineDoc = @()
Try {
    ForEach ($Item in $ListItems) {
        # Progress bar
        $Counter++
        $Progress = [Math]::Round(($Counter / $ListItems.Count) * 100, 0)
        $ProgressText = "Processing item $($Counter) of $($ListItems.Count): $($Item.TCM_DN) - $($Item.Index) - $($Item.IsCurrent)"
        Write-Progress -Activity 'Processing items' -Status $ProgressText -PercentComplete $Progress

        If ($DocumentiRevisionati.Contains($Item.TCM_DN)) {
            Set-IsCurrentDocumentRevision -SiteUrl $SiteUrl -Item $Item -IsCurrent $False @SystemUpdateSetting
            $NewLineDoc += $Item.TCM_DN
        }
        Else {
            If (!($NewLineDoc.Contains($Item.TCM_DN))) {
                Write-Host ''
            }
            # If the Document has been transmitted or only 1 revision exists, set IsCurrent to True, else set it to False
            $DocRevsCount = @($ListItems | Where-Object -FilterScript { $_.TCM_DN -eq $Item.TCM_DN }).Count

            If ($Null -ne $Item.LastTrn -or $DocRevsCount -eq 1) {
                $DocumentiRevisionati += $Item.TCM_DN
                Set-IsCurrentDocumentRevision -SiteUrl $SiteUrl -Item $Item -IsCurrent $True @SystemUpdateSetting
            }
            Else {
                Set-IsCurrentDocumentRevision -SiteUrl $SiteUrl -Item $Item -IsCurrent $False @SystemUpdateSetting
            }
        }
        Start-Sleep -Seconds 0.150
    }
    Write-Progress -Activity 'Completed' -Completed
}
Catch {
    Write-Progress -Activity 'Error' -Completed
    $Msg = ("ERROR ON '{0}' - '{1}':`n{2}" -f $($Item.TCM_DN), $($Item.Index), ($_ | Out-String))
    Write-Host $Msg -ForegroundColor Red
    Exit
}