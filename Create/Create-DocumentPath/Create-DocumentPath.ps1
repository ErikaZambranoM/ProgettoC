<#
    Questo script serve a modifica il percorso di un documento.
    Richiede un csv con l'attuale percorso del documento (DocPath) e il percorso del documento corretto (DocPathCalc)

    TODO:
    - Al momento supporta Document List (DD lato TCM);
    - At ($null -eq $DocumentLibrary), log and continue instead of Exit
    - In case of duplicates on List, return all IDs
    - Sub-ProgressBar
#>

param(
    [parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [parameter(Mandatory = $true)]
    [ValidateSet(
        'DocumentList',
        'Document List',
        'DL',
        'VendorDocumentsList', # Currently not supported
        'Vendor Documents List', # Currently not supported
        'VDL', # Currently not supported
        'Client Document List', # Currently not supported
        'ClientDocumentList', # Currently not supported
        'CDL' # Currently not supported
    )]
    [string]$ListName,

    [parameter(Mandatory = $true)]
    [string]$csvPath,

    [parameter(Mandatory = $true)]
    [ValidateSet(';', ',')]
    [string]$CSVDelimiter

)

$SiteUrl = $SiteUrl.TrimEnd('/')
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop

# Import CSV containing path to be corrected
[Array]$docPathToCreate = Import-Csv -Path $csvPath -Delimiter $CSVDelimiter

# Get List internal name from alias
Switch ($ListName)
{
    { $_ -in ('Document List', 'DocumentList', 'DL') }
    {
        $ListName = 'DocumentList'
    }

    { $_ -in ('Client Document List', 'ClientDocumentList', 'CDL') }
    {
        $ListName = 'Client Document List'
    }

    { $_ -in ('Vendor Documents List', 'VendorDocumentsList', 'VDL') }
    {
        $ListName = 'VendorDocumentsList'
    }

    default
    {
        Write-Host ("ERROR - List '$($ListName)' not supported! Exiting... ") -ForegroundColor Red
        Exit
    }
}

# Check if list exists
$ListCheck = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if ($null -eq $ListCheck)
{
    Write-Host ("ERROR - List '$($ListName)' not found on site $($SiteUrl)! Exiting... ") -ForegroundColor Red
    Exit
}

# Load list items and revision validation list
Switch ($ListName)
{
    'DocumentList'
    {
        $List = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID     = $_['ID']
                TCM_DN = $_['Title']
                Rev    = $_['IssueIndex']
                Path   = $_['DocumentsPath']
            }
            $item
        }

        <# Not needed, just compare Revision on provided paths and on Item's attribute
        $RevisionValidationList = Get-PnPListItem -List 'IssueIndexSequence' -PageSize 5000 | ForEach-Object {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID  = $_["ID"]
                Rev = $_["Title"]
            }
            $item
        }
        #>
    }

    'Client Document List'
    {
        # Currently not supported. All List fields must be adapted and tested. Also, the RevisionValidationList must be point to TCM AREA)
        Write-Host ("ERROR - List '$($ListName)' currently not supported! Exiting... ") -ForegroundColor Red
        Exit

        <#
        $List = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID     = $_["ID"]
                TCM_DN = $_["Title"]
                Rev    = $_["IssueIndex"]
                Path   = $_["DocumentsPath"]
            }
            $item
        }

        <# Not needed, just compare Revision on provided paths and on Item's attribute
        $RevisionValidationList = Get-PnPListItem -List 'IssueIndexSequence' -PageSize 5000 | ForEach-Object {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID     = $_["ID"]
                TCM_DN = $_["Title"]
                Rev    = $_["IssueIndex"]
                Path   = $_["DocumentsPath"]
            }
            $item
        }
        #>
        #>
    }

    'VendorDocumentsList'
    {
        # Currently not supported. All List fields must be adapted and tested. # Also check DocNumber and Rev validation during ForEach

        Write-Host ("ERROR - List '$($ListName)' currently not supported! Exiting... ") -ForegroundColor Red
        Exit

        <#
        $List = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID     = $_["ID"]
                TCM_DN = $_["Title"]
                Rev    = $_["IssueIndex"]
                Path   = $_["DocumentsPath"]
            }
            $item
        }

        <# Not needed, just compare Revision on provided paths and on Item's attribute (?)
        $RevisionValidationList = Get-PnPListItem -List 'RevisionIndexSequence' -PageSize 5000 | ForEach-Object {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID     = $_["ID"]
                TCM_DN = $_["Title"]
                Rev    = $_["IssueIndex"]
                Path   = $_["DocumentsPath"]
            }
            $item
        }
        #>
        #>
    }

    default
    {
        Write-Host ("ERROR - List '$($ListName)' not supported! Exiting... ") -ForegroundColor Red
        Exit
    }
}

# Create CSV file to log results
$SiteTitle = (Get-PnPWeb).Title
$ExecutionDate = Get-Date -Format 'yyyy_MM_dd_hh_mm_ss'
$filePath = "$($PSScriptRoot)\Logs\$($SiteTitle)_DocPathCreation$($ExecutionDate).csv";
if (!(Test-Path -Path $filePath)) { New-Item $filePath -Force -ItemType File | Out-Null }

# Process each item on CSV
foreach ($path in $docPathToCreate)
{
    # Progress bar for each item
    Write-Progress -Activity "Change path from $($path.DocPath) to $($path.DocPathCalc)" -Status "Processing row $($docPathToCreate.IndexOf($path) + 1)" -PercentComplete (($docPathToCreate.IndexOf($path) + 1) / $docPathToCreate.Count * 100)

    try
    {
        $Result = ''
        $ResultDetails = ''
        $DocMatchError = $null

        # Filter the list by both old and new document path provided on CSV
        $Doc = $List | Where-Object -FilterScript {
            $_.Path -eq $path.DocPath -or $_.Path -eq $path.DocPathCalc
        }

        # Path parts calculation
        $oldDocRelativePathSplit = $path.DocPath.Split('/')
        $oldDocRelativePath = $oldDocRelativePathSplit[5..($oldDocRelativePathSplit.Length - 2)] -join ('/')
        $newDocRelativePathSplit = $path.DocPathCalc.Split('/')
        $newDocRelativePath = $newDocRelativePathSplit[5..($newDocRelativePathSplit.Length - 3)] -join ('/')
        $documentLibraryInternalName = $newDocRelativePath[0].ToString()

        # Check if document path exists on list and skip if so
        If ($Doc.count -eq 0)
        {
            Write-Host ("WARNING - DocumentsPath '{0}' or '{1}' not found on List" -f $($path.DocPath), $($path.DocPathCalc)) -ForegroundColor Yellow
            $path | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Skipped'
            $path | Add-Member -MemberType NoteProperty -Name 'Result Details' -Value 'Both DocumentsPath were not found on List'
            $path | Export-Csv -Path $filePath -Delimiter ';' -Append -NoTypeInformation
            Continue
        }
        # Check if document path is duplicated on list and skip if so
        ElseIf ($Doc.count -gt 1 )
        {
            Write-Host ("WARNING - Duplicate found for path $($Doc.Path)") -ForegroundColor Yellow
            $path | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Skipped'
            $path | Add-Member -MemberType NoteProperty -Name 'Result Details' -Value 'Duplicate found'
            $path | Export-Csv -Path $filePath -Delimiter ';' -Append -NoTypeInformation
            Continue
        }
        # Check if item's DocumentsPath is coherent with its DocumentNumber and Revision
        Else
        {
            #Check if DocumentNumber is coherent with DocumentPath
            $CSVDocPathDN = $newDocRelativePathSplit[($newDocRelativePathSplit.Lenght - 2)]
            If ($CSVDocPathDN -ne $Doc.TCM_DN)
            {
                $DocMatchError = $true
                $ResultDetails = ("List DocumentNumber '{0}' is not coherent with '{1}' in provided DocumentPath. " -f $Doc.TCM_DN, $CSVDocPathDN)
            }

            # Check if Revision is coherent with DocumentPath
            $CSVDocPathRev = $newDocRelativePathSplit[-1]
            If ($CSVDocPathRev -ne $Doc.Rev)
            {
                $DocMatchError = $true
                $ResultDetails += ("List DocumentNumber '{0}' is not coherent with '{1}' in provided DocumentPath." -f $Doc.Rev, $CSVDocPathRev)
            }

            If ($DocMatchError)
            {
                Write-Host $ResultDetails -ForegroundColor Yellow
                $path | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Skipped'
                $path | Add-Member -MemberType NoteProperty -Name 'Result Details' -Value $ResultDetails
                $path | Export-Csv -Path $filePath -Delimiter ';' -Append -NoTypeInformation
                Continue
            }
        }


        #Controlla se la cartella è già stata spostata nella posizione corretta
        $isNewPathAlreadyValid = if ($null -eq (Get-PnPFolder -Url $path.DocPathCalc -ErrorAction SilentlyContinue)) { $false }else { $true }
        if ($isNewPathAlreadyValid -eq $false)
        {
            # Check if Document Library exists
            $DocumentLibrary = Get-PnPFolder -Url "/$($documentLibraryInternalName)" -ErrorAction SilentlyContinue
            If ($null -eq $DocumentLibrary)
            {
                # ToDo: Add log and continue
                Write-Host ("ERROR - Document Library $($documentLibraryInternalName) doesn't exist! Exiting... ") -ForegroundColor Red
                Exit
            }

            # Check if new path exists and create it if not
            $FolderToCheck = ''
            $newDocRelativePath.Split('/').Where({ $_ -ne $documentLibraryInternalName }) | ForEach-Object {
                $FolderToCheck += "$_/"
                $FolderToCheckRelativePath = $DocumentLibrary.ServerRelativeUrl + '/' + $FolderToCheck.TrimEnd('/') #$_
                $Folder = Get-PnPFolder -Url $FolderToCheckRelativePath -ErrorAction SilentlyContinue
                if ($null -eq $Folder)
                {
                    Add-PnPFolder -Name $_ -Folder $((Split-Path $FolderToCheckRelativePath).Replace('\', '/')) -ErrorAction SilentlyContinue | Out-Null
                    Write-Host ("SUCCESS - Missing Folder $($FolderToCheckRelativePath) created") -ForegroundColor Green
                }
            }

            # Move folder to new path
            Move-PnPFolder -Folder $oldDocRelativePath -TargetFolder $newDocRelativePath | Out-Null
            Write-Host ("SUCCESS - Folder moved from '{0}' to {1} on item $($doc.ID)" -f $($doc.Path), $($path.DocPathCalc)) -ForegroundColor Green
            $Result = 'Success'
            $ResultDetails = 'Document folder moved'
        }
        else
        {
            # Folder is already on correct path
            $ResultDetails = 'Document folder was already in correct path'
            $Result = 'Skipped'
            Write-Host ("SUCCESS - $ResultDetails '{0}' for item $($doc.ID)" -f $($path.DocPathCalc)) -ForegroundColor Gray

        }
        $path | Add-Member -MemberType NoteProperty -Name 'Result' -Value $Result
        $path | Add-Member -MemberType NoteProperty -Name 'Result Details' -Value $ResultDetails

        # Update document path on list item if needed
        If ($Doc.Path -ne $path.DocPathCalc)
        {
            Set-PnPListItem -List $ListName -Identity $Doc.ID -Values @{DocumentsPath = $Path.DocPathCalc } | Out-Null
            Write-Host ("SUCCESS - DocumentsPath changed from {0} to {1} on item $($Doc.ID)" -f $($doc.Path), $($path.DocPathCalc)) -ForegroundColor Green

            $ResultDetails = $ResultDetails + ', DocumentsPath path updated.'
            $path.'Result Details' = $ResultDetails
        }
        else
        {
            Write-Host ("SUCCESS - Document path '{0}' already correct  for item $($doc.ID)" -f $($path.DocPathCalc)) -ForegroundColor Gray
            $ResultDetails = $ResultDetails + ', DocumentsPath already correct.'
            $path.'Result Details' = $ResultDetails
        }
        $path | Export-Csv -Path $filePath -Delimiter ';' -Append -NoTypeInformation
    }

    catch
    {
        # Append error to CSV Result value
        Write-Host ($_ | Out-String) -ForegroundColor Red
        if ($null -eq $Path.Result)
        {
            $path | Add-Member -MemberType NoteProperty -Name 'Result' -Value 'Error'
            $path | Add-Member -MemberType NoteProperty -Name 'Result Details' -Value ($_ | Out-String)
        }
        else
        {
            $Path.Result = 'Partial error'
            $Path.'Result Details' += ("but`nPartial process error:`n" + ($_ | Out-String))
        }
        $path | Export-Csv -Path $filePath -Delimiter ';' -Append -NoTypeInformation
    }

    Write-Host ''
}

# Progress bar completed
Write-Progress -Activity "Processing item $($path.TCM_DN) - Rev $($path.Rev)" -Status "Processing item $($path.TCM_DN) - Rev $($path.Rev)" -Completed