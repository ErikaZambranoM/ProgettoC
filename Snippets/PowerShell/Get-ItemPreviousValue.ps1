<#
.SYNOPSIS
    Get the previous value of a field for a SharePoint Online list item.
    Returns a custom object containing the version with the previous value and the comparison with current value.

    ToDo:
        - Add parameter validation
        - Add parameter to return all versions
        - Add parameter to return first available version
        - Add default field to add to output for common lists
#>

Param (
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Url.')]
    [ValidateNotNullOrEmpty()]
    [String]
    $SiteUrl,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the List Name.')]
    [ValidateNotNullOrEmpty()]
    [String]
    $ListName,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the ID of the item to get the values from.')]
    [ValidateNotNullOrEmpty()]
    [String]
    $ItemID,

    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the column internal name to retrieve the value for.')]
    [ValidateNotNullOrEmpty()]
    [String]
    $FieldInternalName,

    [Parameter(Mandatory = $false, HelpMessage = 'SharePoint Online Connection object')]
    [ValidateNotNullOrEmpty()]
    [PnP.PowerShell.Commands.Base.PnPConnection]
    $SPOConnection
)
Function Get-ItemPreviousValue {
    Param (
        [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Url.')]
        [ValidateNotNullOrEmpty()]
        [String]
        $SiteUrl,

        [Parameter(Mandatory = $true, HelpMessage = 'Please insert the List Name.')]
        [ValidateNotNullOrEmpty()]
        [String]
        $ListName,

        [Parameter(Mandatory = $true, HelpMessage = 'Please insert the ID of the item to get the values from.')]
        [ValidateNotNullOrEmpty()]
        [String]
        $ItemID,

        [Parameter(Mandatory = $true, HelpMessage = 'Please insert the column internal name to retrieve the value for.')]
        [ValidateNotNullOrEmpty()]
        [String]
        $FieldInternalName,

        [Parameter(Mandatory = $false, HelpMessage = 'SharePoint Online Connection object')]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection
    )
    Try {

        # Connect to SharePoint Online Site if not already connected
        If (-not $SPOConnection) {
            $SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue
        }

        # Calculate the properties to add to the object based on the list name
        Switch ($ListName) {
            'AppCatalog' {
                $SwitchProperties = [ScriptBlock]::Create('
                   [PSCustomObject]@{
                        SPPKG        = $_.Values.FileLeafRef
                        AppDisplayName = $_.Values.Title
                    }
                ')
                Break
            }

            'DocumentList' {
                $SwitchProperties = [ScriptBlock]::Create('
                    [PSCustomObject]@{
                        TCM_DN = $_.Values.Title
                        Rev    = $_.Values.IssueIndex
                    }
                ')
                Break
            }

            Default {
                $SwitchProperties = $null
            }
        }

        # Get all versions of the item
        $AllVersions = Get-PnPListItemVersion -List $ListName -ID $ItemID -Connection $SPOConnection | ForEach-Object {
            # Create a custom object with the properties we want to return
            $VersionProperties = [PSCustomObject]@{
                VersionNumber             = [System.Version]$_.VersionLabel
                VersionID                 = $_.Id
                IsCurrentVersion          = $_.IsCurrentVersion
                VersionCreationDate       = $_.Created
                VersionCreatedByUser      = $_.Values.Editor.LookupValue
                VersionCreatedByUserEmail = $_.Values.Editor.Email
                ItemID                    = $_.Values.ID
                $FieldInternalName        = $_.Values.$FieldInternalName
            }

            # Add the properties from the switch statement if any
            If ($SwitchProperties) {
                $PropertiesToAdd = & $SwitchProperties

                ForEach ($Property in $PropertiesToAdd.PSObject.Properties) {
                    $VersionProperties | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value
                }
            }

            # Return the object to the pipeline
            $VersionProperties

        } | Sort-Object -Property VersionNumber -Descending

        $CurrentVersion = $AllVersions[0]
        $LatestValueVersion = $AllVersions | Where-Object { $_.$FieldInternalName -ne $CurrentVersion.$FieldInternalName } |
            Select-Object -First 1 -Property *, @{
                Name       = "$FieldInternalName - Property Change";
                Expression = { "$($_.$FieldInternalName) => $($CurrentVersion.$FieldInternalName)" }
            }
    }
    Catch {
        Throw
    }

    #Return $LatestValueVersion
    Return $LatestValueVersion
}

# Call the function
Get-ItemPreviousValue -ListName $ListName -SiteUrl $SiteUrl -ItemID $ItemID -FieldInternalName $FieldInternalName
