# ExportColumn=Title returns error
Param(
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [String]$ListName,

    [Parameter(Mandatory = $true)]
    [String]$ViewName,

    [Parameter(Mandatory = $true)]
    [ValidateSet('InternalName', 'DisplayName', 'Title')] # DisplayName or Title are the same
    [String]$ExportColumn,

    [String]$CSVExportFolder
)

<# Parameter samples

VDM Discipline:
https://tecnimont.sharepoint.com/sites/vdm_4305
Vendor Documents List
Building
DisplayName or InternalName

DDC
https://tecnimont.sharepoint.com/sites/4305DigitalDocumentsC
Client Document List
Documents To Send
DisplayName or InternalName

#>

If ('' -eq $CSVExportFolder) {
    $inputValue = Read-Host -Prompt 'CSVExportFolder'
    If ($inputValue -ne '') {
        $CSVExportFolder = $inputValue
        If (!(Test-Path $CSVExportFolder)) {
            Write-Host ('{0} is not a valid path' -f $CSVExportFolder) -ForegroundColor Yellow
            Exit
        }
    }
    Else {
        $CSVExportFolder = (Split-Path -Parent $MyInvocation.MyCommand.Path) + '\'
    }
}
Else {
    If (!(Test-Path $CSVExportFolder)) {
        Write-Host ('{0} is not a valid path' -f $CSVExportFolder) -ForegroundColor Yellow
        Exit
    }
}
If (!($CSVExportFolder.EndsWith('\'))) {
    $CSVExportFolder = $CSVExportFolder + '\'
}

# Function to convert a View CAML query to a Powershell FilterScript
Function Convert-ViewQueryToFilterScript {
    Param (
        [Parameter(Mandatory)]
        [String] $ViewQuery
    )

    Function Convert-CamlCondition {
        Param (
            [Parameter(Mandatory)]
            [String] $CamlCondition
        )

        $regexFieldRef = [regex]'<FieldRef Name="([^"]+)"\s?\/>'
        $regexValueType = [regex]'<Value Type="([^"]+)">(.+)<\/Value>'

        $fieldRefMatches = $regexFieldRef.Matches($CamlCondition)
        $valueTypeMatches = $regexValueType.Matches($CamlCondition)

        $fieldName = $fieldRefMatches[0].Groups[1].Value

        If ($CamlCondition.Contains('<IsNull>') -or $CamlCondition.Contains('<IsNotNull>')) {
            if ($CamlCondition.Contains('<IsNull>')) {
                $operator = '-eq'
            }
            else {
                $operator = '-ne'
            }
            $value = "''"
        }
        else {
            if ($CamlCondition.Contains('<Eq>')) {
                $operator = '-eq'
            }
            else {
                $operator = '-ne'
            }
            $valueType = $valueTypeMatches[0].Groups[1].Value
            $value = $valueTypeMatches[0].Groups[2].Value

            if ($valueType -eq 'Boolean') {
                if ($value -eq '1') {
                    $value = '$true'
                }
                else {
                    $value = '$false'
                }
            }
            else {
                $value = "`"$value`""
            }
        }

        return "`$_.$fieldName $operator $value"
    }

    Function Convert-CamlElement {
        param (
            [Parameter(Mandatory)]
            [string] $CamlElement
        )

        if ($CamlElement.Contains('<And>') -or $CamlElement.Contains('<Or>')) {
            if ($CamlElement.Contains('<And>')) {
                $operator = '-and'
            }
            else {
                $operator = '-or'
            }
            $innerElements = $CamlElement -replace '<And>|<Or>|<\/And>|<\/Or>', ''

            $conditions = [regex]::Matches($innerElements, '<(?:Eq|Neq|IsNull|IsNotNull)>.*?<\/(?:Eq|Neq|IsNull|IsNotNull)>') | ForEach-Object { $_.Value }
            $operands = $conditions | ForEach-Object { Convert-CamlCondition -CamlCondition $_ }

            return "($($operands -join " $operator "))"
        }
        else {
            return Convert-CamlCondition -CamlCondition $CamlElement
        }
    }

    $whereElement = [regex]::Match($ViewQuery, '<Where>.*<\/Where>').Value
    $whereElement = $whereElement -replace '<Where>|<\/Where>', ''

    $filterScriptText = '(& { ' + (Convert-CamlElement -CamlElement $whereElement) + ' })'
    $filterScript = [scriptblock]::Create($filterScriptText)

    Return $FilterScript
}

Function Convert-FilterScriptInternalToDisplayName {
    Param(
        [ScriptBlock]$FilterScript,
        [Array]$ViewFieldsColumns
    )

    $updatedFilterScript = $FilterScript.ToString()

    $ViewFieldsColumns | ForEach-Object {
        If ($FilterScript.ToString().contains("$($_.InternalName)")) {
            $DisplayName = ("'{0}'" -f $_.Title)
            $InternalName = $_.InternalName
            $updatedFilterScript = $updatedFilterScript.Replace($InternalName, $DisplayName)
        }
    }

    $UpdatedFilterScript = [scriptblock]::Create($updatedFilterScript)
    Return $UpdatedFilterScript
}

If ($ExportColumn -eq 'DisplayName') {
    $ExportColumn = 'Title'
}

# Compose CSV export file path
$CSVExportFilePath = $CSVExportFolder + ($SiteUrl.Split('/')[-1], $ListName, $ViewName, (Get-Date -Format 'dd-MM-yyyy-HH-mm-ss') -join '_') + '.csv'

# Connect to the SharePoint site
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Get the view
$ListView = Get-PnPView -List $ListName -Identity $ViewName -Includes ViewQuery, ViewFields, ListViewXml

# Convert CAML query to Powershell FilterScript
$FilterScript = Convert-ViewQueryToFilterScript -ViewQuery $ListView.ViewQuery

# Get all List Columns
$ViewFieldsColumns = Get-PnPField -List $ListName <#-ReturnTyped#> | Select-Object -Property InternalName, Title, TypeAsString, ReadOnlyField | Where-Object -FilterScript { $_.InternalName -in $ListView.ViewFields }

# Compose view fields
$ViewFields = '<View><ViewFields>'
$ViewFieldsColumns.InternalName | ForEach-Object {
    $ViewFields += ("<FieldRef Name='{0}'/>" -f $_)
}
$ViewFields += '</ViewFields></View>'

# Get all list items from the specific view
$CodeStartTime = Get-Date
Write-Host ('Starting at {0}' -f $CodeStartTime) -ForegroundColor Yellow
$AllListItems = Get-PnPListItem -List $ListName -Query $ViewFields -PageSize 5000 | ForEach-Object {
    $ListItem = $_
    $Item = New-Object PSObject
    Foreach ($Field in $ViewFieldsColumns) {

        <# Used for troubleshooting on a specific column
        If ($field.InternalName -eq 'VendorName_x003a_Site_x0020_Url') {
            Write-host "$($Field.InternalName)"
        }
        #Write-Host ('Processing Column {0}/{1}' -f $ViewFieldsColumns.InternalName.IndexOf($($Field.InternalName)), $ViewFieldsColumns.Count)
        #>

        If ($null -ne $ListItem["$($Field.InternalName)"]) {
            Switch ($Field.TypeAsString) {
                {
                    ($_ -eq 'Lookup') -or
                    ($_ -eq 'LookupMulti')
                } {
                    $ItemValue = ($ListItem["$($Field.InternalName)"] | ForEach-Object {
                            (
                                '{0}' -f
                                $_.LookupValue
                            )
                        }) -join ', '
                    Break
                }

                {
                    ($_ -eq 'User') -or
                    ($_ -eq 'UserMulti')
                } {
                    $ItemValue = ($ListItem["$($Field.InternalName)"] | ForEach-Object {
                            (
                                '{0} ({1})' -f
                                $_.LookupValue,
                                $_.Email
                            )
                        }) -join ', '
                    Break
                }

                default {
                    $ItemValue = $ListItem["$($Field.InternalName)"]
                }
            }
            $Item | Add-Member -MemberType NoteProperty -Name $Field.$($ExportColumn) -Value $ItemValue
        }
        Else {
            $Item | Add-Member -MemberType NoteProperty -Name $Field.$($ExportColumn) -Value ''
        }

    }
    $Item
}
$CodeEndTime = Get-Date
$CodeExecutionTime = New-TimeSpan -Start $CodeStartTime -End $CodeEndTime
Write-Host ('Execution time: {0}h {1}m {2}s' -f $($CodeExecutionTime.Hours), $($CodeExecutionTime.Minutes), $($CodeExecutionTime.Seconds)) -ForegroundColor Green

# If asked, update  the FilterScript to filter with DisplayName instead of InternalName
If ($ExportColumn -eq 'Title') {
    $FilterScript = Convert-FilterScriptInternalToDisplayName -FilterScript $FilterScript -ViewFieldsColumns $ViewFieldsColumns
}

# Filter List items as intended on specified view
$FilteredItems = $AllListItems | Where-Object -FilterScript $FilterScript

# Export to CSV
$FilteredItems | Export-Csv -Path $CSVExportFilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'