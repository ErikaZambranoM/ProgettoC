<# Prerequisites

  PnP Powershell module

  Date fields on CSV must be in format mm/dd/yyyy
  Format IssueIndex as text (otherwise revisions like '01' become '1')

#>

<# ToDo

  Add log option:
      - Log: No update needed
        Reminder for trim check missing:
          don't trim list value while checking if values is already updated (it would not correct the trim);
      - Update DateTime

Check duplicates in both KeyColumns and updateColumns
Check if KeyMappingColumn is ReadOnly before TrimUpdate

DONE: Hide set-pnpitem output

  Process bar and counter
  Add total exec time

  Add -BulkUpdate parameter
    check how to adapt functions to work for bot Single and BulkUpdate

  Mapped Column Name Check:
      $internalName = (Get-PnPField -List $ListName -Identity $FieldName).InternalName
        possible conflict: this command check for internal name first
          If internal name is found it won't search for a displayname
        All fields should be extracted, then filtered for both DisplayName and InternalName equal to mapped column
          if $result $null, column not found
          if $result = 1, ok
          if $result -gt 1, too may columns found; return error and exit

  Choice prompt -
    Provide details about current run with U
    Add -WhatIf parameter
    Check colors for U

  Check true/false column

  Notify how many duplicate rows

    Add -CSVSkipRows parameter
    Add -Remediation parameter
      provide single remediation scripts
    Config CSV validation


    ~ UI:
      - "Connect to Sharepoint" Button

      - TextField&DropDown (startsEmpty) per $SiteUrl
          Search Site Button that get the list of available Sites
      - TextField&DropDown for available or custom $ListName
          Search Site Button that get the list of available Lists

      - "Connect To List" Button
      - "Update from CSV" Button (Disabled)
      - "Connect To List".OnClick(
          Get all Columns names (both internalName and DisplayName) from List
          Show empty ListViewTable with a ColumnFilterTextField.Disabled filter on every column and a disabled "Search" button
            show an empty EnableColumnFilterTickBox beside every ColumnFilterTextField
            show an empty ShowColumnTickBox beside every ColumnFilterTextField to include the Column in the output without using a filter on it
          "Update from CSV".Enabled
        )

      - ListViewTable.Filter.IsNotNull(Enable "Search" button)
      - "Search".OnClick(
          Return on ListViewTable as many as possible items from list with given filter
            Show only Column wher (ShowColumnTickBox = true || EnableColumnFilterTickBox = true)
        )

      - "Update from CSV".OnClick(
          File picker for CSV/Excel
          Show file Columns for item[0] with most populated values (as sample)
            show an empty KeyColumnTickBox beside every Column
            show an empty UpdateColumnTickBox beside every Column
            show an empty ShowOnlyTickBox beside every Column
        )

      - "Edit item" Button (Disabled)
      - ListViewTable.Item.OnClick("Edit item".Enabled)
      - "Edit item".OnClick(
          Show ItemViewTable with items' editable and readOnly fields
          Apply button (Disabled)
          Item.Propert.OnChange(ApplyButton.Enable)
      )

      Other:
        - Add option to run automatic Site/List Search on start
        - Edit in Grid view button on ListViewTable?
        - CSV/Excel TopCount
#>

<#
Param(
  [Parameter(Mandatory=$true)]
  [string]$SiteUrl, #URL del sito

	[Parameter(Mandatory=$true)]
  [String]$ListName, #URL relativo o nome della lista/document library

	[String]$FieldName, #Nome Colonna normale o interno (opzionale)

	#[Switch]$System #System Update (opzionale)
)
#>

##### START - Variables to edit #####

# Site url
#$SiteUrl = 'https://tecnimont.sharepoint.com/sites/DDWave2'
$SiteUrl = 'https://tecnimont.sharepoint.com/sites/43U4DigitalDocuments'

# List Display Name
$ListName = 'DocumentList'

# Name of the CSV file containing values to be updated on Sharepoint List or Document Library.
# Set the value to $null to let user pick file from dialog.
#$CSVPath = 'C:\Users\ST-442\OneDrive\Desktop\WorkDesk\AMS-Utilities\Update\SPListFromCSV\DDWave2_Test.csv' # $null
$CSVPath = 'C:\Users\ST-442\Downloads\INC0926749.csv' # $null

<# Top Count
  Limit the number of CSV rows processed (mainly for test purpose)
  Set to $null if all rows need to be processed
#>
$TopCount = 1

<# KeyColumnsMapping
  Set one or more columns to be used as Key Columns on each side.
  Tested MatchingOperator:
    - Eq
    - Contains
#>
$KeyColumnsMapping = @(
  <#
    Copy this commented block to add another Key Column
      @{
        'ListColumnInternalName' = 'IssueIndex'
        'CSVColumnName' = 'Revision'
        'MatchingOperator' = 'Eq'
      }
  #>
  @{
    'ListColumnInternalName' = 'Title'
    'CSVColumnName'          = 'TCM_DN'
    'MatchingOperator'       = 'Eq'
  },
  @{
    'ListColumnInternalName' = 'Issue Index'
    'CSVColumnName'          = 'Rev'
    'MatchingOperator'       = 'Eq'
  }
)

# Set the columns to be mapped for being updated on List
$UpdateColumnsMapping = @(
  <#
      Copy this commented block to add another column to update
      @{
        'ListColumnInternalName' = 'ClientDepartmentDescription'
        'CSVColumnName' = 'ClientDepartmentDescription NEW'
      }
  #>
  @{
    'ListColumnInternalName' = 'DocumentStatus' #ReasonForIssue
    'CSVColumnName'          = 'DocumentStatus'
  }
  <#@{
    'ListColumnInternalName' = 'LastTransmittalDate' #ReasonForIssue
    'CSVColumnName'          = 'TrnDate'
  }#>
)

<# CSVDuplicatesRemediation and ListDuplicatesRemediation

    In case of duplicates on CSV:
    Skip (Skip duplicate items and log them as 'Duplicated CSV row' on the report)
    IncludeDuplicates (Process all csv rows)
    Unique (Only process the first occurrence on the CSV. Processed row gets logged as 'First occurence of duplicated CSV value')

    In case of duplicates on List:
    Skip (Skip duplicate items and log them as 'More then one row found on list matching the Key Column' on the report)
    IncludeDuplicates (Every item matching the Key Column will be updated with csv values)
    Unique (Only update the first occurrence of found Items on list. Processed row gets logged as 'First List item occurence updated')

#>
$CSVDuplicatesRemediation = 'Skip'
$ListDuplicatesRemediation = 'Skip'

<# NullNotEmptyValues
  Add here every column data type which expects null value instead of empty string when are being cleared
    check Boolean
    check UserGroup
#>
$NullNotEmptyValues = @(
  'DateTime'
)

#####  END - Variables to edit  #####



##### START - Functions region  #####

# Function to create a CSV log starting from imported CSV, adding columns needed for processing and columns with List items values before editing
Function New-CSVReportFromImportedCSV
{
  [CmdletBinding()]
  Param(

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        If (Test-Path $_ -PathType Leaf)
        {
          Return $True
        }
        Return $False
      })]
    [String]$CSVPath,

    [Parameter(Mandatory = $false)]
    [Int]$TopCount,

    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [String]$ListName,

    [Parameter(Mandatory = $true)]
    [Object]$ImportedCSV,

    [Parameter(Mandatory = $true)]
    [Object]$KeyColumnsMapping,

    [Parameter(Mandatory = $true)]
    [Object]$UpdateColumnsMapping,

    [Parameter(Mandatory = $true)]
    [String]$CSVDuplicatesRemediation
  )

  # Get the CSV file information
  $CSVItem = Get-Item -Path $CSVPath

  # Check if the $TopCount variable is defined and retrieve the valute to include in $CSVLogName
  If (0 -ne $TopCount)
  {
    # If $TopCount is defined, use it to generate the TopCountString
    $TopCountString = ('Top_{0}' -f ($TopCount))
  }
  Else
  {
    # If $TopCount is not defined, use the count of imported rows to generate the TopCountString
    $TopCountString = ('All_{0}' -f $ImportedCSV.Count)
  }

  # Generate the log file name
  $CSVLogName = ('Log {0}({1}) - {2}({3}) - {4}{5}') -f
  $SiteUrl.Split('/')[-1],
  $ListName.Replace(' ', ''),
  $CSVItem.BaseName,
  $TopCountString,
    (Get-Date -Format 'dd-MM-yyyy_HH_mm_ss'),
  $CSVItem.Extension

  # Generate the log file path
  $Global:CSVLogPath = Join-Path -Path $CSVItem.DirectoryName -ChildPath $CSVLogName

  # Create a new log file
  $null = New-Item -Path $Global:CSVLogPath -ItemType File

  # Write a message indicating that the log file was created
  Write-Host ''
  Write-Host ('Created log CSV file:{0}"{1}"' -f "`n", ($Global:CSVLogPath)) -ForegroundColor Black -BackgroundColor Yellow
  Write-Host ''

  # Creating unique id column by merging Key Columns values
  $Global:DuplicateCheckerColumnName = 'KC_'

  # Initialize the string builder object
  $StringBuilder = New-Object System.Text.StringBuilder

  # Loop through each key column in the KeyColumnsMapping object
  ForEach ($CSVKeyColumn in $KeyColumnsMapping.CSVColumnName)
  {
    # Replace spaces and square brackets with underscores and hyphens, respectively
    $CSVKeyColumn = $CSVKeyColumn -replace ' ', '_' -replace '\[', '-' -replace '\]' , '-'

    # Append the modified key column to the string builder
    $StringBuilder.Append($CSVKeyColumn) | Out-Null

    # Check if the current key column is not the last one
    If ($KeyColumnsMapping.CSVColumnName.IndexOf($CSVKeyColumn) -ne (@($KeyColumnsMapping.CSVColumnName).Length - 1))
    {
      # Append an underscore to the string builder if it's not the last key column
      $StringBuilder.Append('_') | Out-Null
    }
  }

  # Concatenate the final string to the DuplicateCheckerColumnName variable
  $Global:DuplicateCheckerColumnName += $StringBuilder.ToString()

  # Get a list of all properties in $ImportedCSV that are not in $KeyColumnsMapping.CSVColumnName and in $UpdateColumnsMapping.CSVColumnName
  $Global:CSVColumnsToUpdate = ($ImportedCSV | Get-Member -MemberType NoteProperty | Select-Object -Property Name | Where-Object { $_.Name -notin $KeyColumnsMapping.CSVColumnName -and $_.Name -in $UpdateColumnsMapping.CSVColumnName }).Name

  # Get a list of all properties in $ImportedCSV that are not in $Global:CSVColumnsToUpdate and not in $KeyColumnsMapping.CSVColumnName
  $ColumnsNotToReport = ($ImportedCSV | Get-Member -MemberType NoteProperty | Select-Object -Property Name | Where-Object { $_.Name -notin $Global:CSVColumnsToUpdate -and $_.Name -notin $KeyColumnsMapping.CSVColumnName }).Name

  # Add the values for the created unique id column to the $ImportedCSV object
  $ImprovedCSV = $ImportedCSV | Select-Object -Property *, @{
    L = "$Global:DuplicateCheckerColumnName"
    E = {

      # Initialize the value for the unique id column
      $DuplicateCheckerColumnValue = 'KC_'

      # Loop through each key column in the KeyColumnsMapping object
      ForEach ($CSVKeyColumn in $KeyColumnsMapping.CSVColumnName)
      {

        # Add the current key column value to the unique id column value
        $DuplicateCheckerColumnValue += $_.$CSVKeyColumn.Trim()

        # Check if the current key column is not the last one
        If ($KeyColumnsMapping.CSVColumnName.IndexOf($CSVKeyColumn) -ne (@($KeyColumnsMapping.CSVColumnName).Length - 1))
        {
          # Add an underscore to the unique id column value if it's not the last key column
          $DuplicateCheckerColumnValue += '_'
        }
      }

      # Return the final value for the unique id column
      $DuplicateCheckerColumnValue
    }
  },
  @{
    L = 'ListItemID'
    E = { $null }
  },
  @{
    L = 'CSVRowID'
    E = { $(@($ImportedCSV).IndexOf($_) + 2) }
  } -ExcludeProperty $ColumnsNotToReport

  # Adding a backup column for values to update
  # The backup column will store the original value before it is updated
  ForEach ($Column in $Global:CSVColumnsToUpdate)
  {
    # Add a new property to the $ImprovedCSV object
    # The name of the property is "OLD_<ColumnName>"
    # The value is set to $null
    $ImprovedCSV | Add-Member -MemberType NoteProperty -Name ('{0}' -f ('OLD_' + $Column)) -Value $Null
  }

  # Separate unique rows from the duplicate ones
  $DuplicateFreeCSV = @()

  # Get all unique rows in the $ImprovedCSV object by grouping on the unique id column
  # Filter the groups where the count of rows is equal to 1
  # Select only the group of rows and add it to the $DuplicateFreeCSV array
  $DuplicateFreeCSV += $ImprovedCSV | Group-Object -Property "$Global:DuplicateCheckerColumnName" | Where-Object { $_.Count -eq 1 } | Select-Object -ExpandProperty Group

  # Get all duplicate rows in the $ImprovedCSV object by grouping on the unique id column
  # Filter the groups where the count of rows is greater than 1
  # Select only the group of rows and store it in the $DuplicateRows variable
  $DuplicateRows = $ImprovedCSV | Group-Object -Property "$Global:DuplicateCheckerColumnName" | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group

  # Filter duplicates based on the value of $CSVDuplicatesRemediation
  $Global:RowsToProcess = @()
  Switch ($CSVDuplicatesRemediation)
  {
    # If the value is "Skip", log the duplicated rows and process only the unique ones
    'Skip'
    {
      # Log the duplicated rows to the specified CSV log file
      $DuplicateRows | Group-Object -Property "$Global:DuplicateCheckerColumnName" | ForEach-Object {
        $DuplicateItem = $_ | Select-Object -ExpandProperty Group
        $DuplicateItem | Select-Object -Property *,
        @{
          L = 'Log'
          E = { ('Skipped: Duplicated CSV row (CSV Row IDs: {0})' -f ($DuplicateItem.CSVRowID -join ', ')) }
        } | Export-Csv -Path $Global:CSVLogPath -NoTypeInformation -Delimiter ';' -Append
      }

      # Set the rows to process as the unique ones
      $Global:RowsToProcess = $DuplicateFreeCSV
    }

    # If the value is "IncludeDuplicates", process all the rows (including duplicates)
    'IncludeDuplicates'
    {
      $Global:RowsToProcess = $ImprovedCSV
    }

    # If the value is "Unique", process only the first of duplicate rows
    'Unique'
    {
      # Create an array to store the unique duplicated rows
      $UniqueDuplicateRows = @()

      # Loop through the duplicated rows and add only the unique ones to the array
      ForEach ($DuplicateRow in $DuplicateRows)
      {
        If ($DuplicateRow."$Global:DuplicateCheckerColumnName" -notin $UniqueDuplicateRows."$Global:DuplicateCheckerColumnName")
        {
          $UniqueDuplicateRows += $DuplicateRow

          $DuplicateItem = $Null
          $DuplicateRows |
            Where-Object -FilterScript { $_."$Global:DuplicateCheckerColumnName" -eq $DuplicateRow."$Global:DuplicateCheckerColumnName" } |
              Group-Object -Property "$Global:DuplicateCheckerColumnName" | ForEach-Object {
                $DuplicateItem = $_ | Select-Object -ExpandProperty Group
                $DuplicateItem | Select-Object -Skip 1 -Property *,
                @{
                  L = 'Log'
                  E = { ('Skipped: Processed only first CSV Row occurrence ({0}). Ignored CSV Row IDs: ({1})' -f $DuplicateItem.CSVRowID[0], ($DuplicateItem.CSVRowID[1..$DuplicateItem.Count] -join ', ')) }
                }
              } | Export-Csv -Path $Global:CSVLogPath -NoTypeInformation -Delimiter ';' -Append
        }
      }

      # Set the rows to process as the unique duplicated rows and the unique ones, sorted by CSVRowID
      $Global:RowsToProcess = ($DuplicateFreeCSV + $UniqueDuplicateRows) | Sort-Object -Property CSVRowID
    }

    # If the value is not any of the expected values, show an error message and exit the script
    Default
    {
      Write-Host ('Exception not handled, wrong CSVDuplicatesRemediation ({0})' -f $CSVDuplicatesRemediation) -ForegroundColor Red -BackgroundColor Yellow
      Exit
    }
  }
}

# Function to write on CSV log all the details about processed row
Function Write-ToCSVReport
{
  Param (
    $ProcessedRowDetails,

    [String]$LogMessage,

    [Switch]$OnlyFirstListItem
  )

  If ($OnlyFirstListItem)
  {
    ForEach ($Column in $Global:CSVColumnsToUpdate)
    {
      $Row."$('OLD_' + $Column)" = $ListItems[0].$(@($UpdateColumnsMapping_ListColumnInternalName)[$UpdateColumnsMapping.CSVColumnName.IndexOf($Column)]) -join ', '
    }
  }
  Else
  {
    ForEach ($Column in $Global:CSVColumnsToUpdate)
    {
      $Row."$('OLD_' + $Column)" = $ListItems.$(@($UpdateColumnsMapping_ListColumnInternalName)[$UpdateColumnsMapping.CSVColumnName.IndexOf($Column)]) -join ', '
    }
  }

  $Row.ListItemID = $ListItems.ID -join ', '
  If ($null -ne $Row.Log)
  {
    $Row.PsObject.Members.Remove('Log')
  }
  $Row | Add-Member -NotePropertyName 'Log' -NotePropertyValue $LogMessage
  $Row | Export-Csv -Path $Global:CSVLogPath -NoTypeInformation -Delimiter ';' -Append
}

# Function to filter a row from imported CSV using criteria expressed in $KeyColumnsMapping
Function Find-ListItemsFromCSV
{
  Param (
    [Parameter(Mandatory = $true)]
    $ListItems,

    [Parameter(Mandatory = $true)]
    [System.Collections.ArrayList]$FilterMapping,

    [Parameter(Mandatory = $true)]
    [PSCustomObject]$CSVRow
  )

  ForEach ($ListColumn in $FilterMapping)
  {
    If ($ListColumn.MatchingOperator)
    {
      $MatchingOperator = $ListColumn.MatchingOperator.ToLower()
    }
    Else
    {
      $MatchingOperator = $null
    }

    Switch ($MatchingOperator)
    {
      'eq'
      {
        #$ListItems = $ListItems | Where-Object -FilterScript {$_.($ListColumn.ListColumnInternalName).Trim() -eq $Row.($ListColumn.CSVColumnName).Trim()}
        $ListItems = $ListItems | Where-Object "$(($ListColumn.ListColumnInternalName).Trim())" -EQ "$($Row.($ListColumn.CSVColumnName).Trim())"
      }

      $null
      {
        Write-Host ('No matching operator set for columns mapping: {0} / {1}' -f $($ListColumn.ListColumnInternalName, $($ListColumn.CSVColumnName))) -BackgroundColor Yellow -ForegroundColor Red
        Exit
      }

      Default
      {
        Write-Host ('Matching operator "{0}" not handled' -f $($ListColumn.MatchingOperator)) -BackgroundColor Yellow -ForegroundColor Red
        Exit
      }
    }
  }

  Return $ListItems
}

# Function to create new List item values to be updated on List
Function New-UpdatedListItemValuesObject
{
  Param (
    [Array]$UpdateColumnsMapping,
    [Array]$ListColumnsTypes,
    [Array]$NullNotEmptyValues,
    [PSCustomObject]$CSVRow
  )

  $NewItemValues = New-Object -TypeName HashTable

  ForEach ($ListColumn in $UpdateColumnsMapping_ListColumnInternalName)
  {
    If
    (
      '' -eq $($CSVRow.$(@($UpdateColumnsMapping.CSVColumnName)[$UpdateColumnsMapping_ListColumnInternalName.IndexOf($ListColumn)])) -and
      $($ListColumnsTypes[$ListColumnsTypes.InternalName.IndexOf($ListColumn)].TypeAsString) -in $NullNotEmptyValues
    )
    {
      $NewItemValues += @{
        $ListColumn = $null
      }
    }
    Else
    {
      $NewItemValues += @{
        $ListColumn = $($CSVRow.$(@($UpdateColumnsMapping.CSVColumnName)[$UpdateColumnsMapping_ListColumnInternalName.IndexOf($ListColumn)])).Trim()
      }
    }
  }

  Return $NewItemValues
}

Function Set-ListItems
{

  Param (
    [Parameter(Mandatory = $true)]
    [Object]$ListItems,

    [Parameter(Mandatory = $true)]
    [String]$ListDuplicatesRemediation,

    [Parameter(Mandatory = $true)]
    [Object]$UpdatedListItemValues,

    [Parameter(Mandatory = $true)]
    [String]$ListName,

    [Parameter(Mandatory = $true)]
    [Object]$CSVRow,

    [Parameter(Mandatory = $true)]
    [Object]$KeyColumnsMapping
  )

  If (@($ListItems).length -gt 1)
  {

    Switch ($ListDuplicatesRemediation)
    {
      'Skip'
      {
        Write-Host ('Skipped: More then one row found on list matching the Key Columns (List Item IDs: {0})' -f ($ListItems.ID -join ', ')) -ForegroundColor Yellow
        Write-ToCSVReport -LogMessage ('Skipped: More then one row found on list matching the Key Columns (List Item IDs: {0})' -f ($ListItems.ID -join ', ')) -ProcessedRowDetails $Row
      }

      'IncludeDuplicates'
      {
        Write-Host 'More then one item found on List, processing all of them:' -ForegroundColor Yellow
        ForEach ($Item in $ListItems)
        {
          Write-Host ('Processing item with ID: {0}' -f $Item.ID)

          Try
          {
            Set-PnPListItem -List "$ListName" -Identity $($Item.ID) -Values $UpdatedListItemValues | Out-Null
            Write-Host ('Updated item {0}' -f $ListItems.ID) -ForegroundColor Green
            Write-ToCSVReport -LogMessage 'All items updated' -ProcessedRowDetails $Row
          }
          Catch
          {
            $CatchedError = ($_ | Out-String).Trim()
            Write-Host ('Error while trying to edit one of the list items found (Item ID: {0})' -f $Item.ID) -BackgroundColor Red -ForegroundColor Yellow
            Write-Host $CatchedError -ForegroundColor Red -BackgroundColor Yellow

            Write-ToCSVReport -LogMessage ('Error while trying to edit one of the list items found (Item ID: {0}):{1}{2}' -f $Item.ID, "`n", $CatchedError) -ProcessedRowDetails $Row
          }
        }
      }

      'Unique'
      {
        Write-Host ('More then one item found on List, only processing first occurrence: {0}' -f $ListItems[0].ID) -ForegroundColor Yellow
        Write-Host ('Ignored Item IDs: {0}' -f ($ListItems.ID[1..($ListItems.Count - 1)] -join ', '))

        #Update first List Item occurrence
        Try
        {
          Set-PnPListItem -List "$ListName" -Identity $($ListItems[0].ID) -Values $UpdatedListItemValues | Out-Null
          Write-Host ('Updated item {0}' -f $ListItems[0].ID) -ForegroundColor Green
          Write-ToCSVReport -LogMessage ('First List item occurence ({0}) updated. Ignored Item IDs: {1}' -f $ListItems[0].ID, ($ListItems.ID[1..($ListItems.Count - 1)] -join ', ')) -ProcessedRowDetails $Row -OnlyFirstListItem
        }
        Catch
        {
          $CatchedError = ($_ | Out-String).Trim()
          Write-Host 'Error while trying to edit item' -BackgroundColor Yellow -ForegroundColor Red
          Write-Host $CatchedError -ForegroundColor Red -BackgroundColor Yellow

          Write-ToCSVReport -LogMessage ('Error while trying to edit item: {0}' -f $CatchedError) -ProcessedRowDetails $Row
        }
      }

      Default
      {
        Write-Host ('Exception not handled, wrong ListDuplicatesRemediation ({0})' -f $ListDuplicatesRemediation) -ForegroundColor Red -BackgroundColor Yellow
        Exit
      }
    }
  }
  Else
  {
    Write-Host '1 item found on List:' -ForegroundColor Yellow
    Write-Host ('Processing item {0}' -f $ListItems.ID)

    #Update Item
    Try
    {
      # Verify if values for $KeyColumnsMapping on List need to be trimmed
      ForEach ($ListColumn in $KeyColumnsMapping_ListColumnInternalName)
      {
        $CSVRowKeyValueTrimmed = $($Row.$(@($KeyColumnsMapping.CSVColumnName)[$KeyColumnsMapping_ListColumnInternalName.IndexOf($ListColumn)])).Trim()
        $ListItemCurrentKeyValue = $ListItems.$(@($KeyColumnsMapping_ListColumnInternalName)[$KeyColumnsMapping_ListColumnInternalName.IndexOf($ListColumn)])

        If ($CSVRowKeyValueTrimmed -ne $ListItemCurrentKeyValue)
        {
          Write-Host ('List column "{0}" for item {1} has a trailing space to fix:{2}{3}"{4}" will be corrected in "{5}"' -f $ListColumn, $ListItems.ID, "`n", "`t", $ListItemCurrentKeyValue, $CSVRowKeyValueTrimmed) -ForegroundColor Yellow
          $UpdatedListItemValues += @{
            $ListColumn = $CSVRowKeyValueTrimmed
          }
        }
      }

      Set-PnPListItem -List "$ListName" -Identity $($ListItems.ID) -Values $UpdatedListItemValues | Out-Null
      Write-Host ('Updated item {0}' -f $ListItems.ID) -ForegroundColor Green
      Write-ToCSVReport -LogMessage 'Item updated' -ProcessedRowDetails $Row

    }
    Catch
    {
      $CatchedError = ($_ | Out-String).Trim()
      Write-Host 'Error while trying to edit item' -BackgroundColor Red -ForegroundColor Yellow
      Write-Host $CatchedError -ForegroundColor Red -BackgroundColor Yellow

      Write-ToCSVReport -LogMessage ('Error: {0}' -f $CatchedError) -ProcessedRowDetails $Row

    }
  }
}

Function Import-CSVAndValidate
{
  param (
    [Parameter(Mandatory = $true)]
    [String]$CSVPath,

    [Parameter(Mandatory = $true)]
    [Object]$KeyColumnsMapping,

    [Parameter(Mandatory = $true)]
    [Object]$UpdateColumnsMapping,

    [Parameter(Mandatory = $false)]
    [Int]$TopCount
  )

  If (!(Test-Path -Path $CSVPath))
  {
    Write-Host ('CSV file not found at "{0}"!' -f $CSVPath) -ForegroundColor Red -BackgroundColor Yellow
    Exit
  }

  # Import CSV File
  $ImportedCSV = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8

  # Get CSV Column names
  $CSVColumns = $ImportedCSV[0].PSObject.Properties.Name.Trim()

  # Import the CSV again with trimmed Column names
  $ImportedCSV = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8 -Header $CSVColumns | Select-Object -Skip 1

  If (0 -ne $TopCount)
  {
    $ImportedCSV = $ImportedCSV | Select-Object -Property * -First $TopCount
  }

  # Validate CSV columns
  ForEach ($Column in [Array]$KeyColumnsMapping.CSVColumnName + [Array]$UpdateColumnsMapping.CSVColumnName)
  {
    If ($CSVColumns -notcontains $Column)
    {
      Write-Host ('"{0}" mapped column is not found in CSV file!' -f $Column) -ForegroundColor Yellow -BackgroundColor Red
      Write-Host ('Available columns on CSV are:{0}{1}' -f "`n", ($CSVColumns -replace '^|$', '"' | Out-String).Trim()) -ForegroundColor Red -BackgroundColor Yellow
      Exit
    }
  }
  Return $ImportedCSV
}

#####  END - Functions region   #####



##### START - Input validation  #####

# Connect to site
Try
{
  Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue
}
Catch
{
  Write-Host ('Error while trying to connect to Site "{0}"' -f $SiteUrl) -ForegroundColor Red -BackgroundColor Yellow
  Exit
}

# Check if List exists
Try
{
  If ($null -eq (Get-PnPList -Identity "$ListName"))
  {
    Write-Host ('List "{0}" does not exists on Site "{1}"' -f $ListName, $SiteUrl) -ForegroundColor Red -BackgroundColor Yellow
    Exit
  }
}
Catch
{
  Write-Host ('Error while trying to get List "{0}" on Site "{1}"' -f $ListName, $SiteUrl) -ForegroundColor Red -BackgroundColor Yellow
  Exit
}


# Get all List Columns
[Array]$ListColumnsTypes = Get-PnPField -List $ListName <#-ReturnTyped#> | Select-Object -Property InternalName, Title, TypeAsString, ReadOnlyField
[Array]$UpdateColumnsMapping_ListColumnInternalName = $UpdateColumnsMapping.ListColumnInternalName
[Array]$KeyColumnsMapping_ListColumnInternalName = $KeyColumnsMapping.ListColumnInternalName
# Filter from both Column mapping variables ($UpdateColumnsMapping and $KeyColumnsMapping) fields that don't exist on List
$UnexistingListColumns = ($UpdateColumnsMapping_ListColumnInternalName + $KeyColumnsMapping_ListColumnInternalName) |
  Where-Object -FilterScript { $_ -notin $ListColumnsTypes.InternalName } |
    Select-Object -Unique |
      Select-Object -Property @{
        L = 'List Column InternalName'
        E = { $_ }
      },
      @{
        L = 'Mapping Variable to check'
        E = {
          If ($UpdateColumnsMapping_ListColumnInternalName.Contains("$_") -and $KeyColumnsMapping_ListColumnInternalName.Contains("$_"))
          {
            'KeyColumnsMapping + UpdateColumnsMapping'
          }
          ElseIf ($UpdateColumnsMapping_ListColumnInternalName.Contains("$_"))
          {
            'UpdateColumnsMapping'
          }
          ElseIf ($KeyColumnsMapping_ListColumnInternalName.Contains("$_"))
          {
            'KeyColumnsMapping'
          }
        }
      }

# If there is an attempt to write on a non-existant field, exit.
If ($UnexistingListColumns.Count -gt 0)
{
  Write-Host ('Following mapped fields are not present on List:{0}{1}' -f "`n`n", ($UnexistingListColumns | Out-String).Trim()) -ForegroundColor Red -BackgroundColor Yellow
  Exit
}

# Filter from $UpdateColumnsMapping fields that are ReadOnly on List
$TargetedReadOnlyColumns = $UpdateColumnsMapping_ListColumnInternalName |
  Where-Object -FilterScript {
    $ListColumnsTypes[$ListColumnsTypes.InternalName.IndexOf("$_")].ReadOnlyField -eq $true -and
    $_ -notin $UnexistingListColumns.'List Column InternalName'
  }

# If there is an attempt to write on a ReadOnly field, exit.
If ($TargetedReadOnlyColumns.Count -gt 0)
{
  Write-Host ('Following mapped fields cannot be updated on List because they are ReadOnly:{0}{1}' -f "`n", $TargetedReadOnlyColumns) -ForegroundColor Red -BackgroundColor Yellow
  Exit
}



# Import Configs
<#

Get script execution path to search for Config folder, if not found, Exit

$ImportedCSV = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8
If ($null -ne $TopCount) {
  $ImportedCSV = $ImportedCSV | Select-Object -Property * -First $TopCount
}

Later, return in pre-pause Settings output which config files are being used
#>



#####  END - Input validation   #####


##### START - CSV log creation  #####

Try
{
  $ImportedCSV = Import-CSVAndValidate -CSVPath $CSVPath -KeyColumnsMapping $KeyColumnsMapping -UpdateColumnsMapping $UpdateColumnsMapping -TopCount $TopCount

  # Create a hash table for the parameters to be passed to the New-CSVReportFromImportedCSV function
  $SplattingParameters_CSVReportFromImportedCSV = @{
    CSVPath                  = $CSVPath
    TopCount                 = $TopCount
    SiteUrl                  = $SiteUrl
    ListName                 = $ListName
    ImportedCSV              = $ImportedCSV
    KeyColumnsMapping        = $KeyColumnsMapping
    UpdateColumnsMapping     = $UpdateColumnsMapping
    CSVDuplicatesRemediation = $CSVDuplicatesRemediation
  }

  # Call the New-CSVReportFromImportedCSV function and pass the parameters through splatting
  New-CSVReportFromImportedCSV @SplattingParameters_CSVReportFromImportedCSV
}
Catch
{
  Write-Host ('Error while importing CSV or creating Report file') -ForegroundColor Yellow -BackgroundColor Red
  Write-Host ($_ | Out-String).Trim() -ForegroundColor Red -BackgroundColor Yellow
  Exit
}

#####  END - CSV log creation   #####



<# Config files handling
  Provide a FilesPath.config with a semicolon separated mapping for every list (ListInternalName;FilePath;IgnoreItemFieldRule;FileFilter(Name/Extension/Attribute(also based on process stage)))
    es. DocumentList;DocumentsPath;StagingDocumentsPath

  Provide a $($ListName)_MappedListChanges.config file for each List that requires a change in other Lists or Document Libraries (on folders) when a change in $ListName occurs.
    Inside the .config file, list all Lists or Document Libraries (on folders) which need to be updated when an attribute in common with $ListName is edited.
      For Document Libraries, also provide filters on file (ListType(List or Document Library);ListName;FolderList)

  $MappedListChangesConfigFile = Get-ChildItem where name -eq $($ListName)_MappedListChanges.config
  If ($MappedListChangesConfigFile.Count -gt 1){
    stop
  } else {
    Get fields foreach List or Document Library in .config file
    If (there are fields in common) {
      $MappedListChangesRequired = $true
      #Map changes to apply other Lists or Document Libraries
    }
  }

  Later, just after Set-PnpItem
    If ($True -eq MappedListChangesRequired){
      Apply mapped changes
    }

#>



##### START - Resuming settings and asking confirm to start #####

Write-Host ' ~ Settings ~ ' -ForegroundColor Black -BackgroundColor White
Write-Host ''

Write-Host ('Site URL: {0} ' -f $SiteUrl) -ForegroundColor Green -BackgroundColor DarkGray
Write-Host ('List Name: {0} ' -f $ListName) -ForegroundColor Green -BackgroundColor DarkGray
Write-Host ''

Write-Host ('CSV Path: {0} ' -f $CSVPath) -ForegroundColor Magenta -BackgroundColor DarkGray
Write-Host ('TopCount: {0} ' -f $TopCount) -ForegroundColor Magenta -BackgroundColor DarkGray
Write-Host ''

Write-Host ('CSVDuplicatesRemediation: {0} ' -f $CSVDuplicatesRemediation) -ForegroundColor DarkCyan -BackgroundColor DarkGray
Write-Host ('ListDuplicatesRemediation: {0} ' -f $ListDuplicatesRemediation) -ForegroundColor DarkCyan -BackgroundColor DarkGray
Write-Host ''

Write-Host ('Key Columns Mapping ') -ForegroundColor Black -BackgroundColor Cyan
Write-Host ''

Write-Host ($KeyColumnsMapping | Select-Object -Property @{
    L = 'CSV Column'
    E = { $_.CSVColumnName }
  },
  @{
    L = 'Matching Operator'
    E = { $_.MatchingOperator }
  },
  @{
    L = 'List Column'
    E = { $_.ListColumnInternalName }
  } | Out-String).Trim() -ForegroundColor Black -BackgroundColor Cyan
Write-Host ''

Write-Host ('Update Columns Mapping ') -ForegroundColor Black -BackgroundColor Blue
Write-Host ''

Write-Host ($UpdateColumnsMapping | Select-Object -Property @{
    L = 'CSV Column'
    E = { $_.CSVColumnName }
  },
  @{
    L = 'List Column'
    E = { $_.ListColumnInternalName }
  } | Out-String).Trim() -ForegroundColor Black -BackgroundColor Blue

Write-Host ('{0}Found {1} {2} on CSV to be processed' -f
  "`n",
  $Global:RowsToProcess.Count,
  $(
    If ($Global:RowsToProcess.Count -gt 1)
    {
      'rows'
    }
    Else
    {
      'row'
    }
  )
) -ForegroundColor Black -BackgroundColor Green

If (0 -eq $Global:RowsToProcess.Count)
{
  Write-Host 'No valid rows to update, closing.' -ForegroundColor Black -BackgroundColor Yellow
  Exit
}

# Ask confirm to user before proceeding
$Title = 'Do you confirm?'
$Info = 'Please check above settings before confirming!'

$UpdateListChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Update List', (
  'Update List{0}Update items in "{1}" with values from the CSV.{2}' -f
  "`n",
  $ListName,
  "`n`n"
)
$SimulateChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Simulate', (
  'Simulate{0}Only fill the CSV log file with the changes that would be made with normal execution without actually making any changes.{1}' -f
  "`n",
  "`n`n"
)
$CancelChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Cancel', (
  'Cancel{0}Terminate the process and rename the CSV log prepending "Canceled_" to its name.{1}' -f
  "`n",
  "`n`n"
)

$Options = [System.Management.Automation.Host.ChoiceDescription[]] @($UpdateListChoice, $SimulateChoice, $CancelChoice)
[int]$DefaultChoice = 2
$ChoicePrompt = $host.UI.PromptForChoice($Title, $Info, $Options, $DefaultChoice)

Switch ($ChoicePrompt)
{
  0
  {
    Write-Host "`nStarting List Update Process" -ForegroundColor DarkCyan -BackgroundColor Magenta
  }

  1
  {
    Write-Host 'WiP - Exiting' -ForegroundColor Green
    Exit
  }

  2
  {
    Write-Host 'Process canceled!' -ForegroundColor Red -BackgroundColor Yellow
    Exit
  }
}

#####  END - Resuming settings and asking confirm to start  #####



##### START - Getting list items and applying modifies #####

# Compose Query
$Query = "<View><ViewFields><FieldRef Name='ID'/>"
ForEach ($Column in $KeyColumnsMapping_ListColumnInternalName + $UpdateColumnsMapping_ListColumnInternalName)
{
  $Query += ("<FieldRef Name='{0}'/>" -f $Column)
}
$Query += '</ViewFields></View>'

# Get all List items
$AllListItems = Get-PnPListItem -List $listName -Query $Query -PageSize 5000 |
  ForEach-Object {
    $Properties = @{}
    $Properties['ID'] = $_['ID']
    ForEach ($Column in $KeyColumnsMapping_ListColumnInternalName + $UpdateColumnsMapping_ListColumnInternalName)
    {
      $Properties[$Column] = $_[$Column]
    }

    [PSCustomObject]$Properties
  }

# Process rows
ForEach ($Row in $Global:RowsToProcess)
{
  $CatchedError = $null

  # Progress bar for each row
  $Progress = [Math]::Round(($Global:RowsToProcess.IndexOf($Row) / $Global:RowsToProcess.Count) * 100)
  Write-Progress -Activity "Processing row $($Global:RowsToProcess.IndexOf($Row) + 1) of $($Global:RowsToProcess.Count)" -Status "Progress: $Progress%" -PercentComplete $Progress
  Write-Host ("`nProcessing {0}: {1}" -f ($Global:DuplicateCheckerColumnName), $Row.$Global:DuplicateCheckerColumnName)

  # Filtering CSV items on list
  Try
  {
    $ListItems = Find-ListItemsFromCSV -ListItems $AllListItems -FilterMapping $KeyColumnsMapping -CSVRow $Row
  }
  Catch
  {
    Write-Host ('Error while executing function: Find-ListItemsFromCSV') -ForegroundColor Yellow -BackgroundColor Red
    Write-Host ($_ | Out-String).Trim() -ForegroundColor Red -BackgroundColor Yellow
    Exit
  }

  # Execute and log result
  If ($null -ne $ListItems)
  {

    # Create object with new List items values from CSV
    Try
    {
      $UpdatedListItemValues = New-UpdatedListItemValuesObject -UpdateColumnsMapping $UpdateColumnsMapping -ListColumnsTypes $ListColumnsTypes -NullNotEmptyValues $NullNotEmptyValues -CSVRow $Row
    }
    Catch
    {
      Write-Host ('Error while executing function: New-UpdatedListItemValuesObject') -ForegroundColor Yellow -BackgroundColor Red
      Write-Host ($_ | Out-String).Trim() -ForegroundColor Red -BackgroundColor Yellow
      Exit
    }

    # Update List items
    Set-ListItems -ListItems $ListItems -ListDuplicatesRemediation $ListDuplicatesRemediation -UpdatedListItemValues $UpdatedListItemValues -ListName $ListName -CSVRow $Row -KeyColumnsMapping $KeyColumnsMapping
  }
  Else
  {
    Write-Host ('{0} not found' -f $Row.$Global:DuplicateCheckerColumnName) -BackgroundColor Yellow -ForegroundColor Red
    Write-ToCSVReport -LogMessage 'Item not found on List' -ProcessedRowDetails $Row
  }
}
Write-Progress -Activity 'Update Completed' -Completed



# Disconnect
Disconnect-PnPOnline -Verbose

#####  END - Getting list items and applying modifies  #####