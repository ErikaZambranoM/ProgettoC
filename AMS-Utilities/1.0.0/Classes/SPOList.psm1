Class SPOList {
    ##### REGION: Public Class Properties #####

    <# ListDisplayName
        The name of the List to be loaded from SharePoint Online
    #>
    [ValidateNotNullOrEmpty()]
    [String]
    $ListDisplayName

    <# ListColumnsMapping
        A PSCustomObject with the mapping between the List columns and the columns to be used in the output.

        Example:
            $ListColumnsMapping = [PSCustomObject]@{
                TCM_DN = 'Title'
                Rev    = 'Revision Number'
            }
    #>
    [PSCustomObject]
    $ListColumnsMapping

    <# ListItems
        List Items loaded from the SharePoint Online List
    #>
    [PSCustomObject]
    $ListItems



    ##### REGION: Hidden Class Properties #####

    <# SPOConnection
        The SharePoint Online Connection object to be used to load the List (obtaind via Connect-PnPOnline)
    #>
    Hidden
    [ValidateNotNullOrEmpty()]
    [PnP.PowerShell.Commands.Base.PnPConnection]
    $SPOConnection

    <# ListColumns
        List Fields (Columns) loaded from the SharePoint Online List
    #>
    Hidden
    [PSCustomObject]
    $ListColumns

    <# MethodCalledFromAllowedContext
        Boolean property to check if a method has been called from an allowed context (e.g. from a method of the same class)
        This is used to avoid certain methods to be called from outside the class
    #>
    Hidden
    [Boolean]
    $MethodCalledFromAllowedContext = $false



    ##### REGION: Constructors #####

    <# Constructor 1 (2 arguments)
       Check if specified ListDisplayName List exists and create List object with specified SPOConnection
    #>
    SPOList (
        $ListDisplayName,
        $SPOConnection
    ) {
        Try {
            # Assign values to class properties
            $This.ListDisplayName = $ListDisplayName
            $This.SPOConnection = $SPOConnection

            # Check if List exists and throw exception if not
            $ListExists = Get-PnPList -Identity $This.ListDisplayName -Connection $This.SPOConnection -ErrorAction SilentlyContinue
            If ($null -eq $ListExists) {
                Throw ("[ERROR] List '{0}' not found in '{1}'{2}" -f $This.ListDisplayName, $This.SPOConnection.Url, "`n ")
            }
        }
        Catch {
            Throw
        }
    }



    ##### REGION: Public Methods #####

    <# Method GetColumns 1 (no arguments)
        Method to get all Fields (Columns) from a List
        If ListColumnsMapping property has been manually set, return only the mapped columns
    #>
    [System.Object[]]
    GetColumns() {
        Try {
            # Get all List Fields (Columns) and save them in ListColumns property
            $ListFields = Get-PnPField -List $This.ListDisplayName -Connection $This.SPOConnection |
                Select-Object -Property InternalName, Title, TypeAsString, ReadOnlyField, Context

            # Assign value to class properties
            $This.ListColumns = $ListFields

            # If ListColumnsMapping property has been set, return only the mapped columns
            If ($null -ne $This.ListColumnsMapping) {
                $ListFields = $This.GetColumns($This.ListColumnsMapping)
            }
            Return $ListFields
        }
        Catch {
            Throw
        }
    }

    <# Method GetColumns 2 (1 argument)
        Method to get only Fields (Columns) specified (in PSCustomObject $ListColumnsMapping) from a List
    #>
    [System.Object[]]
    GetColumns(
        <# ListColumnsMapping
            A PSCustomObject with the mapping between the List columns and the columns to be used in the output.

            Example:
            $ListColumnsMapping = [PSCustomObject]@{
                TCM_DN = 'Title'
                Rev = 'Revision Number'
            }
        #>
        [PSCustomObject]
        $ListColumnsMapping
    ) {
        Try {
            # Validate $ListColumnsMapping argument data type
            If ($ListColumnsMapping.GetType().Name -ne 'PSCustomObject') {
                Throw (
                    "[ERROR] ListColumnsMapping argument must be a PSCustomObject.{0}Example:{0}`$ListColumnsMappingSample = [PSCustomObject]@{2}{0}{1}TCM_DN = 'Title'{0}{1}Rev = 'Revision Number'{0}{3}" -f
                    "`n",
                    "`t",
                    '{',
                    '}'
                )
            }

            # If needed, assign values to class properties
            If ($null -eq $This.ListColumnsMapping) {
                $This.ListColumnsMapping = $ListColumnsMapping
            }
            If ($null -eq $This.ListColumns) {
                $This.ListColumns = $This.GetColumns()
            }

            # Initialize variables
            $SelectedListFields = @()
            $DuplicatedListColumnNames = @{}

            # $ListColumnsMapping validation
            # Check if $ListColumnsMapping is empty
            If ($This.ListColumnsMapping.PSObject.Properties.Value.Count -eq 0) {
                Throw ("[ERROR] No List column names found in Mapped List Columns.`n ")
            }
            # Check if $ListColumnsMapping contains duplicated column names
            Else {
                ForEach ($ColumnName in $This.ListColumnsMapping.PSObject.Properties.Value) {
                    If ($DuplicatedListColumnNames.ContainsKey($ColumnName)) {
                        $DuplicatedListColumnNames[$ColumnName]++
                    }
                    Else {
                        $DuplicatedListColumnNames[$ColumnName] = 1
                    }
                }
                $MappedListColumnDuplicates = $DuplicatedListColumnNames.Keys | Where-Object -FilterScript { $DuplicatedListColumnNames[$_] -gt 1 }
                If ($MappedListColumnDuplicates) {
                    Throw ("[ERROR] Duplicated List column names found in Mapped List Columns:`n{0}`n " -f ($MappedListColumnDuplicates -join "`n"))
                }
            }

            # Get all List column names mapped on $ListColumnsMapping to create ordered object
            $OrderedColumnNamesObject = New-Object System.Collections.Specialized.OrderedDictionary
            $This.ListColumnsMapping.PSObject.Properties | ForEach-Object {
                $OrderedColumnNamesObject[$_.Name] = $_.Value
            }
            $MappedListColumnsNames = $OrderedColumnNamesObject.Keys | ForEach-Object {
                $ColumnAssignedName = $_
                $ListColumnName = $This.ListColumnsMapping."$($ColumnAssignedName)"
                $ListColumnName
            }

            # Create a Splatting object to be used in Select-Object to add a new column with the mapped column name
            $MappedColumnSplatting = @{
                Name       = 'MappedColumnName'
                Expression = {
                    $This.ListColumnsMapping |
                        Get-Member -MemberType NoteProperty |
                            Where-Object -FilterScript {
                                $This.ListColumnsMapping."$($_.Name)" -eq $ListColumnName
                            } |
                                Select-Object -ExpandProperty Name
                }
            }

            # Check if each column name in $MappedListColumnsNames exists in $This.ListColumns filtering via InternalName. If not, filter via Title.
            ForEach ($ListColumnName in $MappedListColumnsNames) {
                # Check by InternalName
                [Array]$ListColumn = $This.ListColumns | Where-Object { $_.InternalName -eq $ListColumnName }
                If ($null -ne $ListColumn) {
                    $SelectedListFields += $ListColumn | Select-Object -Property *, $MappedColumnSplatting
                }
                # Check by DisplayName
                Else {
                    [Array]$ListColumn = $This.ListColumns | Where-Object { $_.Title -eq $ListColumnName }
                    If ($null -ne $ListColumn -and $ListColumn.Count -eq 1) {
                        $SelectedListFields += $ListColumn | Select-Object -Property *, $MappedColumnSplatting
                    }
                    Else {
                        Throw ("[ERROR] Column '{0}' not found in List '{1}'{2}" -f $ListColumnName, $This.ListDisplayName, "`n ")
                    }
                }
            }

            Return $SelectedListFields
        }
        Catch {
            Throw
        }
    }

    <# Hidden Method GetAllItemsFromFields (1 argument)
        Method to get all items from a List and loop through each item to create a new PSObject with all Fields (Columns) and values
        This method is hidden because it is used internally by other methods.
    #>
    [PSCustomObject]
    Hidden GetAllItemsFromFields(
        [System.Object[]]
        $ListFields,

        <#PSObjectColumnName
            The type of column name to be used in the returned PSObject SPOList.
            Valid values for manual use are:
                'DisplayName' or 'Display'
                'InternalName' or 'Internal'

            'CustomName' is only for Class internal use.
        #>
        [String]
        $PSObjectColumnName
    ) {
        # Ensure the method is called from an allowed context
        Trap {
            # Reset MethodCalledFromAllowedContext to false to avoid calling the hidden method GetAllItemsFromFields
            $This.MethodCalledFromAllowedContext = $false
        }

        Try {
            # Stops the code if this method is called directly
            If (-not $This.MethodCalledFromAllowedContext) {
                Throw 'This method cannot be called directly.'
            }

            # Validate ListFields argument
            If ($ListFields[0].Context.ApplicationName -ne 'SharePoint PnP PowerShell Library' ) {
                Throw 'ListFields argument is not valid.'
            }

            # Set the column name type for the PSObject List
            Switch ($PSObjectColumnName) {
                # 'DisplayName' or 'Display'
                {
                    $_ -eq 'DisplayName' -or
                    $_ -eq 'Display'
                } {
                    $PSObjectColumnName = 'Title'
                    Break
                }

                # 'InternalName' or 'Internal'
                {
                    $_ -eq 'InternalName' -or $_ -eq 'Internal'
                } {
                    $PSObjectColumnName = 'InternalName'
                    Break
                }

                # 'CustomName'. For Class internal use only.
                {
                    $_ -eq 'CustomName' -and
                    ($ListFields | Get-Member -MemberType NoteProperty -Name 'MappedColumnName')
                } {
                    $PSObjectColumnName = 'MappedColumnName'
                    Break
                }

                # Default to error since no other values are allowed
                Default {
                    Throw (
                        "'{0}' is not a valid column name for the PSObject List to be returned!{1}Valid names are:{1}{2}{1}{3}" -f
                        $PSObjectColumnName,
                        "`n ",
                        "- 'DisplayName' or 'Display'",
                        "- 'InternalName' or 'Internal"
                    )
                }
            }

            # Get all items from the List and loop through each item to create a new PSObject with all Fields (Columns) and values
            $AllListItems = Get-PnPListItem -List $This.ListDisplayName -Connection $This.SPOConnection -PageSize 5000 | ForEach-Object {

                # Temporary variable to store the current Item
                $ListItem = $_
                $Item = New-Object PSObject

                # Loop through all Fields (Columns) and get each value for the current Item
                Foreach ($Field in $ListFields) {
                    # If Field value is null assign empty string to the current Item
                    If ($null -eq $ListItem["$($Field.InternalName)"]) {
                        $Item | Add-Member -MemberType NoteProperty -Name $($Field.$($PSObjectColumnName)) -Value ''
                    }
                    Else
                    { # Otherwise, extract correct value based on Field type and assign it to the current Item
                        Switch ($Field.TypeAsString) {

                            # Lookup and LookupMulti fields
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

                            # User and UserMulti fields
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

                            # Default value type that don't need any special treatment
                            Default {
                                $ItemValue = $ListItem["$($Field.InternalName)"]
                            }
                        }
                        $Item | Add-Member -MemberType NoteProperty -Name $($Field.$($PSObjectColumnName)) -Value $ItemValue
                    }
                }
                $Item
            }

            # Assign values to class properties
            $This.ListItems = $AllListItems
            $This.MethodCalledFromAllowedContext = $false

            Return $AllListItems
        }
        Catch {
            # Reset MethodCalledFromAllowedContext to false to avoid calling the hidden method GetAllItemsFromFields
            $This.MethodCalledFromAllowedContext = $false
            Throw
        }
    }

    <# Method GetAllItems 1 (no arguments)
        Method to get all Fields' (Columns) values of all Items from a List
    #>
    [PSCustomObject]
    GetAllItems(
        <#PSObjectColumnName
            The type of column name to be used in the returned PSObject SPOList.
            Valid values for manual use are:
                'DisplayName' or 'Display'
                'InternalName' or 'Internal'

            'CustomName' is only for Class internal use.
        #>
        [String]
        $PSObjectColumnName
    ) {

        Try {
            # Get All Fields (Columns) from the List
            $ListFields = $This.GetColumns()

            # Permit to call the hidden method GetAllItemsFromFields
            $This.MethodCalledFromAllowedContext = $true

            # Get all Items from the List
            $Items = $This.GetAllItemsFromFields($ListFields, $PSObjectColumnName)

            Return $Items
        }
        Catch {
            Throw
        }
    }

    <# Method GetAllItems 2 (1 argument)
        Method to get Fields (Columns) specified (in PSCustomObject) of all Items from a List
    #>
    [PSCustomObject]
    GetAllItems (
        <# ListColumnsMapping
            A PSCustomObject with the mapping between the List columns and the columns to be used in the output.

            Example:
            $ListColumnsMapping = [PSCustomObject]@{
                TCM_DN = 'Title'
                Rev = 'Revision Number'
            }
        #>
        $ListColumnsMapping
    ) {
        Try {

            # Get specified Fields (Columns) from the List
            $ListFields = $This.GetColumns($ListColumnsMapping)

            # Permit to call the hidden method GetAllItemsFromFields
            $This.MethodCalledFromAllowedContext = $true

            # Get all Items from the List
            $Items = $This.GetAllItemsFromFields($ListFields, 'CustomName')

            Return $Items
        }
        Catch {
            Throw
        }
    }
}
