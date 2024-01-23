<#
 ToDo:
    Provide a CSV with list of FTAEvolutionName
    A Document Library and a Registry will be created with this name. All spaces in provided names will be deleted.
#>

Param (
    [Parameter(Mandatory = $true)]
    [String]
    $ProjectCode,

    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [AllowEmptyString()]
    [String]
    $CSVPath
)


$SiteURL = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
If (-not $CSVPath) {
    $CSVPath = "$PSScriptRoot\Setup_FTALists.csv"
}
$arrayFTALists = Import-Csv $CSVPath -Delimiter ';'

Function createFTAEvolution() {
    param
    (
        $Ctx,
        $Lists,
        $ListTitle,
        $ListURL
    )

    Write-Host "Create List: $($ListTitle) - $($ListURL)"

    if (!($Lists.Title -contains $ListTitle)) {
        $lci = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $lci.Title = $ListTitle
        $lci.Url = $ListURL
        $lci.Description = ''
        $lci.TemplateType = 101
        $list = $Ctx.web.lists.add($lci)
        $Ctx.load($list)

        #send the request containing all operations to the server
        $Ctx.executeQuery()
        Write-Host "info: Created $($ListTitle)" -ForegroundColor green
    }
    else {
        Write-Host -f Yellow "List '$ListTitle' already exists!"
    }

    #Aggiungo il Content Type
    addContentType -context $Ctx -ListName $ListTitle -CTypeName 'FTAEvolutionDocument Content Type'
    removeContentType -context $Ctx -ListName $ListTitle -ContentTypeName 'Document'

    #Recupero View standard e la modifico
    $ViewName = 'All Docs'
    $Query = "<OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy>"

    $ViewFields = 'FSObjType',
    'LinkFilename',
    'ClientAction',
    'Title',
    'Discipline',
    'ClientDiscipline',
    'CommentDueDate',
    'Modified',
    'Editor'

    CreateView -Ctx $Ctx -ListName $ListTitle -ViewFields $ViewFields -ViewName $ViewName -ViewQuery $Query

    #Creo Registry
    $RegTitle = "$($ListTitle)Registry"
    $RegURL = "Lists/$($ListURL)Registry"
    if (!($Lists.Title -contains $RegTitle)) {
        $lci = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $lci.Title = $RegTitle
        $lci.Url = $RegURL
        $lci.Description = ''
        $lci.TemplateType = 100
        $list = $Ctx.web.lists.add($lci)
        $Ctx.load($list)

        #send the request containing all operations to the server
        $Ctx.executeQuery()
        Write-Host "info: Created $($RegTitle)" -ForegroundColor green
    }
    else {
        Write-Host -f Yellow "List '$RegTitle' already exists!"
    }

    addContentType -context $Ctx -ListName $RegTitle -CTypeName 'FTAEvolutionRegistry Content Type'
    removeContentType -context $Ctx -ListName $RegTitle -ContentTypeName 'Item'

    #Recupero View standard e la modifico
    $ViewName = 'All Items'
    $Query = "<OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>"

    $ViewFields = 'Transmittal_x0020_Number',
    'FTARevision',
    'Discipline',
    'ClientDiscipline',
    'FTAStatus',
    'TransmittalDate',
    'TransmittalUser',
    'CommentDueDate',
    'ValidationUser',
    'ValidationDate',
    'LastClientTransmittalUser',
    'LastClientTransmittalDate',
    'ID'

    UpdateDefaultView -Ctx $Ctx -ListName $RegTitle -ViewFields $ViewFields -ViewName $ViewName -Query $Query

    setListPermissions -context $Ctx -ListName $RegTitle

    $JsonFormat = @"
        {
        "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
        "elmType": "div",
        "attributes": {
          "class": "=if(Number([`$LastClientTransmittalDate]) == 0, if([`$CommentDueDate] <= @now, 'sp-field-severity--severeWarning', if(@now >= ([`$CommentDueDate] - (86400000*2)), 'sp-field-severity--warning', 'sp-field-severity--good')),'sp-field-severity--good')+ ' ms-fontColor-neutralSecondary'"
        },
        "children": [
          {
            "elmType": "span",
            "style": {
              "display": "inline-block",
              "padding": "0 4px"
            },
            "txtContent": "@currentField"
          },
          {
            "elmType": "span",
            "style": {
              "display": "inline-block",
              "padding": "0 4px"
            },
            "attributes": {
              "iconName": "=if(Number([`$LastClientTransmittalDate]) == 0, if([`$CommentDueDate] <= @now, 'Error', if(@now >= ([`$CommentDueDate] - (86400000*2)), 'Warning', '')),'')",
              "title": "=if(Number([`$LastClientTransmittalDate]) == 0, if([`$CommentDueDate] <= @now, 'Expired from ' + floor((Number(@now)-Number([`$CommentDueDate]))/86400000) + ' days', if(@now >= ([`$CommentDueDate] - (86400000*2)), 'Expiring', @currentField)),@currentField)"
            }
          }
        ]
      }
"@

    formatColumn -Ctx $Ctx -ListName $RegTitle -FieldName 'CommentDueDate' -JsonFormat $JsonFormat

    $JsonFormat = @"
    {
        "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
        "elmType": "a",
        "txtContent": "@currentField",
        "attributes": {
          "target": "_blank",
          "href": "=@currentWeb + '/' + '$($RegURL)' + '/' + [`$Transmittal_x0020_Number] + '-' + [`$FTARevision]"
        }
      }
"@

    formatColumn -Ctx $Ctx -ListName $RegTitle -FieldName 'Transmittal_x0020_Number' -JsonFormat $JsonFormat
}

Function addContentType {
    param (
        $context,
        $ListName,
        $CTypeName
    )

    Try {
        #Get the List
        $List = $context.web.Lists.GetByTitle($ListName)
        $context.Load($List)
        $context.ExecuteQuery()

        #Enable managemnt of content type in list - if its not enabled already
        If ($List.ContentTypesEnabled -ne $True) {
            $List.ContentTypesEnabled = $True
            $List.Update()
            $context.ExecuteQuery()
            Write-Host 'Content Types Enabled in the List!' -f Yellow
        }

        #Get all existing content types of the list
        $ListContentTypes = $List.ContentTypes
        $context.Load($ListContentTypes)

        #Get the content type to Add to list
        $ContentTypeColl = $context.Web.ContentTypes
        $context.Load($ContentTypeColl)
        $context.ExecuteQuery()

        #Check if the content type exists in the site
        $CTypeToAdd = $ContentTypeColl | Where-Object { $_.Name -eq $CTypeName }
        If ($Null -eq $CTypeToAdd) {
            Write-Host "Content Type '$CTypeName' doesn't exists!" -f Yellow
            Return
        }

        #Check if content type added to the list already
        $ListContentType = $ListContentTypes | Where-Object { $_.Name -eq $CTypeName }
        If ($Null -ne $ListContentType) {
            Write-Host "Content type '$CTypeName' already exists in the List!" -ForegroundColor Yellow
        }
        else {
            #Add content Type to the list or library
            $AddedCtype = $List.ContentTypes.AddExistingContentType($CTypeToAdd)
            $context.ExecuteQuery()

            Write-Host "Content Type '$CTypeName' Added to '$ListName' Successfully!" -ForegroundColor Green
        }
    }
    Catch {
        Write-Host -f Red 'Error Adding Content Type to the List!' $_.Exception.Message
    }
}

Function removeContentType() {
    param (
        $context,
        $ListName,
        $ContentTypeName
    )

    Try {
        #Get the List
        $List = $context.Web.Lists.GetByTitle($ListName)
        $context.Load($List)

        #Get the content type from list
        $ContentTypeColl = $List.ContentTypes
        $context.Load($ContentTypeColl)
        $context.ExecuteQuery()

        #Get the content type to remove
        $CTypeToRemove = $ContentTypeColl | Where-Object { $_.Name -eq $ContentTypeName }
        If ($CTypeToRemove -ne $Null) {
            #Remove content type from list
            $CTypeToRemove.DeleteObject()
            $context.ExecuteQuery()

            Write-Host "Content Type '$ContentTypeName' Removed From '$ListName'" -f Green
        }
        else {
            Write-Host "Content Type '$ContentTypeName' doesn't exist in '$ListName'" -f Yellow
            Return
        }
    }
    Catch {
        Write-Host -f Red 'Error Removing Content Type from List!' $_.Exception.Message
    }
}

Function setListPermissions {
    param(
        $context,
        $ListName
    )

    $PermissionLevel = 'MT Contributors'
    $GroupNames = 'FTAEvolutionClientOperators',
    'FTAEvolutionClientValidators',
    'FTAEvolutionTCMOperators',
    'FTAEvolutionTCMValidators'

    try {
        #Get the web and List
        $Web = $context.Web
        $List = $Web.Lists.GetByTitle($ListName)

        $groups = $Web.SiteGroups
        $context.Load($groups)
        $context.executeQuery()

        #Break Permission inheritence - keep existing list permissions & Item level permissions
        $List.BreakRoleInheritance($True, $True)
        $context.ExecuteQuery()
        Write-Host -f Yellow 'Permission inheritance broken...'

        $context.Load($List.RoleAssignments)
        $context.ExecuteQuery()

        #Remove permissions to groups
        foreach ($RoleAssignment in $List.RoleAssignments) {
            $context.Load($RoleAssignment.Member)
            $context.ExecuteQuery()
            $RoleAssignmentName = $RoleAssignment.Member.Title

            if ($GroupNames -contains $RoleAssignmentName) {
                $group = $groups.GetByName($RoleAssignmentName)
                $List.RoleAssignments.GetByPrincipal($group).DeleteObject()
                $context.ExecuteQuery()

                Write-Host -f Green "Permissions for Group $RoleAssignmentName Deleted Successfully!" $_.Exception.Message
            }
        }

        #Assign permissions to all the other groups
        foreach ($GroupName in $GroupNames) {
            $group = $groups.GetByName($GroupName)
            $RoleDef = $context.web.RoleDefinitions.GetByName($PermissionLevel)
            $RoleDefBind = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
            $RoleDefBind.Add($RoleDef)
            $context.Load($List.RoleAssignments.Add($group, $RoleDefBind))

            $context.ExecuteQuery()

            Write-Host -f Green "Permission level $PermissionLevel added to User Group $GroupName in the list $ListName"
        }
    }
    Catch {
        Write-Host -f Red 'Error setting Permissions!' $_.Exception.Message
    }
}

Function CreateView() {
    param
    (
        [Parameter(Mandatory = $true)] $Ctx,
        [Parameter(Mandatory = $true)] $ListName,
        [Parameter(Mandatory = $true)] [String[]] $ViewFields,
        [Parameter(Mandatory = $true)] $ViewName,
        [Parameter(Mandatory = $true)] [string] $ViewQuery
    )

    $ItemLimit = '50'
    $IsDefaultView = $True

    Try {
        #Get the List
        $List = $Ctx.Web.Lists.GetByTitle($ListName)
        $Ctx.Load($List)
        $Ctx.ExecuteQuery()

        #Check if the View exists in list already
        $ViewColl = $List.Views
        $Ctx.Load($ViewColl)
        $Ctx.ExecuteQuery()
        $NewView = $ViewColl | Where-Object { ($_.Title -eq $ViewName) }
        if ($NULL -ne $NewView) {
            Write-Host "View '$ViewName' already exists in the List!" -f Yellow
        }
        else {
            $ViewCreationInfo = New-Object Microsoft.SharePoint.Client.ViewCreationInformation
            $ViewCreationInfo.Title = $ViewName
            $ViewCreationInfo.Query = $ViewQuery
            $ViewCreationInfo.RowLimit = $ItemLimit
            $ViewCreationInfo.ViewFields = $Viewfields
            $ViewCreationInfo.SetAsDefaultView = $IsDefaultView
            $ViewCreationInfo.Paged = 'TRUE'

            #sharepoint online powershell create view
            $NewView = $List.Views.Add($ViewCreationInfo)
            $Ctx.ExecuteQuery()

            Write-Host 'New View Added to the List Successfully!' -ForegroundColor Green
        }
    }
    Catch {
        Write-Host -f Red 'Error Adding View to List!' $_.Exception.Message
    }
}

Function UpdateDefaultView() {
    param
    (
        $Ctx,
        $ListName,
        $ViewFields,
        $ViewName,
        $Query
    )

    #$ViewName="All Documents"

    Try {
        #Get the List
        $List = $Ctx.Web.Lists.GetByTitle($ListName)
        $Ctx.Load($List)
        $Ctx.ExecuteQuery()

        #Get the default view
        $defaultView = $List.Views.getByTitle($ViewName)
        $Ctx.Load($defaultView)
        $Ctx.ExecuteQuery()

        #Check if view exists
        if ($defaultView -eq $Null) { Write-Host "View doesn't exists!" -f Red; return }

        #Update the View
        $defaultView.ViewQuery = $Query
        $defaultView.Update()
        $Ctx.ExecuteQuery()

        #Add fields to the view
        $defaultView = $List.Views.getByTitle($ViewName)
        $Ctx.Load($defaultView)
        $Ctx.ExecuteQuery()
        foreach ($field in $ViewFields) {
            $defaultView.ViewFields.Add($field)
        }
        $defaultView.Update()
        $Ctx.ExecuteQuery()

        Write-Host 'New fields Added to Default View Successfully!' -ForegroundColor Green
    }
    Catch {
        Write-Host -f Red 'Error Adding Fields to Default View!' $_.Exception.Message
    }
}

Function formatColumn {
    param
    (
        $Ctx,
        $ListName,
        $FieldName,
        $JsonFormat
    )

    #Get the List
    $List = $Ctx.Web.Lists.GetByTitle($ListName)

    #Get the Field
    $Field = $List.Fields.GetByInternalNameOrTitle($FieldName)
    $Ctx.Load($Field)
    $Ctx.ExecuteQuery()

    #Apply Column Formatting to the field
    $Field.CustomFormatter = $JsonFormat
    $Field.Update()
    $Ctx.ExecuteQuery()
    Write-Host 'Column has been formatted' -ForegroundColor green
}


# Connect to site
Try {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue
}
Catch {
    Write-Host ('Error while trying to connect to Site "{0}"' -f $SiteUrl) -ForegroundColor Red -BackgroundColor Yellow
}

#$Cred = Get-Credential
#$Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName, $Cred.Password)

Try {
    #$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    #$Ctx.Credentials = $Cred

    $Ctx = Get-PnPContext

    #Get All Lists
    $Lists = $Ctx.Web.Lists
    $Ctx.Load($Lists)
    $Ctx.ExecuteQuery()

    foreach ($ftaList in $arrayFTALists) {
        #Start-Job -ScriptBlock ${Function:createFTAEvolution} -ArgumentList $Ctx, $ftaList.ListTitle, $ftaList.ListURL | Wait-Job | Receive-Job
        $FTAEvolutionName = $ftaList.FTAEvolutionName.Replace(' ', '')
        createFTAEvolution -Ctx $Ctx -Lists $Lists -ListTitle $FTAEvolutionName -ListURL $FTAEvolutionName
    }
}
Catch {
    Write-Host -f Red 'Error Connection!' $_.Exception.Message
}