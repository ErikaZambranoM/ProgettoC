# When Override parameter is needed, run this script from command line
# !Controllare esistenza colonne prima di elminare view esistente e ricrearla
# ! Se ci sono colonne mancanti, aggiungerle automaticamente (con possibilità di elimarle in seguito)

param (
    [Parameter(Mandatory = $True)][String]$SourceSiteURL,
    [Parameter(Mandatory = $True)][String]$DestinationSiteURL,
    [Parameter(Mandatory = $True)][String]$SourceListName,
    [Parameter(Mandatory = $True)][String]$DestinationListName
)
Function Copy-View {
    param (
        [parameter (Mandatory = $true)]
        $SourceSiteConnection,

        [parameter (Mandatory = $true)]
        $DestinationSiteConnection,

        [parameter (Mandatory = $true)]
        [string]$SourceListName,

        [parameter (Mandatory = $true)]
        [string]$DestinationListName,

        [parameter (Mandatory = $true)]
        [string]$ViewName
    )

    Try {

        #Get the Source View
        $View = Get-PnPView -List $SourceListName -Identity $ViewName -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit -Connection $SourceSiteConnection | Where-Object -FilterScript { $_.Hidden -ne $True }

        If ($null -eq $View) {
            Write-Host "View $ViewName not found"
            Exit
        }

        #Get Properties of the source View
        $ViewProperties = @{

            'List'         = $DestinationListName #'Test View Copy - Client Document List'
            'Title'        = $ViewName
            'Paged'        = $View.Paged
            'Personal'     = $View.PersonalView
            'Query'        = $View.ViewQuery
            'RowLimit'     = $View.RowLimit
            'SetAsDefault' = $View.DefaultView
            'Fields'       = @($View.ViewFields)
            'ViewType'     = $View.ViewType
            'Aggregations' = $View.Aggregations

        }

        #Create a New View
        $DestinationView = Get-PnPView -List $DestinationListName -Identity $ViewName -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit -Connection $DestinationSiteConnection -ErrorAction SilentlyContinue | Where-Object -FilterScript { $_.Hidden -ne $True }
        If ($null -eq $DestinationView) {
            Add-PnPView @ViewProperties -Connection $DestinationSiteConnection | Out-Null
            Write-Host ('View "{0}" was missing and has now been created' -f $ViewName) -ForegroundColor Green
        }
        Else {
            If ($true -eq $Override) {
                Write-Host ('View "{0}" already exists. Proceeding with deleting it and then copying it' -f $ViewName) -ForegroundColor Yellow
                Remove-PnPView -Identity $ViewName -List $DestinationListName -Force -Connection $DestinationSiteConnection | Out-Null
                Add-PnPView @ViewProperties -Connection $DestinationSiteConnection | Out-Null
                Write-Host ('Existing view "{0}" has been deleted from destination site and then copied from source site' -f $ViewName) -ForegroundColor Green
            }
            Else {
                Write-Host ('View "{0}" already exists and will not be copied. Use Override parameter to force the copy' -f $ViewName) -ForegroundColor Yellow
            }
        }
    }
    Catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

try {
    $ViewName = Read-Host -Prompt 'View Name (Case sensitive)'
    if ($ViewName -eq '') {
        $ViewName = $null
        Write-Host 'Mode: Copy All' -ForegroundColor Red
    }

    # Ask if to use Override parameter
    Do { $Override = Read-Host -Prompt 'Override (true or false)' }
    While
    (
    ($Override.ToLower() -notin ('true', 'false'))
    )
    $Override = [System.Convert]::ToBoolean($Override)



    $SrcSiteConn = Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ReturnConnection
    $DestSiteConn = Connect-PnPOnline -Url $DestinationSiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ReturnConnection

    # If $VienaName is not specified, copy all views
    If ($null -eq $ViewName -or '' -eq $ViewName) {

        # Ask confirm to user before proceeding
        $Title = 'No ViewName Specified'
        $Info = 'If you proceed, all new Views will be copied'
        $ProceedCopyMsg = $null
        If ($true -eq $Override) {
            $Info = $Info + ' and existing Views will be overwritten'
            $ProceedCopyMsg = "`nOverride parameter has been used, existing Views will be overwritten!"
        }

        $ProceedChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Proceed', (
            'Copy Views{0}Copy all Views of List "{1}" from site "{2}" to List "{3}" of site "{4}".{0}{0}' -f
            "`n",
            $SourceListName,
            $SourceSiteUrl,
            $DestinationListName,
            $DestinationSiteUrl,
            $ProceedCopyMsg
        )
        $CancelChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Cancel', (
            'Cancel{0}Terminate the process without any change{1}' -f
            "`n",
            "`n`n"
        )

        $Options = [System.Management.Automation.Host.ChoiceDescription[]] @($ProceedChoice, $CancelChoice)
        [int]$DefaultChoice = 1
        $ChoicePrompt = $host.UI.PromptForChoice($Title, $Info, $Options, $DefaultChoice)

        Switch ($ChoicePrompt) {
            0 {
                Write-Host "`nStarting Views Copy Process" -ForegroundColor DarkCyan -BackgroundColor Magenta
            }

            1 {
                Write-Host 'Process canceled!' -ForegroundColor Red -BackgroundColor Yellow
                Exit
            }
        }

        $Views = Get-PnPView -List $SourceListName -Connection $SrcSiteConn -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit | Where-Object -FilterScript { $_.Hidden -ne $True }
        ForEach ($View in $Views) {
            Copy-View -SourceSiteConnection $SrcSiteConn -DestinationSiteConnection $DestSiteConn -SourceListName $SourceListName -DestinationListName $DestinationListName -ViewName $($View.Title)
        }
    }
    Else {
        Copy-View -SourceSiteConnection $SrcSiteConn -DestinationSiteConnection $DestSiteConn -SourceListName $SourceListName -DestinationListName $DestinationListName -ViewName $ViewName
    }

}
catch { Throw }

#check aggregation