<# ToDo and notes
    * Check variable $SupportedFlowsAndRemediations to aknowledge which are the supported flows, their Preventive Check Action and Remediation Actions that will be performed.

    ToDo:
        ! Add -WhatIf to script/functions
        ! Add to CSV export: PATicketResolutionNotes
#>

#Region Requirements, Parameters, Variables
#Requires -Version 7 -Modules @{ ModuleName = "Microsoft.PowerApps.PowerShell"; ModuleVersion = "1.0.32" }, @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2.0" }

# Parameters
Param (
    # Path of the CSV exported from ServiceNow with the list of Power Automate tickets to be processed
    [Parameter(
        Mandatory = $true,
        HelpMessage = 'Path of the CSV exported from ServiceNow with the list of tickets with flows to resubmit.'
    )]
    [ValidateScript({ Test-Path -Path $($_ -replace '"', '') -PathType Leaf })]
    [String]
    $PATicketCSVPath,

    # Delimiter of the CSV
    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Delimiter used in the CSV extracted from ServiceNow.'
    )]
    [ValidateSet(',', ';')]
    [String]
    $CSVDelimiter = ',',

    # Path of the PowerShell script that contains the list of supported flows, their Preventive Check Action and Remediation Actions ScriptBlocks
    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Path of the PowerShell script that contains the list of supported flows and their remediation actions as ScriptBlocks.'
    )]
    [AllowNull()]
    [AllowEmptyString()]
    [String]
    $SupportedFlowsAndRemediationsScriptPath
)

# Variables and logging setup
Try
{
    <# URI of the flow used for resubmissions
        Link to the flow: https://make.powerautomate.com/environments/Default-7cc91888-5aa0-49e5-a836-22cda2eae0fc/flows/11dcfd38-a20e-4e6e-aa02-e164bdbecc9f/details
    #>
    $AMSResubmitFlowUri = 'https://prod-209.westeurope.logic.azure.com:443/workflows/8a0ba5d97be94e65a619e92acb032496/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jWKnOGirP7PL8RVu6HSW_yau0yJo5lZ3FKpbUFWs2a4'

    # Compose logs path and create log folder if it doesn't exist
    $ScriptName = (Get-Item -Path $MyInvocation.MyCommand.Path).BaseName
    $ExecutionDateTime = Get-Date -Format 'dd_MM_yyyy-HH_mm_ss'
    $LogRootFolder = "$($PSScriptRoot)\Logs"
    $LogFolder = "$($LogRootFolder)\$($ScriptName)_$($ExecutionDateTime)"
    $CSVLogPath = "$($LogFolder)\$($ScriptName)_CSV_$($ExecutionDateTime).csv"
    $PnPTraceLogPath = "$($LogFolder)\$($ScriptName)_PnPTraceLog_$($ExecutionDateTime).log"
    $TranscriptLogPath = "$($LogFolder)\$($ScriptName)_ConsoleLog_$($ExecutionDateTime).log"
    If (!(Test-Path -Path $LogFolder -PathType Container))
    {
        New-Item $LogFolder -Force -ItemType Directory | Out-Null
    }
    Start-Transcript -Path $TranscriptLogPath -IncludeInvocationHeader | Out-Null

    # Start $ScriptStopwatch to measure execution time
    $ScriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $ScriptExecutionStartDate = (Get-Date -Format 'dd/MM/yyyy - HH:mm:ss')
    Write-Host ('{0}Script execution start date and time: {1}' -f "`n", $ScriptExecutionStartDate) -ForegroundColor Green

    # Copy the CSV exported from ServiceNow inside the logs folder
    $PATicketCSVPath = $PATicketCSVPath -replace '["'']', ''
    $PATicketCSVSourceFileName = Split-Path -Path $PATicketCSVPath -Leaf
    $PATicketCSVSourceFilePath = (Copy-Item -Path $PATicketCSVPath -Destination "$($LogFolder)\SourceFile_$($PATicketCSVSourceFileName)" -Force -PassThru -ErrorAction Stop).FullName

    <# Import $SupportedFlowsAndRemediations script.
        If no value is specified for the parameter, the default script will be used from the same folder of this script
    #>
    If (-not $SupportedFlowsAndRemediationsScriptPath)
    {
        # Using default script
        $SupportedFlowsAndRemediationsScriptPath = "$($PSScriptRoot)\SupportedFlowsAndRemediations.ps1"
        Write-Host ("{0}No value specified for parameter 'SupportedFlowsAndRemediationsScriptPath'. Default script will be used from following path:{0}{1}" -f
            "`n",
            $SupportedFlowsAndRemediationsScriptPath
        ) -ForegroundColor Yellow

        # If the default script is not found, the execution will be stopped
        If (!(Test-Path -Path $SupportedFlowsAndRemediationsScriptPath -PathType Leaf))
        {
            Throw ("{0}Default remediation script not found in '{1}'" -f
                "`n",
                (Split-Path -Path $SupportedFlowsAndRemediationsScriptPath -Parent)
            )
        }
    }
    Else
    {
        # Using the script provided by the user in the parameter
        $SupportedFlowsAndRemediationsScriptPath = $SupportedFlowsAndRemediationsScriptPath -replace '["'']', ''
        If (!(Test-Path -Path $SupportedFlowsAndRemediationsScriptPath -PathType Leaf))
        {
            Throw ("{0}Provided remediation script '{1}' not found in '{2}'" -f
                "`n",
                (Split-Path -Path $SupportedFlowsAndRemediationsScriptPath -Leaf),
                (Split-Path -Path $SupportedFlowsAndRemediationsScriptPath -Parent)
            )
        }
    }
    # Import the script
    $SupportedFlowsAndRemediations = . $SupportedFlowsAndRemediationsScriptPath

    # Compose temporary folder path
    $Global:TmpFolderPath = ('{0}\Temp\#{1}' -f
        $($ENV:LOCALAPPDATA),
        $ScriptName
    )

    # Delete temporary folder if it already exists
    If (Test-Path -Path $Global:TmpFolderPath -PathType Container)
    {
        Remove-Item -Path $Global:TmpFolderPath -Force -Recurse
    }

    # Initialize/reset variables
    $PATicketList = @()
    $ScriptError = $null
    $TotalSkippedTicket = 0
    $TotalProcessedTicket = 0
    $Global:TmpCSVLists = @()
    $ScriptElapsedTime = $null
    #$Global:SPOConnections = @()
    $Global:ProgressBarsIds = @{}
    $ExecutionsLogsDetails = $null
    $Global:ProgressBarsIds[0] = $true
}
Catch
{
    $ScriptStopwatch.Stop()
    Write-Host ('{0}Error:{1}{0}' -f
        "`n",
        ($_ | Out-String).TrimEnd()
    ) -ForegroundColor Red
    Stop-Transcript | Out-Null
    Exit 1
}
#EndRegion Requirements, Parameters, Variables

#Region Functions

# Function to convert a PSCustomObject and its nested PSCustomObject properties to an array of strings
Function Convert-PSCustomObjectToList
{
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $InputPSCustomObject
    )

    # Get properties of the input PSCustomObject
    $Properties = $InputPSCustomObject.PSObject.Properties | Select-Object Name, Value | Select-Object -ExpandProperty Name

    # Loop through properties and expand nested PSCustomObject properties when needed
    ForEach ($Property In $Properties)
    {
        $Value = $InputPSCustomObject.$Property

        If ($Value -is [PSCustomObject])
        {
            Write-Output ($Value | Format-List | Out-String).Trim()
        }
        Else
        {
            Write-Output ('{0}: {1}' -f
                $Property,
                $Value
            )
        }
    }
}

# Function to validate the $SupportedFlowsAndRemediations object
Function Confirm-SupportedFlowsAndRemediations
{
    Param(
        # Dot sourced Object that contains the list of supported flows and their resolution criteria
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Object]
        $SupportedFlowsAndRemediations
    )

    Try
    {
        # Hashtable to track uniqueness of DisplayName, Remediation.Name, and Remediation.ErrorToRemediate
        $UniqueDisplayName = @{}
        $UniqueRemediationName = @{}
        $UniqueRemediationError = @{}
        $UniqueRemediationExecutionOrder = @{}

        # Loop through each flow and remediation action to validate the object
        ForEach ($Flow in $SupportedFlowsAndRemediations)
        {
            # Check DisplayName property
            If (-not $Flow.DisplayName)
            {
                Throw "'DisplayName' is required for each supported flow."
            }

            # Check uniqueness of DisplayName
            If ($UniqueDisplayName.ContainsKey($Flow.DisplayName))
            {
                Throw ("Duplicate flow found by 'DisplayName':{0}'{1}'." -f
                    "`n",
                    $Flow.DisplayName
                )
            }
            Else
            {
                $UniqueDisplayName[$Flow.DisplayName] = $true
            }

            # Check ResubmitIfEmptyFlowErrors property
            If (
                $true -ne $Flow.ResubmitIfEmptyFlowErrors -and
                $false -ne $Flow.ResubmitIfEmptyFlowErrors
            )
            {
                Throw ("Missing required property 'ResubmitIfEmptyFlowErrors' for Flow '{0}'." -f $($Flow.DisplayName))
            }

            # Check ResubmitIfUnsupportedFlowErrors property
            If (
                $true -ne $Flow.ResubmitIfUnsupportedFlowErrors -and
                $false -ne $Flow.ResubmitIfUnsupportedFlowErrors
            )
            {
                Throw ("Missing required property 'ResubmitIfUnsupportedFlowErrors' for Flow '{0}'." -f $($Flow.DisplayName))
            }

            # Check PreventiveCheckAction property type
            If (
                !(-not $Flow.PreventiveCheckAction) -and
                !($Flow.PreventiveCheckAction -is [ScriptBlock])
            )
            {
                Throw ("Property 'PreventiveCheckAction' for Flow '{0}' is not a ScriptBlock." -f $($Flow.DisplayName))
            }

            # Check Remediations property
            If ($Flow.Remediations)
            {
                ForEach ($Remediation in $Flow.Remediations)
                {
                    # Check Name property
                    If (-not $Remediation.Name)
                    {
                        Throw ("Missing required Remediation property 'Name' for flow '{0}'." -f $($Flow.DisplayName))
                    }

                    # Check uniqueness of Name property within the same flow remediations
                    $RemediationKey = ('{0}|{1}' -f
                        $Flow.DisplayName,
                        $Remediation.Name
                    )
                    If ($UniqueRemediationName.ContainsKey($RemediationKey))
                    {
                        Throw ("Duplicate property 'Remediation Name' found for Flow '{0}':{1}'{2}'" -f
                            $Flow.DisplayName,
                            "`n",
                            $Remediation.Name
                        )
                    }
                    Else
                    {
                        $UniqueRemediationName[$RemediationKey] = $true
                    }

                    # Check ErrorToRemediate property
                    If (-not $Remediation.ErrorToRemediate)
                    {
                        Throw ("Missing required Remediation property 'ErrorToRemediate' for Remediation '{0}' of Flow '{1}'." -f
                            $($Remediation.Name),
                            $($Flow.DisplayName)
                        )
                    }

                    # Check uniqueness of Remediation.ErrorToRemediate within the same flow
                    $ErrorKey = '{0}|{1}' -f $Flow.DisplayName, $Remediation.ErrorToRemediate
                    If ($UniqueRemediationError.ContainsKey($ErrorKey))
                    {
                        Throw ("Duplicate property 'ErrorToRemediate' found for Remediation '{0}' of Flow '{1}':{2}'{3}'" -f
                            $($Remediation.Name),
                            $Flow.DisplayName,
                            "`n",
                            $Remediation.ErrorToRemediate
                        )
                    }
                    Else
                    {
                        $UniqueRemediationError[$ErrorKey] = $true
                    }

                    # Check ExecutionOrder property
                    If ($Remediation.ExecutionOrder -isnot [Int])
                    {
                        Throw ("Missing required Remediation property 'ExecutionOrder' for Remediation '{0}' of Flow '{1}'." -f
                            $($Remediation.Name),
                            $($Flow.DisplayName)
                        )
                    }

                    # Check uniqueness of Remediation.ExecutionOrder within the same Flow
                    $ExecutionOrderKey = '{0}|{1}' -f $Flow.DisplayName, $Remediation.ExecutionOrder
                    If ($UniqueRemediationExecutionOrder.ContainsKey($ExecutionOrderKey))
                    {
                        Throw ("Duplicate property 'ExecutionOrder' found for Remediation '{0}' of Flow '{1}':{2}'{3}'" -f
                            $($Remediation.Name),
                            $Flow.DisplayName,
                            "`n",
                            $Remediation.ErrorToRemediate
                        )
                    }
                    Else
                    {
                        $UniqueRemediationExecutionOrder[$ExecutionOrderKey] = $true
                    }

                    # Check IsSPOConnectionRequired property
                    If ($null -eq $Remediation.IsSPOConnectionRequired)
                    {
                        Throw ("Missing required Remediation property 'IsSPOConnectionRequired' for Remediation '{0}' of flow '{1}'." -f
                            $($Remediation.Name),
                            $($Flow.DisplayName)
                        )
                    }

                    # Check Action property
                    If (-not $Remediation.Action)
                    {
                        Throw ("Missing required Remediation property 'Action' for Remediation '{0}' of Flow '{1}'." -f
                            $($Remediation.Name),
                            $($Flow.DisplayName)
                        )
                    }

                    # Check Action property type
                    If (-not ($Remediation.Action -is [ScriptBlock]))
                    {
                        Throw ("Property 'Action' for Flow '{0}' is not a ScriptBlock." -f $($Flow.DisplayName))
                    }
                }
            }
        }

        # If all validations pass, return $null
        Return $null
    }
    Catch { Throw }
}

# Function that returns the list of tickets to resubmit the flow for, got from the CSV exported from ServiceNow
Function Get-PATicketList
{
    #Requires -Module Microsoft.PowerApps.PowerShell

    Param(
        # Path of the CSV exported from ServiceNow with the list of tickets with flow to resubmit
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $PATicketCSVPath,

        # Dot sourced Object that contains the list of supported flows and their resolution criteria
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Object]
        $SupportedFlowsAndRemediations
    )

    Try
    {
        # Check if the module 'Microsoft.PowerApps.PowerShell' is imported
        If (-not (Get-Module -Name Microsoft.PowerApps.PowerShell))
        {
            Throw "This function requires the module 'Microsoft.PowerApps.PowerShell' to be imported."
        }


        # Import the CSV with the list of tickets to resubmit the flow for
        $PATicketCSV = Import-Csv -Path $PATicketCSVPath -Delimiter $CSVDelimiter

        # Ensure that the CSV contains at least one ticket
        If ($PATicketCSV.Count -eq 0)
        {
            Throw 'The CSV must contain at least one ticket.'
        }

        # Validate columns of the CSV
        If ($PATicketCSV[0].PSObject.Properties.Name -notcontains 'number' -or $PATicketCSV[0].PSObject.Properties.Name -notcontains 'Description')
        {
            Throw "The CSV must contain the following columns: 'number', 'Description'."
        }

        # Create object with needed properties for each ticket
        $Global:PA_Environment = (Get-PowerAppEnvironment | Where-Object -FilterScript { $_.IsDefault }).EnvironmentName
        [Array]$PATicketList = $PATicketCSV | ForEach-Object {
            $PATicketHyperTextLink = New-PowerShellHyperlink -LinkURL $('https://tecnimont.service-now.com/incident.do?sysparm_query=number={0}' -f $_.number) -LinkDisplayText $_.number
            $EnvironmentID = ($((($_.Description -Split 'Environment: ')[1] -Split "`n")[0] | Where-Object -FilterScript { $_ -ne '' }) ?? $Global:PA_Environment)
            $RunID = $((($_.Description -Split "`nRun ID: ")[1] -Split "`n")[0] | Where-Object -FilterScript { $_ -ne '' })
            $FlowID = $((($_.Description -Split "`nFlow ID: ")[1] -Split "`n")[0] | Where-Object -FilterScript { $_ -ne '' })
            If ($RunID -and $FlowID)
            {
                $FlowRunLink = ('https://make.powerautomate.com/environments/{0}/flows/{1}/runs/{2}' -f $EnvironmentID, $FlowID, $RunID)
            }
            else
            {
                Throw "Run ID or Flow ID not found in description of ticket $($_.number)."
            }

            [PSCustomObject]@{
                PATicketID                    = $_.number
                PATicketHyperTextLink         = $PATicketHyperTextLink
                FlowID                        = $FlowID
                EnvironmentID                 = $EnvironmentID
                FlowDisplayName               = $null
                SupportedFlowDisplayName      = $null
                FlowHyperTextLink             = $null
                IsSupportedFlow               = $null
                RunID                         = $RunID
                FlowRunLink                   = $FlowRunLink ?? 'N/A'
                SiteUrl                       = $((($_.Description -Split "`nSite URL: ")[1] -Split "`n")[0] | Where-Object -FilterScript { $_ -ne '' })
                AMSIdentifier                 = $null
                FlowErrorDetails              = $(($_.Description -Split "`nFlow Error Details:")[1] -Split "`n" | Where-Object -FilterScript { $_ -ne '' })
                PATicketDescription           = $($_.Description)
                TriggerName                   = $null
                PATicketResolutionTime        = $null
                PATicketAverageResolutionTime = $null
            }
        }

        # Get the display name of each flow and check if it is supported
        $Global:FlowList = $PATicketList | Select-Object -Property FlowID, EnvironmentID -Unique
        Write-Host ('{0}______________________________{0}{0}Starting Flow search by IDs...{0}______________________________' -f "`n") -ForegroundColor DarkMagenta
        $Counter = 0
        ForEach ($Flow in $Global:FlowList)
        {
            $Counter++
            $PercentComplete = ($Counter / $($Global:FlowList.Count) * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Flow search')
                Status          = ("Searching Flow '{0}'..." -f $Flow.FlowID)
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = 0
            }
            Start-Sleep -Milliseconds 250
            Write-Progress @ProgressBarSplatting

            # Get flow DisplayName
            $FlowDisplayName = $null
            $FlowURL = ('https://make.powerautomate.com/environments/{0}/flows/{1}/details' -f
                $Flow.EnvironmentID,
                $Flow.FlowID
            )
            $FlowLink = New-PowerShellHyperlink -LinkURL $FlowURL -LinkDisplayText $FlowURL
            $FlowHyperTextLink = New-PowerShellHyperlink -LinkURL $FlowURL -LinkDisplayText $($FlowDisplayName ?? $Flow.FlowID)
            Write-Host ("{0}Searching Flow ID: '{1}'..." -f "`n", $FlowHyperTextLink) -ForegroundColor Cyan

            # Save current ProgressPreference
            $OldProgressPreference = $ProgressPreference
            # Temporarily change ProgressPreference to suppress progress bar from Get-Flow cmdlet
            $ProgressPreference = 'SilentlyContinue'
            $FlowObject = Get-Flow -FlowName $Flow.FlowID -EnvironmentName $Flow.EnvironmentID
            # Restore original ProgressPreference
            $ProgressPreference = $OldProgressPreference

            $FlowDisplayName = $FlowObject.DisplayName
            $FlowHyperTextLink = New-PowerShellHyperlink -LinkURL $FlowURL -LinkDisplayText $($FlowDisplayName ?? $Flow.FlowID)
            $Flow | Add-Member -MemberType NoteProperty -Name FlowURL -Value $FlowURL
            $Flow | Add-Member -MemberType NoteProperty -Name FlowHyperTextLink -Value $FlowHyperTextLink
            If (!(-not $FlowDisplayName))
            {
                $FlowSearchResult = ("Found Flow '{1}'.{0}Flow URL: {2}" -f
                    "`n",
                    $FlowHyperTextLink,
                    $FlowLink
                )
                $ForegroundColor = 'DarkGreen'
            }
            ElseIf (!(-not $FlowObject.Internal))
            {
                $FlowSearchResult = ("Error getting Flow '{1}'.{0}Flow URL: {2}{0}Error: {3} ({4})" -f
                    "`n",
                    $FlowHyperTextLink,
                    $FlowLink,
                    $FlowObject.Internal.StatusCode,
                    $FlowObject.Internal.Error.code
                )
                $ForegroundColor = 'DarkRed'
                $FlowSearchError = $true
            }
            Else
            {
                $FlowSearchResult = ("Flow '{1}' not found (no response).{0}Flow URL: {2}" -f
                    "`n",
                    $FlowHyperTextLink,
                    $FlowLink
                )
                $ForegroundColor = 'Yellow'
                $FlowSearchError = $true
            }
            $ConsoleOutputForegroundColor = @{
                ForegroundColor = $ForegroundColor
            }
            Write-Host $FlowSearchResult @ConsoleOutputForegroundColor

            # Check if the flow is found
            If (-not $FlowDisplayName -or -not $Flow.FlowID)
            {
                $FlowDisplayName = 'Not found'
                $FlowTriggerName = 'N/A'
                $AMSIdentifierName = 'N/A'
                $IsSupportedFlow = $false
            }
            Else
            {
                # Check if the flow is supported
                ForEach ($SupportedFlow in $SupportedFlowsAndRemediations)
                {
                    If ($FlowDisplayName -like "*$($SupportedFlow.DisplayName)*")
                    {
                        $IsSupportedFlow = $true
                        $AMSIdentifierName = $SupportedFlow.AMSIdentifier ?? 'N/A'
                        $SupportedFlowDisplayName = $SupportedFlow.DisplayName
                        Break
                    }
                    Else
                    {
                        $IsSupportedFlow = $false
                        $AMSIdentifierName = 'N/A'
                    }
                }
            }

            # Set the trigger name of the flow
            If ($IsSupportedFlow -eq $false)
            {
                $FlowTriggerName = 'N/A'
            }
            Else
            {
                $FlowTriggerName = ($FlowObject.Internal.properties.definition.triggers | Get-Member -MemberType NoteProperty).Name
            }

            # Add the flow display name to the ticket object
            $PATicketList | Where-Object { $_.FlowID -eq $Flow.FlowID } | ForEach-Object {
                $_.FlowDisplayName = $FlowDisplayName
                $_.SupportedFlowDisplayName = $SupportedFlowDisplayName ?? 'N/A'
                $_.FlowHyperTextLink = $FlowHyperTextLink
                $_.TriggerName = $FlowTriggerName
                $_.IsSupportedFlow = $IsSupportedFlow
                $_.AMSIdentifier = [PSCustomObject]@{
                    AMSIdentifierName  = $AMSIdentifierName
                    AMSIdentifierValue = (
                        $((($_.PATicketDescription -Split ('{0}: ' -f $AMSIdentifierName))[1] -Split "`n")[0] | Where-Object -FilterScript { $_ -ne '' }) ??
                        'N/A'
                    )
                }
            }
        }
        Write-Host ('{0}______________________{0}{0}Flow search completed.{0}______________________' -f "`n") -ForegroundColor DarkMagenta
        Write-Progress -Activity 'Flow Search' -Completed -Id 0
        if ($FlowSearchError)
        {
            Write-Host ''
            Throw 'Error searching for one or more flows.'
        }
        Return $PATicketList
    }
    Catch
    {
        Write-Progress -Activity 'Flow Search' -Completed -Id 0
        Throw
    }
}

# Function to get the details of the error of a flow from the PATicket object
Function Get-PATicketFlowErrorDetails
{
    Param(
        # Display name of the flow to get the error details for
        [AllowNull()]
        [String]
        $PATicketFlowDisplayName,

        # Array of Flow Error Details passed from the PATicket object rertieved from Get-PATicketList function
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [Array]
        $PATicketFlowError
    )

    Try
    {
        # Create object with needed properties
        $PATicketFlowErrorObject = [PSCustomObject]@{
            FlowDisplayName  = $PATicketFlowDisplayName
            FlowErrorDetails = [Array]@()
        }

        # If $PATicketFlowError contains Flow Error Details, create a custom object to be added to $PATicketFlowErrorObject.FlowErrorDetails
        If ($PATicketFlowError)
        {
            # Parse the Flow Error Details by checking blocks made of the 3 required properties (Error Action, Error Code and Error Message)
            $ErrorLines = $PATicketFlowError | Where-Object -FilterScript { $null -ne $_ -and '' -ne $_ -and "`r" -ne $_ }
            ForEach ($Index in 0..($ErrorLines.Count - 1) | Where-Object { $_ % 3 -eq 0 })
            {
                $ErrorAction = $ErrorLines[$Index] -replace 'Error on action: '
                $ErrorCode = $ErrorLines[$Index + 1] -replace 'Error code: '
                $ErrorMessages = $ErrorLines[$Index + 2] -replace 'Error message: '

                $ErrorObject = [PSCustomObject]@{
                    ErrorAction   = $ErrorAction.Trim()
                    ErrorCode     = $ErrorCode.Trim()
                    ErrorMessages = $ErrorMessages.Trim()
                }

                $PATicketFlowErrorObject.FlowErrorDetails += $ErrorObject
            }
        }
        Else
        {
            # If PATicket doesn't contain Flow Error Details, create a custom object with null values
            $ErrorObject = [PSCustomObject]@{
                ErrorAction   = $null
                ErrorCode     = $null
                ErrorMessages = $null
            }
            $PATicketFlowErrorObject.FlowErrorDetails = $ErrorObject
        }

        Return $PATicketFlowErrorObject
    }
    Catch { Throw }
}

# Function to connect to SharePoint Online Site if a Connection object for the specified SiteUrl is not already present in the global variable $Global:SPOConnections
Function Connect-SPOSite
{
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
                If ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$')
                {
                    $True
                }
                Else
                {
                    Throw "`n'$($_)' is not a valid SharePoint Online site or subsite URL."
                }
            })]
        [String]
        $SiteUrl
    )

    Try
    {

        # Initialize Global:SPOConnections array if not already initialized
        If (-not $Global:SPOConnections)
        {
            $Global:SPOConnections = @()
        }
        Else
        {
            # Check if SPOConnection to specified Site already exists
            $SPOConnection = ($Global:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $SiteUrl }).Connection
        }

        # Create SPOConnection to specified Site if not already established
        If (-not $SPOConnection)
        {
            # Create SPOConnection to SiteURL
            $SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

            # Add SPOConnection to the list of connections
            $Global:SPOConnections += [PSCustomObject]@{
                SiteUrl    = $SiteUrl
                Connection = $SPOConnection
            }
        }

        Return $SPOConnection
    }
    Catch
    {
        Throw
    }
}

# Function to search for a setting key in available settings lists of a site
Function Search-SettingKey
{
    Param(
        # Name of the settings list
        [Parameter(Mandatory = $true)]
        [ValidateSet(
            'Configuration List',
            'Settings'
        )]
        [String]
        $SettingsListName,

        # Name of the setting to search for
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $SettingName,

        # Sharepoint Online Connection Object
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection
    )

    Try
    {

        # Get the site type from SiteURL
        $SiteURL = $SPOConnection.Url.ToLower().TrimEnd('/')
        If ($SiteURL.Contains('vdm_'))
        {
            $SiteType = 'VD'

        }
        ElseIf ($SiteURL.Contains('DigitalDocuments'))
        {
            $SiteType = 'DD'
        }

        # Get the name of the setting value column based on site type
        Switch ($SiteType)
        {
            'VD'
            {
                $SettingColumnName = 'VD_ConfigValue'
                Break
            }

            'DD'
            {
                $SettingColumnName = 'Value'
                Break
            }

            Default
            { Throw 'Site type not found.' }
        }

        # Check if the settings list provided is correct for site type
        If ($SiteType -eq 'DD' -and $SettingsListName -eq 'Configuration List')
        {
            Throw ("Settings list '{0}' not available on site type '{1}'." -f
                $SettingsListName,
                $SiteType
            )
        }

        # Get all settings from the settings list
        $SettingsList = Get-PnPListItem -List $SettingsListName -Connection $SPOConnection -PageSize 5000 | ForEach-Object {
            [PSCustomObject]@{
                ID          = $_['ID']
                SettingName = $_['Title']
                Value       = $_["$($SettingColumnName)"]
            }
        }

        # Get the value of the setting
        $SettingValue = ($SettingsList | Where-Object -FilterScript { $_.SettingName -eq $SettingName }).Value

        # Return the setting value if found
        If ($null -eq $SettingValue)
        {
            Throw ("Setting '{0}' not found on List '{1}' of Site '{2}'." -f
                $SettingName,
                $SettingsListName,
                $SiteURL
            )
        }
        Return $SettingValue
    }
    Catch { Throw }
}

# Function to get remediation actions from $SupportedFlowsAndRemediations from the current Flow Error Details
Function Get-PATKFlowRemediationActions
{
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $PATicketFlowErrorDetails,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Object]
        $SupportedFlowsAndRemediations
    )

    Try
    {
        # Filter unique Flow Error Details
        $Unique_PATicketFlowErrorDetails = ($PATicketFlowErrorDetails.FlowErrorDetails | Select-Object -Unique -Property *)

        # Loop through all the Flow Error Details to get their remediations
        $Remediations = @()
        ForEach ($FlowError in $Unique_PATicketFlowErrorDetails)
        {
            # Get remediations for the current Flow sorted by ExecutionOrder
            $Remediation = ($SupportedFlowsAndRemediations |
                    Where-Object -FilterScript {
                        ForEach ($ErrorToRemediate in $_.Remediations.ErrorToRemediate)
                        {
                            $FlowError.ErrorMessages -like "*$($ErrorToRemediate)*" -and
                            $PATicketFlowErrorDetails.FlowDisplayName -like "*$($_.DisplayName)*"
                        }
                    } | Select-Object -ExpandProperty Remediations) |
                    Where-Object -FilterScript { $FlowError.ErrorMessages -like "*$($_.ErrorToRemediate)*" }

            # If no remediation has been found, add a dummy remediation to the list
            If (-not $Remediation)
            {
                # Set string value for the remediation action
                If (-not $FlowError.ErrorMessages)
                {
                    $StringAction = 'N/A'
                }
                Else
                {
                    $StringAction = 'Missing remediation'
                }

                # Dummy remediation action object
                $RemediationActionObject = [PSCustomObject][Ordered]@{}
                $RemediationActionObjectProperties = [PSCustomObject]@{
                    ErrorToRemediate = $($FlowError.ErrorMessages)
                    Action           = $StringAction
                }

                # Get all properties from the dummy remediation action object
                $RemediationActionObjectPropertiesList = @($RemediationActionObjectProperties | Get-Member -MemberType NoteProperty).Name

                # Get all properties from a standard remediation object in the Supported Flows List
                $StandardRemediationsProperties = @($SupportedFlowsAndRemediations.Remediations[0].PSObject.Properties).Name

                # Add all other properties from the remediation object
                $StandardRemediationsProperties | ForEach-Object {
                    If ( $_ -notin $RemediationActionObjectPropertiesList )
                    {
                        # Add unsupported value to the property
                        $RemediationActionObject | Add-Member -NotePropertyName $_ -NotePropertyValue 'N/A'
                    }
                    Else
                    {
                        # Add value to the property from the temporary remediation action object
                        $RemediationActionObject | Add-Member -NotePropertyName $_ -NotePropertyValue $RemediationActionObjectProperties.$($_)
                    }
                }

                # Add remediation to the list
                $Remediations += $RemediationActionObject
            }
            Else
            {
                # Create the ScriptBlock and add the remediation to the Remediations list
                # Create ScriptBlock from the remediation action and replace the string value in Remediation object with it
                $Remediation.Action = [ScriptBlock]::Create($($($Remediation.Action).ToString() -replace '(?m)^\s+').Trim(@(' ', "`n", "`r")))

                # Add remediation to the list
                $Remediations += $Remediation
            }
        }

        # Select unique error message to handle every error only once with its specific remediation action
        $Remediations = $Remediations | Select-Object -Unique -Property * | Sort-Object -Property ExecutionOrder

        Return $Remediations
    }
    Catch
    { Throw }
}

# Function to get all items from a SharePoint Online list with custom properties
Function Get-AMSSPOList
{
    [CmdletBinding(DefaultParameterSetName = 'CSVExpirationSet')]
    Param (
        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CSVExpirationOverrideSet', Mandatory = $true)]
        [ValidateSet(
            'Vendor Documents List',
            'Process Flow Status List',
            'Revision Folder Dashboard'
        )]
        [String]
        $ListName,

        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CSVExpirationOverrideSet', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection,

        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -PathType Container -IsValid })]
        [String]
        $TmpFolderPath,

        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $false)]
        [ValidateRange(1, 60)]
        [ArgumentCompleter({
                Param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                New-Object -Type System.Management.Automation.CompletionResult -ArgumentList 20, 20, 'ParameterValue', 'Set a value from 1 to 60. If not specified, default value is 15.'
            })]
        [Int]
        $CSVListExpirationMinutes = 15, # If this parameter default value is changed, the default value in the information message in the ArgumentCompleter will need to be changed as well

        # Parameter to override the CSVListExpirationMinutes parameter and force the function to load the list from SharePoint Online instead of previously saved temporary CSV file
        [Parameter(ParameterSetName = 'CSVExpirationOverrideSet', Mandatory = $true)]
        [Switch]
        $ForceOnlineList
    )

    Try
    {
        Write-Host ("Items from List '{0}' of Site '{1}' are required to proceed..." -f $ListName, $SPOConnection.Url) -ForegroundColor DarkMagenta

        # Checking for provided CSVListExpirationMinutes or ForceOnlineList parameter
        If ($ForceOnlineList -ne $true)
        {
            # If exists, use a previously saved temporary CSV file to load the list items
            # Warn that fucntion is checking for temporary CSV files within provided or default time range in $CSVListExpirationMinutes parameter
            If ($PSBoundParameters.ContainsKey('CSVListExpirationMinutes'))
            {
                $ParamInformationMessages = ('CSVListExpirationMinutes parameter was provided by the user with value: {0}' -f $CSVListExpirationMinutes)
                Write-Host $ParamInformationMessages -ForegroundColor Magenta
            }
            Else
            {
                $ParamInformationMessages = ('The CSVListExpirationMinutes parameter was not provided by the user, using the default value: {0}.' -f $CSVListExpirationMinutes)
                Write-Host ('The CSVListExpirationMinutes parameter was not provided by the user, using the default value: {0}.' -f $CSVListExpirationMinutes) -ForegroundColor Yellow
            }
            Write-Host ('Searching valid temporary CSV file containing required List Items...') -ForegroundColor Cyan

            <# Check if:
               - a the temporary CSV file exists for given List on provided SiteUrl in $Global:TmpCSVLists array within provided time range
               - a the temporary CSV file exists for given List on provided SiteUrl inside the temporary folder with correct naming convention
               - the file is not empty
               - the file is not older than the minutes set with $CSVListExpirationMinutes parameter
            #>
            $Filter_TmpCSVFilePath = ('{0}\{1}_{2}_{3}.csv' -f
                $Global:TmpFolderPath,
                $SPOConnection.Url.Split('/')[-1].ToUpper(),
                ($ListName -Split ' ' -Join '_'),
                '*'
            )
            $TmpCSVList = $Global:TmpCSVLists | Where-Object -FilterScript {
                $_.ListName -eq $ListName -and
                $_.SiteUrl -eq $SPOConnection.Url -and
                $_.ListItemsCSVFilePath -like $Filter_TmpCSVFilePath -and
                ((Get-Date) - $_.CSVListCreationDate).TotalMinutes -le $_.CSVListExpirationMinutes -and
                (Test-Path -Path $_.ListItemsCSVFilePath -PathType Leaf) -and
                ((Get-Content -Path $_.ListItemsCSVFilePath -ErrorAction SilentlyContinue).Count -gt 2) -and
                (((Get-Date) - (Get-Item -Path $_.ListItemsCSVFilePath -ErrorAction SilentlyContinue).CreationTime).TotalMinutes) -le $_.CSVListExpirationMinutes
            }

            # Return error if more than 1 temporary CSV file is found for the same List, Site and DateTime
            If ($TmpCSVList.Count -gt 1)
            {
                Throw ("[ERROR] More than 1 temporary CSV file found for List '{0}' and Site '{1}':{2}{3}" -f
                    $ListName,
                    $SPOConnection.Url,
                    "`n",
                    $TmpCSVList
                )
            }
            # Since file exists, import list items from the temporary CSV file
            ElseIf ($null -ne $TmpCSVList)
            {
                $StartImportInformationMessage = ('Importing list items from found temporary CSV file:{0}{1}' -f
                    "`n",
                    $TmpCSVList.ListItemsCSVFilePath
                )
                Write-Host $StartImportInformationMessage -ForegroundColor Cyan
                [Array]$ListItems = Import-Csv -Path $TmpCSVList.ListItemsCSVFilePath -Delimiter ';' -Encoding UTF8BOM
                $EndImportInformationMessage = ('List items imported from found temporary CSV file' -f $TmpCSVList.ListItemsCSVFilePath)
                Write-Host $EndImportInformationMessage -ForegroundColor DarkGreen

                # Add information messages to the global CSVOutput object and return the list items
                #$Global:CSVOutput.AdditionalDetails += ($ParamInformationMessages, $StartImportInformationMessage, $EndImportInformationMessage) -join "`n" + "`n"
                Return $ListItems
            }
            Else
            {
                $CSV_SearchResultMessage = ('No valid CSV file found for this purpose.{2}List will now be downloaded from SharePoint Online to a temporary CSV file for future uses within {3} minutes during current execution.' -f
                    $ListName,
                    $SPOConnection.Url,
                    "`n",
                    $CSVListExpirationMinutes
                )
                Write-Host $CSV_SearchResultMessage -ForegroundColor Cyan
            }
        }
        Else
        {
            # Warn that fucntion will not check for temporary CSV files and will load the list from SharePoint Online
            $ParamInformationMessages = ('ForceOnlineList parameter was provided by the user, ignoring default CSVListExpirationMinutes parameter ({0}).{1}List will be loaded from SharePoint Online and will not be saved to temporary CSV file for future uses during current execution.' -f
                $CSVListExpirationMinutes,
                "`n"
            )
            Write-Host $ParamInformationMessages -ForegroundColor DarkMagenta
        }

        # Since ForceOnlineList parameter was provided or CSV file was not found, load the list from SharePoint Online

        # Get Fields from List
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')] # Suppress IDE's linter warning for next line ($ListFields variable)
        [Array]$ListFields = Get-PnPField -List $ListName -Connection $SPOConnection | Select-Object -Property InternalName, Title, TypeAsString

        # Compose PSCustomObject based on ListName
        Switch ($ListName)
        {
            'Vendor Documents List'
            {
                [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')] # Suppress IDE's linter warning for next line ($VendorSiteUrl_ColumnInteralName variable)
                $VendorSiteUrl_ColumnInteralName = $(($ListFields | Where-Object -FilterScript {
                            $_.InternalName -eq 'VD_VendorName' -or
                            $_.InternalName -eq 'VendorName_x003a_Site_x0020_Url'
                        }).InternalName | Sort-Object -Descending -Top 1)

                $ListFilter = [ScriptBlock]::Create('
                    [PSCustomObject]@{
                        ID                 = $_["ID"]
                        TCM_DN             = $_["VD_DocumentNumber"]
                        Rev                = $_["VD_RevisionNumber"]
                        Index              = $_["VD_Index"]
                        DocTitle           = $_["VD_EnglishDocumentTitle"]
                        PONumber           = $_["VD_PONumber"]
                        MRCode             = $_["VD_MRCode"]
                        DisciplineOwnerTCM = $_["VD_DisciplineOwnerTCM"].LookupValue
                        DisciplinesTCM     = [Array]$_["VD_DisciplinesTCM"].LookupValue -join ''","''
                        VendorName         = $_["VD_VendorName"].LookupValue
                        Path               = $_["VD_DocumentPath"]
                        VendorSiteUrl      = $_[$VendorSiteUrl_ColumnInteralName].LookupValue
                    }
                ')
                Break
            }

            'Process Flow Status List'
            {
                $ListFilter = [ScriptBlock]::Create('
                    [PSCustomObject]@{
                        ID              = $_["ID"]
                        TCM_DN          = $_["VD_DocumentNumber"]
                        Index           = $_["VD_Index"]
                        Rev             = $_["VD_RevisionNumber"]
                        VDL_ID          = $_["VD_VDL_ID"]
                        Status          = $_["VD_DocumentStatus"]
                        CommentsEndDate = $_["VD_CommentsEndDate"]
                        SubSitePOURL    = $($_["VD_PONumberUrl"].Url)
                        VendorSiteUrl    = $($_["VD_PONumberUrl"].Url -replace ''/[^/]*$'', '''')
                    }
                ')
                Break
            }

            'Revision Folder Dashboard'
            {
                $ListFilter = [ScriptBlock]::Create('
                    [PSCustomObject]@{
                        ID     = $_["ID"]
                        TCM_DN = $_["VD_DocumentNumber"]
                        Rev    = $_["VD_RevisionNumber"]
                        Status = $_["VD_DocumentSubmissionStatus"]
                    }
                ')
                Break
            }

            Default
            { Throw ("List '{0}' not supported!" -f $ListName) }
        }

        # Get all list items
        $StartImportInformationMessage = ("Getting all items from list '{0}'..." -f $ListName)
        Write-Host $StartImportInformationMessage -ForegroundColor Cyan
        [Array]$ListItems = Get-PnPListItem -List $ListName -Connection $SPOConnection -PageSize 5000 | ForEach-Object {
            & $ListFilter
        }
        $EndImportInformationMessage = ("List '{0}' loaded." -f $ListName)
        Write-Host $EndImportInformationMessage -ForegroundColor DarkGreen

        # Save list items to temporary CSV file if ForceOnlineList parameter was not provided
        If ($ForceOnlineList -ne $true)
        {
            # Create temporary folder if it doesn't exist
            If (!(Test-Path -Path $Global:TmpFolderPath -PathType Container))
            {
                New-Item -Path $Global:TmpFolderPath -ItemType Directory | Out-Null
            }

            # Compose temporary CSV file path
            $CSVListCreationDate = (Get-Date)
            $TmpCSVFilePath = ('{0}\{1}_{2}_{3}.csv' -f
                $Global:TmpFolderPath,
                $SPOConnection.Url.Split('/')[-1].ToUpper(),
                ($ListName -Split ' ' -Join '_'),
                $CSVListCreationDate.ToString('dd_MM_yyyy-HH_mm_ss')
            )

            # Save list items to temporary CSV file
            $ListItems | Export-Csv -Path $TmpCSVFilePath -NoTypeInformation -Delimiter ';' -Encoding UTF8BOM -UseQuotes Always
            Write-Host ("Items from List '{0}' of Site '{1}' saved to:{2}{3}" -f
                $ListName,
                $SPOConnection.Url,
                "`n",
                $TmpCSVFilePath
            ) -ForegroundColor Cyan

            # Create item with properties ListItemsCSVFilePath, SiteUrl, ListName to be added to $Global:TmpCSVLists
            $Global:TmpCSVLists += [PSCustomObject]@{
                ListName                 = $ListName
                SiteUrl                  = $SPOConnection.Url
                ListItemsCSVFilePath     = $TmpCSVFilePath
                CSVListCreationDate      = $CSVListCreationDate
                CSVListExpirationMinutes = $CSVListExpirationMinutes ?? 'N/A'
            }
        }

        Return $ListItems
    }
    Catch { Throw }
}

<# Function to check for Document matching in Vendor Documents List and Process Flow Status List.
    Function will return object from both lists.
#>
Function Get-VDMDocument
{
    Param(
        # Document Number with Index as suffix
        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CSVExpirationOverrideSet', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
                If ($_ -eq 'N/A')
                {
                    Throw ("Invalid document number: '{0}'" -f $_)
                }
                Else { $True }
            })]
        [String]
        $FullDocumentNumber,

        # Sharepoint Online Connection Object
        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $true)]
        [Parameter(ParameterSetName = 'CSVExpirationOverrideSet', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection,

        [Parameter(ParameterSetName = 'CSVExpirationSet', Mandatory = $false)]
        [ValidateRange(1, 60)]
        [ArgumentCompleter({
                Param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                New-Object -Type System.Management.Automation.CompletionResult -ArgumentList 20, 20, 'ParameterValue', 'Set a value from 1 to 60. If not specified, default value is 15.'
            })]
        [Int]
        $CSVListExpirationMinutes = 15, # If this parameter default value is changed, the default value in the information message in the ArgumentCompleter will need to be changed as well

        # Parameter to force function 'Get-AMSSPOList' to load the list from SharePoint Online instead of previously saved temporary CSV file
        [Parameter(ParameterSetName = 'CSVExpirationOverrideSet', Mandatory = $true)]
        [Switch]
        $ForceOnlineList
    )

    Try
    {
        # Create splatting for Get-AMSSPOList function based on $ForceOnlineList parameter
        If ($ForceOnlineList)
        {
            $ListLoadBehaviour = @{
                ForceOnlineList = $true
            }
        }
        Else
        {
            $ListLoadBehaviour = @{
                TmpFolderPath            = $Global:TmpFolderPath
                CSVListExpirationMinutes = $CSVListExpirationMinutes
            }
        }

        # Set variables for Document
        $DocumentNumber = $FullDocumentNumber.Substring(0, $FullDocumentNumber.Length - 4)
        [Int]$DocumentIndex = $FullDocumentNumber.Substring($FullDocumentNumber.Length - 3)

        # Get all Vendor Documents List items
        $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
        $CurrentProgressBarId = $ParentProgressBarId + 1
        $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
        $PercentComplete = (8 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document Search')
            Status          = ("Loading 'Vendor Documents List'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        [Array]$VDL_Items = Get-AMSSPOList -ListName 'Vendor Documents List' -SPOConnection $SPOConnection @ListLoadBehaviour

        # Get all Process Flow Status List items
        $PercentComplete = (37 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document Search')
            Status          = ("Loading 'Process Flow Status List'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        [Array]$PFSL_Items = Get-AMSSPOList -ListName 'Process Flow Status List' -SPOConnection $SPOConnection @ListLoadBehaviour

        # Filter Document from Vendor Documents List
        $PercentComplete = (52 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document Search')
            Status          = ("Filtering 'Vendor Documents List'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        [Array]$VDL_Document = $VDL_Items | Where-Object -FilterScript { $_.TCM_DN -eq $DocumentNumber -and $_.Index -eq $DocumentIndex }

        # Filter Document from Process Flow Status List
        $PercentComplete = (72 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document Search')
            Status          = ("Filtering 'Process Flow Status List'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        [Array]$PFSL_Document = $PFSL_Items | Where-Object -FilterScript {
            ($_.TCM_DN -eq $DocumentNumber -and $_.Index -eq $DocumentIndex) -or
            $_.VDL_ID -eq $VDL_Document.ID
        }
        # Previous FilterScript changed because item on PFSL could not be found if VDL item was not present (so couldn't match the ID) { $_.VDL_ID -eq $VDL_Document.ID }

        $PercentComplete = (88 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document Search')
            Status          = ('Checking results...')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting

        # Check if Document exists in Vendor Documents List
        If (-not $VDL_Document)
        {
            $ErrorMessage = ("Document '{0}' (Index {1}) not found in 'Vendor Documents List'." -f
                $DocumentNumber,
                $DocumentIndex
            )

            Switch ($ErrorActionPreference)
            {
                'SilentlyContinue'
                {
                    Write-Host $ErrorMessage -ForegroundColor Yellow
                    Break
                }

                Default
                { Throw $ErrorMessage }
            }
        }

        # Ensure only 1 Document is returned from Vendor Documents List
        If ($VDL_Document.Count -gt 1)
        {
            $SPOLinkToFilteredItems = Get-SPOItemsLink -SPOConnection $SPOConnection -ListName 'Vendor Documents List' -ItemIDs $VDL_Document.ID

            Throw ("More than 1 item found in 'Vendor Documents List' for '{0}' - Index {1}):{2}{3}" -f
                $DocumentNumber,
                $DocumentIndex,
                "`n",
                $SPOLinkToFilteredItems
            )
        }

        # Check if Document exists in Process Flow Status List
        If (-not $PFSL_Document)
        {
            $ErrorMessage = ("Document '{0}' (Index {1}) not found in 'Process Flow Status List'." -f
                $DocumentNumber,
                $DocumentIndex
            )

            Switch ($ErrorActionPreference)
            {
                'SilentlyContinue'
                {
                    Write-Host $ErrorMessage -ForegroundColor Yellow
                    Break
                }

                Default
                { Throw $ErrorMessage }
            }
        }

        # Ensure only 1 Document is returned from Process Flow Status List
        If ($PFSL_Document.Count -gt 1 -and -not $ErrorMessage)
        {
            $SPOLinkToFilteredItems = Get-SPOItemsLink -SPOConnection $SPOConnection -ListName 'Process Flow Status List' -ItemIDs $PFSL_Document.ID

            Throw ("More than 1 item found in 'Process Flow Status List' for '{0}' - Index {1}):{2}{3}" -f
                $DocumentNumber,
                $DocumentIndex,
                "`n",
                $SPOLinkToFilteredItems
                # Add link to vdl on filtered doc number / index / id
            )
        }

        # Ensure that Document Number inside the ticket is equal to the one in Vendor Documents List and in Process Flow Status List
        If (($VDL_Document.TCM_DN -ne $PFSL_Document.TCM_DN ) -or ($VDL_Document.Index -ne $PFSL_Document.Index ) -and -not $ErrorMessage)
        {
            Throw ("Document found in 'Vendor Documents List' ('{0}' - Index: {1}) is different from the one in 'Process Flow Status List' ('{2}' - Index: {3})" -f
                $VDL_Document.TCM_DN,
                $VDL_Document.Index,
                $PFSL_Document.TCM_DN,
                $PFSL_Document.Index)
        }

        # Return Document properties from both lists
        $VDM_Document = [PSCustomObject]@{
            FullDocumentNumber = $FullDocumentNumber
            VDL_Item           = $VDL_Document
            PFSL_Item          = $PFSL_Document
        }
        $PercentComplete = (100 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document Search')
            Status          = ('Search completed')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        Write-Progress -Activity 'VDM Document Search' -Completed -Id $CurrentProgressBarId
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Return $VDM_Document
    }
    Catch
    {
        Write-Progress -Activity 'VDM Document Search' -Completed -Id $CurrentProgressBarId
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Throw
    }
}

# Function to trigger a Power Automate HTTP flow
Function Invoke-PAHTTPFlow
{
    Param (
        # URI of the HTTP Request Trigger of the Flow
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Uri,

        # Method to be used to trigger the flow
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'PUT', 'POST', 'PATCH', 'DELETE')]
        [String]
        $Method,

        # JSON body to be passed to HTTP Request Trigger of the Flow
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Body
    )

    Try
    {
        # Create a new HTTP request to trigger the flow
        $Headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
        $Headers.Add('Content-Type', 'application/json')
        $Headers.Add('CharSet', 'charset=UTF-8')
        $Headers.Add('Accept', 'application/json')
        $EncodedBody = [System.Text.Encoding]::UTF8.GetBytes($Body)

        # Invoke the HTTP request
        $Response = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $EncodedBody

        If (-not $Response)
        {
            $ResubmissionType = ($Body | ConvertFrom-Json).ResubmissionType ?? 'Resubmission type not provided'
            $Response = [PSCustomObject]@{
                ResubmitActionStatusCode = $null
                ResubmitActionResult     = $null
                ResubmissionType         = $ResubmissionType
                LinkToResubmittedRun     = $null
                FlowResubmissionDateTime = $null
                LinkToRun                = $null
                FlowRunDateTime          = $null
            }
        }
        Return $Response
    }
    Catch { Throw }
}

# Function that converts a CSV file to an Excel file and adds hyperlinks to the cells containing URLs.
Function Convert-CSVToExcel
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $CSVPath
    )

    Try
    {
        # Determine the Excel path based on the CSV path
        $ExcelPath = [System.IO.Path]::ChangeExtension($CSVPath, 'xlsx')
        $CSVFileName = [System.IO.Path]::GetFileName($CSVPath)
        $ExcelFileName = [System.IO.Path]::GetFileName($ExcelPath)

        # Create Excel COM object
        $ExcelApp = New-Object -ComObject Excel.Application

        # Make Excel invisible
        $ExcelApp.Visible = $false

        # Disable alerts to suppress overwrite confirmation
        $ExcelApp.DisplayAlerts = $false

        # Open CSV
        $Workbook = $ExcelApp.Workbooks.Open($CSVPath)
        $Worksheet = $Workbook.Worksheets.Item(1)

        # Get the Range of data and create a table
        $Range = $Worksheet.UsedRange
        $Worksheet.ListObjects.Add(1, $Range, $null, 1) | Out-Null

        # Get the Columns with names containing 'Link'
        $LinkColumns = @()
        For ($i = 1; $i -le $Worksheet.UsedRange.Columns.Count; $i++)
        {
            If ($Worksheet.Cells.Item(1, $i).Text -like '*Link*')
            {
                $LinkColumns += $i
            }
        }

        # Find the last Row with data
        $LastRow = $Worksheet.UsedRange.Rows.Count

        # Iterate through link Columns and turn Cell Values into hyperlinks
        ForEach ($Column in $LinkColumns)
        {
            $ColumnName = $Worksheet.Cells.Item(1, $Column).Text # Getting the column name
            For ($Row = 2; $Row -le $LastRow; $Row++)
            {
                # Starting from Row 2 to skip the header
                $Cell = $Worksheet.Cells.Item($Row, $Column)
                $Value = $Cell.Text
                If (
                    $Value -match 'https?://\S+' -or
                    $Value -match 'http?://\S+' -or
                    $Value -match 'www\.\S+' -or
                    $Value -match 'mailto:\S+'
                )
                {
                    If ($Value.Length -le 2083)
                    {
                        $HyperLink = $Worksheet.Hyperlinks.Add($Cell, $Value)
                        Switch ($ColumnName)
                        {
                            'Link to Tickets'
                            {
                                # Customizing hyperlink text based on pattern
                                $LinkText = ((($Value -split 'numberIN')[1] -split '%255EORDERBY' | Select-Object -First 1) -split '%252C') -join ', '
                                $Hyperlink.TextToDisplay = $LinkText
                                Break
                            }
                            Default
                            {}
                        }
                    }
                }
            }
        }

        # Autofit the Columns and Wrap the Text
        $Range.WrapText = $false
        $Range.EntireColumn.AutoFit() | Out-Null

        # Attempt to save the file with retries
        $Success = $false
        While (-not $Success)
        {
            Try
            {
                # Save the Workbook
                $Workbook.SaveAs($ExcelPath, 51) # 51 = xlsx format
                $Success = $true
            }
            Catch
            {
                Write-Host ''
                Write-Warning ('Failed to save the Excel file. It may be in use by another process.')

                $ChoiceTitle = 'Excel file in use during CSV to XLSX conversion'
                $ChoiceMessage = ("File '{0}' is in use by another process.{1}Do you want to manually close the file and try again?" -f $ExcelFileName, "`n")
                $RetryChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Retry', ('{0}Retry saving after manually closing the file:{0}{1}{0}{0}' -f "`n", $ExcelPath)
                $TerminateChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Abort', ('{0}Abort the operation. It can be run later using this command:{0}{1}{0}' -f "`n", $MyInvocation.Line )
                $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($RetryChoice, $TerminateChoice)

                $Result = $Host.UI.PromptForChoice($ChoiceTitle, $ChoiceMessage, $Choices, 0)

                Switch ($Result)
                {
                    # Retry
                    0
                    {
                        Write-Host ''
                    }

                    # Abort
                    1
                    {
                        Write-Host ''
                        Write-Warning ('You choosed to abort the save process. To try again later, you can run this command:{0}{1}' -f "`n", $MyInvocation.Line)
                        Return $false
                    }
                }
            }

        }

        # Close Excel
        $ExcelApp.Quit()

        Write-Host ("{0}File '{1}' converted in '{2}'" -f "`n", $CSVFileName, $ExcelFileName) -ForegroundColor Green
        Return $true
    }
    Catch
    {
        Throw
    }
    Finally
    {
        # Clean up by releasing all COM objects created
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null

        # Force garbage collection to clean up any lingering objects
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function that returns the URL of a filtered List on one or more provided SharePoint list items
Function Get-SPOItemsLink
{
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection,

        [Parameter(Mandatory = $true)]
        [String]
        $ListName,

        [Parameter(Mandatory = $true)]
        [Int[]]
        $ItemIDs
    )

    Try
    {
        $ListObject = Get-PnPList -Identity $ListName -Includes ParentWeb -Connection $SPOConnection
        $ListFilter = ('FilterField{0}1=ID&FilterValue{0}1={1}&FilterType1=Counter' -f
            (($ItemIDs.Count -gt 1) ? 's' : $null),
            $($ItemIDs -Join '%3B%23')
        )
        $FilteredListItemsURL = ('{0}{1}{2}?{3}' -f
            $SPOConnection.Url,
            $ListObject.DefaultViewUrl.Substring(0, $ListObject.DefaultViewUrl.LastIndexOf('/')).Replace($ListObject.ParentWeb.ServerRelativeUrl, ''),
            '/AllItems.aspx',
            $ListFilter
        )

        Return $FilteredListItemsURL
    }
    Catch
    {
        Throw
    }
}

# Function that returns a Hyperlink clickable in the PowerShell console
# Add URL validation
Function New-PowerShellHyperlink
{
    Param(
        [Parameter(Mandatory = $true)]
        [String]
        $LinkURL,

        [Parameter(Mandatory = $true)]
        [String]
        $LinkDisplayText
    )

    Try
    {
        $PowerShellHyperlink = ("`e]8;;{0}`e\{1}`e]8;;`e\" -f $LinkURL, $LinkDisplayText)
        Return $PowerShellHyperlink
    }
    Catch
    {
        Throw
    }
}
#EndRegion Functions

#Region Remediations Functions

# Function to create JSON body for HTTP Flow remediation based on flow remediation name
Function New-JSONBodyForHTTPFlowRemediation
{
    Param(
        # Flow DisplayName
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $FlowDisplayName,

        # Name of the remediation action to perform
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $RemediationName,

        # Array of supported flows' DisplayNames
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Object]
        $SupportedFlowsAndRemediations,

        # Single PATicket object, memeber of $PATicketList returned by Get-PATicketList
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $PATicket,

        # Sharepoint Online Connection Object
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection,

        # Array of supported flows' DisplayNames
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $SpecificRemediationArguments
    )

    Try
    {
        # Check if flow is supported
        If ($SupportedFlowsAndRemediations.DisplayName -notcontains $FlowDisplayName)
        {
            Throw ("Flow '{0}' not supported" -f $FlowDisplayName)
        }

        # Get flow remediation
        $FlowRemediation = $SupportedFlowsAndRemediations |
            Where-Object -FilterScript {
                $_.DisplayName -eq $FlowDisplayName -and $_.Remediations.Name -eq $RemediationName
            } |
                Select-Object -ExpandProperty Remediations |
                    Where-Object -FilterScript { $_.Name -eq $RemediationName }


        # Create body based on flow remediation name
        Switch ($FlowRemediation.Name)
        {
            # Remediation for flow error 'DueDate cannot be earlier than the StartDate' in flow 'VDM - Site Independent - Document Disciplines Task Creation'
            'Invalid DueDate'
            {
                # Get required settings from Configuration List
                $FlowURL_TaskCreation = Search-SettingKey -SettingsListName 'Configuration List' -SettingName 'FlowUrl_TaskCreation' -SPOConnection $SPOConnection
                $Project_TeamsID = Search-SettingKey -SettingsListName 'Configuration List' -SettingName 'Project_TeamsID' -SPOConnection $SPOConnection
                $FlowURL_DisciplineNotifications = Search-SettingKey -SettingsListName 'Configuration List' -SettingName 'FlowURL_DisciplineNotifications' -SPOConnection $SPOConnection

                # Set variables and merge DisciplineOwnerTCM and DisciplinesTCM into a single array
                $VDL_Document = $Global:VDM_Document.VDL_Item
                $FullDocumentNumber = $Global:VDM_Document.FullDocumentNumber
                $DisciplinesArray = '"' + $VDL_Document.DisciplineOwnerTCM + '"'
                If (!(-not $VDL_Document.DisciplinesTCM)) { $DisciplinesArray += ',"' + $VDL_Document.DisciplinesTCM + '"' }

                # Create JSON body for 'Task Creation' Flow
                $TaskCreationFlowTriggerBody = '{
                    "ChosenDisciplines": [' + $DisciplinesArray + '],
                    "RootSiteUrl": "' + $PATicket.SiteUrl + '",
                    "ProjectTeamsID": "' + $Project_TeamsID + '",
                    "DocumentNumber": "' + $FullDocumentNumber + '",
                    "MRCode": "' + $VDL_Document.MRCode + '",
                    "PONumber": "' + $VDL_Document.PONumber + '",
                    "VendorSiteUrl": "' + $VDL_Document.VendorSiteUrl + '",
                    "VendorName": "' + $VDL_Document.VendorName + '",
                    "VDL_DocumentNumber": "' + $VDL_Document.TCM_DN + '",
                    "VDL_RevisionNumber": "' + $VDL_Document.Rev + '",
                    "EnglishDocumentTitle": "' + $VDL_Document.DocTitle + '",
                    "CommentsEndDate": "' + $((Get-Date).AddDays(1).ToString('yyyy-MM-dd')) + '",
                    "VDL_Index": ' + $VDL_Document.Index + ',
                    "VDL_ID": ' + $VDL_Document.ID + ',
                    "FlowURL_DisciplineNotifications": "' + $FlowURL_DisciplineNotifications + '"
                }' | ConvertTo-Json -Depth 100

                # Create JSON body for 'AMS - Resubmit Flow' Flow
                $Body = '{
                    "PATicketID": "' + $PATicket.PATicketID + '",
                    "EnvironmentID": "' + $PATicket.EnvironmentID + '",
                    "FlowID": "' + $PATicket.FlowID + '",
                    "RunID": "' + $PATicket.RunID + '",
                    "TriggerName": "' + $PATicket.TriggerName + '",
                    "ResubmissionType": "Manual",
                    "ManualResubmissionMethod": "POST",
                    "ManualResubmissionURI": "' + $FlowURL_TaskCreation + '",
                    "ManualResubmissionJSONBody": '+ $TaskCreationFlowTriggerBody + '
                }'

                # End switch
                Break
            }

            # Return error if remediation is not supported
            Default { Throw ("Remediation for flow '{0}' not supported" -f $FlowDisplayName) }
        }

        #Return output
        Return $Body
    }
    Catch { Throw }
}

<# Remediation for 'VDM - Site Independent - Document Disciplines Task Creation'
Used to remove a user that doesn't exist anymore on the tenant from the Disciplines list
#>
Function Remove-PAErrorUserFromDisciplines
{
    Param(
        # Sharepoint Online Connection Object
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection,

        # Error object from the PATicketFlow (obtained with Get-PATicketFlowErrorDetails function)
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Array]
        $ErrorMessages
    )

    Try
    {

        $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
        $CurrentProgressBarId = $ParentProgressBarId + 1
        $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
        $PercentComplete = (12 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('User removal')
            Status          = ("Loading list 'Disciplines'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting

        # Load Disciplines list in a PSCustomObject
        $DisciplinesList = Get-PnPListItem -List 'Disciplines' -Connection $SPOConnection -PageSize 5000 | ForEach-Object {
            [System.Collections.Generic.List[System.String]]$UserEmails = $_['VD_PersonID'].Email | ForEach-Object {
                If (!(-not $_)) { $_.ToLower() }
            }
            [PSCustomObject]@{
                ID         = $_['ID']
                Discipline = $_['Title']
                People     = $UserEmails
            }
        }

        # Loop through unique error messages
        $PercentComplete = (50 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('User removal')
            Status          = ('Removing users from list...')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        $Result = @()
        $ErrorMessages = $ErrorMessages | Where-Object -FilterScript { $_ -like 'Referenced User or Group (*) is not found.' } | Select-Object -Unique
        ForEach ($ErrorMessage in $ErrorMessages)
        {
            # Get the user's email address from the error message
            $UserEmailToRemove = (($ErrorMessage -split '\(')[1] -split '\)')[0].ToLower()

            # Filter disciplines list to only include only disciplines the user is a member of
            $UserDisciplines = $DisciplinesList | Where-Object { $_.People -contains $UserEmailToRemove }

            # Check if the user is still member of any disciplines
            If ( $UserDisciplines.Count -eq 0 )
            {
                $Result += ("User '{0}' is not a member of any disciplines" -f $UserEmailToRemove)
            }
            Else
            {
                # Remove the user from the disciplines he is a member of
                # Loop through the disciplines list and remove the user from each discipline
                ForEach ($Discipline in $UserDisciplines)
                {
                    # Remove user email from PSCustomObject
                    $Discipline.People.Remove($UserEmailToRemove) | Out-Null

                    # Update value in the list
                    Set-PnPListItem -List 'Disciplines' -Connection $SPOConnection -Identity $Discipline.ID -Values @{VD_PersonID = [Array]$Discipline.People } | Out-Null

                    # Compose output message
                    $Result += ("User '{0}' removed from discipline '{1}'{2}" -f
                        $UserEmailToRemove,
                        $Discipline.Discipline,
                        "`r"
                    )
                }
            }
        }
        $Result = $Result.Trim()
        $PercentComplete = (100 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('User removal')
            Status          = ('User removal completed')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        Write-Progress -Activity 'User removal completed' -Completed -Id $CurrentProgressBarId
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Return $Result
    }
    Catch
    {
        Write-Progress -Activity 'User removal completed' -Completed -Id $CurrentProgressBarId
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Throw
    }
}

# Function to delete a VDM Document. It deletes folder inside Vendor Subsite PO, list item from Revision folder Dashboard and Process Flow Status List
Function Remove-VDMDocument
{
    Param
    (
        # Vendor Documents document object as returned by function Get-VDMDocument
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $VDMDocument,

        # Sharepoint Online Connection Object
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection
    )

    Try
    {
        $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
        $CurrentProgressBarId = $ParentProgressBarId + 1
        $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
        $PercentComplete = (12 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document deletion')
            Status          = ('Connecting to Vendor subsite...')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        # Connect to the sub site
        $VDMSubSiteConnection = Connect-SPOSite -SiteUrl $($VDMDocument.PFSL_Item[0].VendorSiteUrl)

        Write-Host ('Cleaning leftovers for Document {0} - Rev {1}...' -f
            $($VDMDocument.PFSL_Item[0].TCM_DN),
            $($VDMDocument.PFSL_Item[0].Rev)
        ) -ForegroundColor Cyan

        # Check if folder for the document exists and delete it
        Write-Host 'Searching orphaned folder...' -ForegroundColor Cyan

        $VDMDocumentPath = "$($VDMDocument.PFSL_Item.SubSitePOURL)/$($VDMDocument.FullDocumentNumber)"
        $PercentComplete = (24 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document deletion')
            Status          = ('Cleaning folder...')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        $VDMDocFolder = Get-PnPFolder -Url $VDMDocumentPath -Connection $VDMSubSiteConnection -ErrorAction SilentlyContinue
        If ($null -eq $VDMDocFolder)
        {
            Write-Host ('Folder already not present.') -ForegroundColor DarkGreen
        }
        Else
        {
            # Delete orphaned document folder
            $VDMDocPathSplit = $VDMDocumentPath.Split('/')
            $VDMDocParentFolderPath = ($VDMDocPathSplit[6..($VDMDocPathSplit.Length - 2)] -join '/')
            Remove-PnPFolder -Name $($VDMDocument.FullDocumentNumber) -Folder $VDMDocParentFolderPath -Recycle -Force -Connection $VDMSubSiteConnection | Out-Null
            Write-Host ('Orphaned folder deleted.') -ForegroundColor DarkGreen
        }

        # Check if item for the document exists in Revision Folder Dashboard and delete it
        $PercentComplete = (50 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document deletion')
            Status          = ("Cleaning 'Revision Folder Dashboard'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        Write-Host ("Searching orphaned item in 'Revision Folder Dashboard'...") -ForegroundColor Cyan
        $RevisionFolderDashboardListItem = Get-AMSSPOList -ListName 'Revision Folder Dashboard' -SPOConnection $VDMSubSiteConnection -ForceOnlineList |
            Where-Object { $_.TCM_DN -eq $VDMDocument.PFSL_Item.TCM_DN -and $_.Rev -eq $VDMDocument.PFSL_Item.Rev }

        If ($RevisionFolderDashboardListItem.Count -eq 0)
        {
            Write-Host ("Document not found on 'Revision Folder Dashboard'.") -ForegroundColor DarkGreen
        }
        <# Replaced error with warning since these items need to be deleted even if duplicated
        ElseIf ($RevisionFolderDashboardListItem.Count -gt 1)
        {
            Throw ("More then one item found on 'Revision Folder Dashboard' for Document '{0}' (Rev: {1})." -f
                $($VDMDocument.PFSL_Item.TCM_DN),
                $($VDMDocument.PFSL_Item.Rev)
            )
        }
        #>
        Else
        {
            If ($RevisionFolderDashboardListItem.Count -gt 1)
            {
                Write-Host ("More then one item found on 'Revision Folder Dashboard' for Document '{0}' (Rev: {1}). Deleting all {2} items..." -f
                    $($VDMDocument.PFSL_Item.TCM_DN),
                    $($VDMDocument.PFSL_Item.Rev),
                    $RevisionFolderDashboardListItem.Count
                ) -ForegroundColor Yellow
            }
            ForEach ($Item in $RevisionFolderDashboardListItem)
            {
                Remove-PnPListItem -List 'Revision Folder Dashboard' -Identity $Item.ID -Recycle -Force -Connection $VDMSubSiteConnection | Out-Null
            }
            Write-Host ("Document deleted from 'Revision Folder Dashboard'.") -ForegroundColor DarkGreen
        }

        # Set document status to Deleted on Process Flow Status List and then delete the item
        $PercentComplete = (80 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document deletion')
            Status          = ("Cleaning 'Process Flow Status List'...")
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        If (!(-not $VDMDocument.PFSL_Item))
        {
            Write-Host ("Searching orphaned item in 'Process Flow Status List'...") -ForegroundColor Cyan
            ($VDMDocument.PFSL_Item.Count -gt 1) ? (Write-Host ("More then one item found on 'Process Flow Status List': {0}" -f ($VDMDocument.PFSL_Item.ID -Join ', ')) -ForegroundColor Yellow) : $null
            ForEach ($PFSL_Item in $VDMDocument.PFSL_Item)
            {
                Set-PnPListItem -List 'Process Flow Status List' -Identity $PFSL_Item.ID -Values @{'VD_DocumentStatus' = 'Deleted' } -Force -Connection $SPOConnection | Out-Null
                Remove-PnPListItem -List 'Process Flow Status List' -Identity $PFSL_Item.ID -Recycle -Force -Connection $SPOConnection | Out-Null
                $DeletedStatusMessageString = ($PFSL_Item.Status -eq 'Deleted') ? "already set to 'Deleted', item has just been deleted" : "has been set to 'Deleted', then the item has been deleted"
                Write-Host ("Document Status for item {0} on 'Process Flow Status List' {1}." -f $PFSL_Item.ID, $DeletedStatusMessageString) -ForegroundColor DarkGreen
            }
        }
        Else
        {
            Write-Host ("Item on 'Process Flow Status List' already not present.") -ForegroundColor DarkGreen
        }

        Write-Host ('Leftovers cleaning completed.') -ForegroundColor DarkGreen
        $PercentComplete = (100 / 100 * 100)
        $ProgressBarSplatting = @{
            Activity        = ('VDM Document deletion')
            Status          = ('Leftover cleaning completed')
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = $CurrentProgressBarId
            ParentId        = $ParentProgressBarId
        }
        Start-Sleep -Milliseconds 250
        Write-Progress @ProgressBarSplatting
        Write-Progress -Activity 'VDM Document deletion' -Completed -Id $CurrentProgressBarId
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
    }
    Catch
    {
        Write-Progress -Activity 'VDM Document deletion' -Completed -Id $CurrentProgressBarId
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Throw
    }
}
#EndRegion Remediations Functions

#Region Main
Try
{
    # Start PnP Trace Log
    Set-PnPTraceLog -On -LogFile $PnPTraceLogPath -Level Debug

    # Import Power Automate PowerShell module
    Import-Module -Name Microsoft.PowerApps.PowerShell -WarningAction SilentlyContinue

    # Validate the Supported Flows and Remediations object
    Confirm-SupportedFlowsAndRemediations -SupportedFlowsAndRemediations $SupportedFlowsAndRemediations

    # Get the list of tickets to resubmit the flow for
    $PATicketList = Get-PATicketList -PATicketCSVPath $PATicketCSVSourceFilePath -SupportedFlows $SupportedFlowsAndRemediations
    $LinkToTickets = ('https://tecnimont.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3DnumberIN{0}%255EORDERBYopened_at%255EGROUPBYshort_description%26sysparm_first_row%3D1%26sysparm_view%3D' -f ($PATicketList.PATicketID -join '%252C'))

    # Filter unsupported or not found flows from the list
    $PAUnsupportedTicketList = $PATicketList | Where-Object { $_.IsSupportedFlow -eq $false }

    # Uncomment the following line to pause the script after importing the tickets
    #Pause

    # Loop through the tickets and trigger the flow for each of them
    Write-Host ("`n`n==============================`nStarting to process tickets...`n==============================") -ForegroundColor DarkMagenta
    $Counter = 0
    ForEach ($PATicket in $PATicketList)
    {
        #Reset variables
        $PATicketStopwatch = $null
        $RemediationOutput = $null
        $RemediationActions = $null
        $Global:VDM_Document = $null
        #$Global:CommandError = $null
        $FlowNeedResubmission = $true
        $RemediationOutputString = $null
        $FlowResubmissionResponse = $null
        $PATicketFlowErrorDetails = $null
        $FlowHasRemediationActions = $null
        $ConsoleOutputForegroundColor = $null
        $FlowHasGenericRemediationActions = $null

        # Progess bar and $PATicketStopwatch
        $PATicketStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $Counter++
        $PercentComplete = ($Counter / $PATicketList.Count * 100)
        $ProgressBarSplatting = @{
            Activity        = ("Processing Flow '{0}'" -f $($PATicket.FlowDisplayName))
            Status          = ("Ticket: '{0}' - {1}/{2} ({3}%)" -f
                $($PATicket.PATicketID),
                $Counter,
                $PATicketList.Count,
                $([Math]::Round($PercentComplete))
            )
            PercentComplete = $([Math]::Round($PercentComplete))
            Id              = ([Array]$Global:ProgressBarsIds.Keys)[0]
        }
        Write-Progress @ProgressBarSplatting

        # Get the details of the Flow Error
        $PATicketFlowErrorDetails = Get-PATicketFlowErrorDetails -PATicketFlowDisplayName $PATicket.FlowDisplayName -PATicketFlowError $PATicket.FlowErrorDetails
        $AMSIdentifierString = ($PATicket.AMSIdentifier | Format-List | Out-String).Trim()

        # Create object with basic ticket properties for the output
        If (-not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorAction -and
            -not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorCode -and
            -not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorMessages
        )
        {
            $CSVOutputErrorDetails = 'N/A'
        }
        Else
        {
            $CSVOutputErrorDetails = $(($PATicketFlowErrorDetails.FlowErrorDetails | Select-Object -Unique -Property * | Format-List | Out-String).Trim(@("`n", "`r")).TrimEnd())
        }
        $Global:CSVOutput = [PSCustomObject]@{
            PATicket              = $PATicket.PATicketID
            FlowDisplayName       = $PATicket.FlowDisplayName #$(($PATicket.FlowDisplayName -split "`e")[2] -replace '\\')
            SiteUrl               = $PATicket.SiteUrl ?? 'N/A'
            FlowErrorDetails      = $CSVOutputErrorDetails
            PATicktDescription    = $($PATicket.PATicketDescription.Trim(@("`n", "`r")).TrimEnd())
            AMSIdentifier         = $AMSIdentifierString
            IsSupportedFlow       = $PATicket.IsSupportedFlow
            FlowID                = $PATicket.FlowID
            FlowRunLink           = $PATicket.FlowRunLink
            AdditionalDetails     = $null
            RemediationActionName = $null
            TicketProcessingTime  = $null
            #FlowResubmissionDateTime = $null
            #FlowRunDateTime          = $null
        }

        # Skip the Flow resubmission if it's not supported or not found
        If ($PATicket.PATicketID -in $PAUnsupportedTicketList.PATicketID)
        {
            # Write to console the current ticket details
            Write-Host ('{0}[PROCESSING {1}/{2}]{0}Ticket: {3}{0}Flow: {4}{0}Flow Run ID: {5}{0}Site Url: {6}{0}{7}{8}{9}' -f
                "`n",
                $Counter,
                $PATicketList.Count,
                $PATicket.PATicketHyperTextLink,
                $PATicket.FlowHyperTextLink, #("{0} ({1})" -f $PATicket.FlowHyperTextLink, $PATicket.FlowID),
                (New-PowerShellHyperlink -LinkURL $PATicket.FlowRunLink -LinkDisplayText $PATicket.RunID),
                $PATicket.SiteUrl,
                (($null -ne $PATicket.FlowErrorDetails) ? 'ERROR DETAILS' : 'MISSING ERROR DETAILS'),
                (($null -ne $PATicket.FlowErrorDetails) ? ('{0}{1}' -f "`n", $Global:CSVOutput.FlowErrorDetails) : $null),
                (('N/A' -ne $PATicket.AMSIdentifier.AMSIdentifierValue) ? ('{0}{1}' -f "`n", $AMSIdentifierString ) : $null)
            ) -ForegroundColor Blue

            # Set the returned flow name and the skip reason for console and log output
            If ($PATicket.FlowDisplayName -eq 'Not found')
            {
                $ReturnedFlowName = $PATicket.FlowID
                $SkipReason = 'NOT FOUND'
            }
            ElseIf ($PATicket.IsSupportedFlow -eq $false)
            {
                $ReturnedFlowName = $PATicket.FlowDisplayName
                $SkipReason = 'NOT SUPPORTED'
            }

            # Create object with fake FlowResubmissionResponse properties for the output
            $FlowResubmissionResponse = [PSCustomObject]@{
                ResubmitActionStatusCode = 'N/A'
                ResubmitActionResult     = $SkipReason
                ResubmissionType         = 'No'
                LinkToResubmittedRun     = 'N/A'
                FlowResubmissionDateTime = 'N/A'
                LinkToRun                = 'N/A'
                FlowRunDateTime          = 'N/A'
            }

            # Append to the output the response from the flow resubmission
            $FlowResubmissionResponse.PSObject.Properties | ForEach-Object {
                # Check if the property is already present in the output object, then add or update it
                If ($null -eq $Global:CSVOutput.$($_.Name))
                {
                    $Global:CSVOutput | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value
                }
                Else
                {
                    $Global:CSVOutput.$($_.Name) = $_.Value
                }
            }

            # Write to the console that the flow resubmission was skipped
            Write-Host ("{0}[SKIPPED] {1} - Flow '{2}' {3}" -f "`n", $PATicket.PATicketHyperTextLink, $ReturnedFlowName, $SkipReason) -ForegroundColor Yellow
            $TotalSkippedTicket++
        }
        Else
        {
            # If the Flow is found and supported, proceed with possible remediations and resubmission
            # Get Flow object from the list of Supported Flows matching the current ticket FlowDisplayName
            $CurrentPATicketFlow = $SupportedFlowsAndRemediations | Where-Object -FilterScript {
                $PATicket.FlowDisplayName -like "*$($_.DisplayName)*"
            }

            # Write to console the current ticket details
            Write-Host ('{0}[PROCESSING {1}/{2}]{0}Ticket: {3}{0}Flow: {4}{0}Flow Run ID: {5}{0}Site Url: {6}{0}{7}{8}{9}' -f
                "`n",
                $Counter,
                $PATicketList.Count,
                $PATicket.PATicketHyperTextLink,
                $PATicket.FlowHyperTextLink, #("{0} ({1})" -f $PATicket.FlowHyperTextLink, $PATicket.FlowID),
                (New-PowerShellHyperlink -LinkURL $PATicket.FlowRunLink -LinkDisplayText $PATicket.RunID),
                $PATicket.SiteUrl,
                (($null -ne $PATicket.FlowErrorDetails) ? 'ERROR DETAILS' : 'MISSING ERROR DETAILS'),
                (($null -ne $PATicket.FlowErrorDetails) ? ('{0}{1}' -f "`n", $Global:CSVOutput.FlowErrorDetails) : $null),
                (('N/A' -ne $PATicket.AMSIdentifier.AMSIdentifierValue) ? ('{0}{1}' -f "`n", $AMSIdentifierString ) : $null)
            ) -ForegroundColor Blue

            # Check if the Flow has one or more mapped remediation actions
            $FlowHasRemediationActions = !(-not $CurrentPATicketFlow.Remediations.Action)

            # Check if Flow has a 'ResubmissionPreventiveCheck' ScriptBlock property and, if yes, execute it to determine if the Flow should be resubmitted
            If (!(-not $CurrentPATicketFlow.PreventiveCheckAction))
            {
                If (!(-not $PATicket.SiteUrl))
                {
                    # Create SPOConnection to specified Site if not already established
                    <#
                        Temporary exception to split url when subsite is present instead of root site.
                        Comment when not needed.
                    #>
                    If ($PATicket.SiteUrl.Split('/').count -eq 6)
                    {
                        Write-Error -Message ("Subsite '{0}' is being processed instead of a main site. Continue?" -f $PATicket.SiteUrl.Split('/')) -ErrorAction Inquire
                        <#
                        $URL_Parts = $PATicket.SiteUrl -split "/"
                        $URL_Parts = $URL_Parts[0..($URL_Parts.Length - 2)]
                        $PATicket.SiteUrl = $URL_Parts -join "/"
                        #>
                    }
                    $SPOConnection = Connect-SPOSite -SiteUrl $PATicket.SiteUrl

                    # Run the 'PreventiveCheckAction' ScriptBlock property
                    $PreventiveCheckAction = [ScriptBlock]::Create($($($CurrentPATicketFlow.PreventiveCheckAction).ToString() -replace '(?m)^\s+').Trim(@(' ', "`n", "`r")))
                    Write-Host ("{0}Running 'PreventiveCheckAction' ScriptBlock property..." -f "`n") -ForegroundColor DarkMagenta

                    # Reset variables used in PreventiveCheckAction
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')] # Suppress IDE's linter warning for $ProgressBarStatus variable
                    $OutputToReturn = $null
                    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')] # Suppress IDE's linter warning for $ProgressBarStatus variable
                    $ProgressBarStatus = $null

                    $PreventiveCheckActionOutput = (& $PreventiveCheckAction)
                    Write-Progress -Activity 'Running PreventiveCheckAction' -Completed -Id ([Array]$Global:ProgressBarsIds.Keys)[0]
                    ($Global:ProgressBarsIds.Count -gt 1) ? $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1) : $null

                    If ($PreventiveCheckActionOutput -eq $False)
                    {
                        $FlowNeedResubmission = $False
                        $WarningMessage = ("Flow '{0}' will not be resubmitted because the 'PreventiveCheckAction' ScriptBlock property returned {1}." -f
                            $($CurrentPATicketFlow.DisplayName),
                            #$($PATicket.FlowHyperTextLink), # bad for csv
                            $PreventiveCheckActionOutput.ToString().ToUpper()
                        )
                        Write-Host $WarningMessage -ForegroundColor Yellow
                    }
                }
                Else
                {
                    $WarningMessage = ('{0}Missing Site URL on Ticket to create SharePoint Online Connection for PreventiveCheckAction.{0}Proceeding with automatic Flow resubmission.' -f "`n")
                    Write-Host $WarningMessage -ForegroundColor Yellow
                    $FlowNeedResubmission = $true
                }
                $Global:CSVOutput.AdditionalDetails += $WarningMessage + "`r"

            }

            # If $FlowNeedResubmission is still $True after skipping or running PreventiveCheckAction, check for Remediation Actions
            # Get remediation actions for the Flow
            $RemediationActions = Get-PATKFlowRemediationActions -PATicketFlowErrorDetails $PATicketFlowErrorDetails -SupportedFlowsAndRemediations $SupportedFlowsAndRemediations
            $FlowHasGenericRemediationActions = ($null -ne ($RemediationActions | Where-Object -FilterScript { $_.ErrorToRemediate -eq '*' }))
            If ($FlowNeedResubmission -eq $True)
            {
                <# If the flow has one or more remediation actions, but no Flow Error Details are available in the PATicket,
                simply warn that the flow will be resubmitted without remediation or skip it if the property 'ResubmitIfEmptyFlowErrors' is set to false.
            #>
                If (
                    $FlowHasRemediationActions -and
                    (-not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorMessages) -and
                    $FlowHasGenericRemediationActions -eq $false
                )
                {
                    If ($CurrentPATicketFlow.ResubmitIfEmptyFlowErrors -eq $true)
                    {
                        $FlowNeedResubmission = $true
                        $WarningMessage = ("Flow has mapped remediation actions but 'Flow Error Details' have not been found on ticket '{0}'.{1}Property 'ResubmitIfEmptyFlowErrors' is set to {2}, the Flow will be resubmitted without any remediation." -f
                            $($PATicket.PATicketHyperTextLink),
                            "`n",
                            $($CurrentPATicketFlow.ResubmitIfEmptyFlowErrors.ToString().ToUpper())
                        )
                        $Global:CSVOutput.AdditionalDetails += $WarningMessage
                        Write-Host $WarningMessage -ForegroundColor Yellow
                    }
                    Else
                    {
                        $FlowNeedResubmission = $false
                        $WarningMessage = ("Flow has mapped remediation actions but 'Flow Error Details' have not been found on ticket '{0}'.{1}Property 'ResubmitIfEmptyFlowErrors' is set to {2}, the Flow won't be resubmitted." -f
                            $($PATicket.PATicketHyperTextLink),
                            "`n",
                            $($CurrentPATicketFlow.ResubmitIfEmptyFlowErrors.ToString().ToUpper())
                        )
                        $Global:CSVOutput.AdditionalDetails += $WarningMessage
                        Write-Host $WarningMessage -ForegroundColor Yellow
                    }
                }
                Else
                {
                    # Run remediation actions for the flow
                    # Filter Missing remediation actions
                    [Array]$MissingRemediationActions = $RemediationActions | Where-Object -FilterScript { $_.Action -eq 'Missing remediation' }

                    # If the flow has one or more Missing remediation actions, skip the PATicket unless property 'ResubmitIfUnsupportedFlowErrors' is set to true
                    If (
                        $MissingRemediationActions.Count -gt 0 -and
                        $CurrentPATicketFlow.ResubmitIfUnsupportedFlowErrors -ne $true
                    )
                    {
                        If (-not $FlowResubmissionResponse)
                        {
                            $FlowResubmissionResponse = [PSCustomObject]@{
                                ResubmitActionStatusCode = 'N/A'
                                ResubmitActionResult     = $MissingRemediationActions.Action
                                ResubmissionType         = 'No'
                                LinkToResubmittedRun     = 'N/A'
                                FlowResubmissionDateTime = 'N/A'
                                LinkToRun                = 'N/A'
                                FlowRunDateTime          = 'N/A'
                            }

                            # Append to the output the response from the flow resubmission
                            $FlowResubmissionResponse.PSObject.Properties | ForEach-Object {
                                # Check if the property is already present in the output object, then add or update it
                                If ($null -eq $Global:CSVOutput.$($_.Name))
                                {
                                    $Global:CSVOutput | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value
                                }
                                Else
                                {
                                    $Global:CSVOutput.$($_.Name) = $_.Value
                                }
                            }
                        }

                        $FlowNeedResubmission = $false
                        $FlowResubmissionResponse.ResubmissionType = 'No'
                        $WarningMessage = ("{0}Flow Error Details contain one or more unsupported error, the Flow won't be resubmitted because property 'ResubmitIfUnsupportedFlowErrors' is set to {1}.{0}Missing remediation actions:{0}{2}" -f
                            "`n",
                            $CurrentPATicketFlow.ResubmitIfUnsupportedFlowErrors.ToString().ToUpper(),
                        ($MissingRemediationActions | Select-Object -Property ErrorToRemediate, Action | Format-List | Out-String).Trim(@("`n", "`r")).TrimEnd()
                        )
                        $Global:CSVOutput.AdditionalDetails += $WarningMessage
                        $Global:CSVOutput.FlowResubmissionDateTime = 'N/A'
                        $Global:CSVOutput.FlowRunDateTime = 'N/A'
                        Write-Host $WarningMessage -ForegroundColor Yellow
                    }

                    # Run remediation action if needed
                    If (
                        !(-not $RemediationActions) -and
                        $RemediationActions.Action -ne 'N/A' -and
                        $FlowNeedResubmission
                    )
                    {
                        ForEach ($RemediationAction in $RemediationActions)
                        {
                            # Run remediation action if supported
                            If ($RemediationAction -notin $MissingRemediationActions)
                            {
                                # Write output to console
                                Write-Host ("{0}Remediation '{1}' found.{0}Running remediation action '{1}'..." -f
                                    "`n",
                                    $RemediationAction.Name
                                ) -ForegroundColor DarkMagenta

                                # Check if the remediation action requires a connection to SharePoint Online SiteURL provided in the PATicket
                                $IsSPOConnectionRequired = $null
                                $IsSPOConnectionRequired = ($CurrentPATicketFlow.Remediations | Where-Object -FilterScript { $_.Name -eq $RemediationAction.Name }).IsSPOConnectionRequired
                                If ($IsSPOConnectionRequired -eq $true)
                                {
                                    # Check if the SharePoint Online SiteURL is present in the PATicket description
                                    If (-not $PATicket.SiteUrl)
                                    {
                                        If ($FlowHasGenericRemediationActions -ne $true)
                                        {
                                            Throw ("[ERROR] {0} - Required property 'Site URL' for Remediation Action '{1}' of Flow '{2}' could not be found inside ticket description." -f
                                                $PATicket.PATicketHyperTextLink,
                                                $($RemediationAction.Name),
                                                $($CurrentPATicketFlow.DisplayName)
                                                #$($PATicket.FlowHyperTextLink) # bad for csv

                                            )
                                        }
                                    }
                                    Else
                                    {
                                        # Create SPOConnection to specified Site if not already established
                                        $SPOConnection = Connect-SPOSite -SiteUrl $PATicket.SiteUrl
                                    }
                                }

                                # Skip remediation action if the flow has generic remediation actions
                                If ($FlowHasGenericRemediationActions -eq $true -and -not $PATicket.SiteUrl)
                                {
                                    Write-Host ('Missing Site URL on Ticket to create SharePoint Online Connection for RemediationOutput.{0}Proceeding with automatic Flow resubmission.' -f "`n") -ForegroundColor Yellow
                                }
                                Else
                                {
                                    #Run remediation action
                                    $RemediationOutput = (& $RemediationAction.Action)
                                    if ($RemediationOutput.RemediationActionOutput -eq 'SKIPPED') { continue }
                                    Write-Progress -Activity 'Running PreventiveCheckAction' -Completed -Id ([Array]$Global:ProgressBarsIds.Keys)[0]
                                    ($Global:ProgressBarsIds.Count -gt 1) ? $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1) : $null
                                    $RemediationOutputString = ($RemediationOutput | Convert-PSCustomObjectToList | Out-String).TrimEnd()
                                }

                                # Set property to track if the Flow still needs to be resubmitted
                                If ($null -eq $RemediationOutput.FlowNeedResubmission)
                                {
                                    $FlowNeedResubmission = $true
                                }
                                Else
                                {
                                    $FlowNeedResubmission = $RemediationOutput.FlowNeedResubmission
                                }

                                # Append remediation action name and its output to the output to be exported to CSV
                                $Global:CSVOutput.AdditionalDetails += $RemediationOutputString + "`r"
                                $Global:CSVOutput.RemediationActionName += $RemediationAction.Name + "`r"

                                If (
                                    (
                                        $RemediationAction.IsSPOConnectionRequired -eq $true -and
                                        !(-not $PATicket.SiteUrl)
                                    ) -or
                                    $RemediationAction.IsSPOConnectionRequired -ne $true
                                )
                                {
                                    # Return error if remediation action returns null output, otherwise return the output
                                    If (-not $RemediationOutputString)
                                    {
                                        Throw ("Remediation action failed for ticket '{0}'. Null output returned." -f $($PATicket.PATicketHyperTextLink))
                                    }
                                    Else
                                    {
                                        Write-Host 'Remediation action output:' -ForegroundColor DarkMagenta
                                        Write-Host $RemediationOutputString -ForegroundColor Cyan
                                    }
                                }

                                # Break the loop if the flow doesn't need more resubmissions
                                If (!($FlowNeedResubmission))
                                {
                                    # Compose console output details
                                    If (
                                        -not $RemediationOutput.FlowResubmissionResponse.ResubmitActionResult -and
                                        $RemediationOutput.FlowResubmissionResponse.ResubmissionType -eq 'Manual'
                                    )
                                    {
                                        $RemediationOutput.FlowResubmissionResponse.ResubmitActionResult = 'Unhandled exception'
                                    }
                                    Break
                                }
                            }
                            Else
                            {
                                # Warn that Flow will be automatically resubmitted because property 'ResubmitIfUnsupportedFlowErrors' is set to $True
                                $WarningMessage = Write-Host ("{0}Flow Error Details contain one or more unsupported error.{0}Missing remediation actions:{0}{1}{0}The Flow will be automatically resubmitted because property 'ResubmitIfUnsupportedFlowErrors' is set to {2}." -f
                                    "`n",
                                    ($MissingRemediationActions | Select-Object -Property * | Format-List | Out-String).Trim(@("`n", "`r")).TrimEnd(),
                                    $CurrentPATicketFlow.ResubmitIfUnsupportedFlowErrors.ToString().ToUpper()
                                )
                                $Global:CSVOutput.AdditionalDetails += $WarningMessage
                                Write-Host $WarningMessage -ForegroundColor Yellow
                            }
                        }
                    }
                    Else
                    {
                        $Global:CSVOutput.RemediationActionName = 'N/A'
                    }
                }
            }

            # Compose output if resubmission is not needed or has already been performed manually
            If (
                (
                    (
                        !(-not $CurrentPATicketFlow.PreventiveCheckAction) -and
                        $PreventiveCheckActionOutput -eq $False
                    ) -or
                    $FlowNeedResubmission -eq $False
                ) -or
                $RemediationOutput.RemediationActionOutput -eq 'SKIPPED'
            )
            {
                # Create object with FlowResubmissionResponse properties for the output
                $FlowResubmissionResponse = [PSCustomObject]@{
                    ResubmitActionStatusCode = $($RemediationOutput.FlowResubmissionResponse.ResubmitActionStatusCode) ?? 'N/A'
                    ResubmitActionResult     = $($RemediationOutput.FlowResubmissionResponse.ResubmitActionResult) ?? 'Skipped'
                    ResubmissionType         = $($RemediationOutput.FlowResubmissionResponse.ResubmissionType) ?? 'No'
                    LinkToResubmittedRun     = $($RemediationOutput.FlowResubmissionResponse.LinkToResubmittedRun) ?? 'N/A'
                    FlowResubmissionDateTime = $($RemediationOutput.FlowResubmissionResponse.FlowResubmissionDateTime) ?? 'N/A'
                    LinkToRun                = $($RemediationOutput.FlowResubmissionResponse.LinkToResubmittedRun) ?? 'N/A'
                    FlowRunDateTime          = $($RemediationOutput.FlowResubmissionResponse.FlowRunDateTime) ?? 'N/A'
                }
            }
            Else
            {
                # Run automatic resubmission if still needed
                # Resubmit the Flow if it has not been already resubmitted manually
                If ($FlowNeedResubmission -eq $true)
                {
                    # Body of the flow resubmission request
                    #$ResubmissionType = 'Automatic'
                    $AMSResubmitFlowBody = '{
                        "PATicketID": "' + $PATicket.PATicketID + '",
                        "EnvironmentID": "' + $PATicket.EnvironmentID + '",
                        "FlowID": "' + $PATicket.FlowID + '",
                        "RunID": "' + $PATicket.RunID + '",
                        "TriggerName": "' + $PATicket.TriggerName + '",
                        "ResubmissionType": "Automatic"
                    }'

                    # Trigger the flow 'AMS - Resubmit flow' to resubmit the flow from PATicket with given parameters
                    Write-Host ('{0}Running automatic resubmission...' -f "`n") -ForegroundColor DarkMagenta
                    $FlowResubmissionResponse = Invoke-PAHTTPFlow -Uri $AMSResubmitFlowUri -Body $AMSResubmitFlowBody -Method POST
                    Write-Host (($FlowResubmissionResponse | Convert-PSCustomObjectToList | Out-String).TrimEnd()) -ForegroundColor Cyan
                }
                Else
                {
                    # Set the flow resubmission response to the one returned by the remediation action
                    $FlowResubmissionResponse = $RemediationOutput.FlowResubmissionResponse
                }
            }

            # Append to the output the response from the flow resubmission
            [Array]$CSVOutputProperties = ($Global:CSVOutput | Get-Member -MemberType NoteProperty).Name
            $FlowResubmissionResponse.PSObject.Properties | ForEach-Object {
                # Check if the property is already present in the output object, then add or update it
                #If ($null -eq $Global:CSVOutput.$($_.Name))
                If ($_.Name -notin $CSVOutputProperties)
                {
                    $Global:CSVOutput | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value
                }
                Else
                {
                    $Global:CSVOutput.$($_.Name) = $_.Value
                }
            }

            # Set the color of the console output based on the resubmission result
            Switch ($FlowResubmissionResponse.ResubmitActionResult.ToUpper())
            {
                'SUCCESS'
                {
                    $TotalProcessedTicket++
                    $ForegroundColor = 'Green'
                    Break
                }

                'SKIPPED'
                {
                    $TotalProcessedTicket++
                    $ForegroundColor = 'Yellow'
                    Break
                }

                { $_ -in ('FAILED', 'UNHANDLED EXCEPTION') }
                {
                    $TotalSkippedTicket++
                    $ForegroundColor = 'Red'
                    Break
                }

                Default
                {
                    $ForegroundColor = 'Red'
                }
            }
            $ConsoleOutputForegroundColor = @{
                ForegroundColor = $ForegroundColor
            }

            # Write to the console flow resubmission result
            Write-Host ("{0}[{1}] {2} - Flow '{3}' ({4} resubmission)" -f
                "`n",
                $FlowResubmissionResponse.ResubmitActionResult.ToUpper(),
                $PATicket.PATicketHyperTextLink,
                $PATicket.FlowHyperTextLink,
                $FlowResubmissionResponse.ResubmissionType
            ) @ConsoleOutputForegroundColor
        }
        $TotalSkippedTicket++

        $PATicketStopwatch.Stop()
        $PATicketElapsedTime = $(Get-Date -Date $($PATicketStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
        $PAticket.PATicketResolutionTime = $PATicketElapsedTime
        $Global:CSVOutput.TicketProcessingTime = $PATicketElapsedTime
        # Add FlowResubmissionDateTime and FlowRunDateTime properties and export the output to the log file
        #$Global:CSVOutput.OperationDateTime = $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
        $Global:CSVOutput.FlowResubmissionDateTime = $FlowResubmissionResponse.FlowResubmissionDateTime ?? 'N/A'
        $Global:CSVOutput.FlowRunDateTime = $FlowResubmissionResponse.FlowRunDateTime ?? 'N/A'
        (-not $Global:CSVOutput.AdditionalDetails) ? ($Global:CSVOutput.AdditionalDetails = 'N/A') : $null | Out-Null
        (-not $Global:CSVOutput.RemediationActionName) ? ($Global:CSVOutput.RemediationActionName = 'N/A') : $null | Out-Null
        $Global:CSVOutput | Export-Csv -Path $CSVLogPath -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8BOM -UseQuotes Always

        # Write to the console the elapsed time to handle the PATicket
        $ScriptElapsedTime = $(Get-Date -Date $($ScriptStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
        Write-Host ('{0}Ticket processing execution time: {1}' -f "`n", $PATicketElapsedTime) -ForegroundColor Green
        Write-Host ('Current script execution time: {0}' -f $ScriptElapsedTime) -ForegroundColor Green

        # Uncomment the following line to pause the script after each flow resubmission
        #Pause
    }

    Write-Progress -Activity 'Resubmitting flows' -Completed -Id ([Array]$Global:ProgressBarsIds.Keys)[0]
}
Catch
{
    $ScriptError = $true

    # Stop the Stopwatch if exists
    If ($PATicketStopwatch)
    {
        $PATicketStopwatch.Stop()
    }

    # Compose the error message
    $CaughtError = ('Error:{0}' -f ($_ | Out-String)).TrimEnd()

    # Get the details of the Flow Error from the PATicket if available
    If (-not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorAction -and
        -not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorCode -and
        -not $PATicketFlowErrorDetails.FlowErrorDetails.ErrorMessages
    )
    {
        $CSVOutputErrorDetails = 'N/A'
    }
    Else
    {
        $ReturnedFlowErrorDetails = $PATicketFlowErrorDetails ?? $PATicket.FlowErrorDetails ?? ''
        Switch ($ReturnedFlowErrorDetails.GetType().Name)
        {
            'PSCustomObject'
            {
                $CSVOutputErrorDetails = $(($PATicketFlowErrorDetails.FlowErrorDetails | Select-Object -Unique -Property * | Format-List | Out-String).Trim(@("`n", "`r")).TrimEnd())
                Break
            }

            'Object[]'
            {
                $CSVOutputErrorDetails = $(($PATicket.FlowErrorDetails | Format-List | Out-String).Trim(@("`n", "`r")).TrimEnd())
                Break
            }

            Default
            { $CSVOutputErrorDetails = 'Script error getting property' }
        }
    }

    # Create object with basic ticket properties for the output
    $Global:CSVOutput = [PSCustomObject]@{
        PATicket              = $PATicket.PATicketID ?? 'Script error getting property'
        FlowDisplayName       = $PATicket.FlowDisplayName #$(($PATicket.FlowDisplayName -split "`e")[2] -replace '\\') ?? 'Script error getting property'
        SiteUrl               = $PATicket.SiteUrl ?? 'Script error getting property'
        FlowErrorDetails      = $CSVOutputErrorDetails
        PATicktDescription    = $($PATicket.PATicketDescription?.Trim(@("`n", "`r"))?.TrimEnd()) ?? 'Script error getting property'
        AMSIdentifier         = $AMSIdentifierString ?? 'Script error getting property'
        IsSupportedFlow       = $PATicket.IsSupportedFlow ?? 'Script error getting property'
        FlowID                = $PATicket.FlowID ?? 'Script error getting property'
        FlowRunLink           = $PATicket.FlowRunLink ?? 'Script error getting property'
        AdditionalDetails     = $CaughtError ?? 'Script error getting property'
        RemediationActionName = $PATicket.RemediationActionName ?? 'Script error getting property'
        TicketProcessingTime  = 'N/A'
    }

    # Append to the output the response from the flow resubmission
    If ($null -eq $FlowResubmissionResponse)
    {
        # Create object with FlowResubmissionResponse properties for the output
        $FlowResubmissionResponse = [PSCustomObject]@{
            ResubmitActionStatusCode = $($RemediationOutput.FlowResubmissionResponse.ResubmitActionStatusCode) ?? 'N/A'
            ResubmitActionResult     = $($RemediationOutput.FlowResubmissionResponse.ResubmitActionResult) ?? 'Skipped'
            ResubmissionType         = $($RemediationOutput.FlowResubmissionResponse.ResubmissionType) ?? 'N/A'
            LinkToResubmittedRun     = $($RemediationOutput.FlowResubmissionResponse.LinkToResubmittedRun) ?? 'N/A'
            FlowResubmissionDateTime = $($RemediationOutput.FlowResubmissionResponse.FlowResubmissionDateTime) ?? 'N/A'
            LinkToRun                = $($RemediationOutput.FlowResubmissionResponse.LinkToResubmittedRun) ?? 'N/A'
            FlowRunDateTime          = $($RemediationOutput.FlowResubmissionResponse.FlowRunDateTime) ?? 'N/A'
        }
    }

    # Append to the output the response from the flow resubmission
    [Array]$CSVOutputProperties = ($Global:CSVOutput | Get-Member -MemberType NoteProperty).Name
    $FlowResubmissionResponse.PSObject.Properties | ForEach-Object {
        # Check if the property is already present in the output object, then add or update it
        #If ($null -eq $Global:CSVOutput.$($_.Name))
        If ($_.Name -notin $CSVOutputProperties)
        {
            $Global:CSVOutput | Add-Member -NotePropertyName $_.Name -NotePropertyValue $_.Value
        }
        Else
        {
            $Global:CSVOutput.$($_.Name) = $_.Value
        }
    }

    # Export the output to the log file
    $Global:CSVOutput | Export-Csv -Path $CSVLogPath -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8BOM

    # Return the error
    Write-Host $CaughtError -ForegroundColor Red

    # Stop all progress bars
    $([Array]$Global:ProgressBarsIds.Keys) | Sort-Object | ForEach-Object {
        Write-Progress -Activity '*' -Completed -Id $_
    }
}
Finally
{
    Write-Host ("`n`n==============================`nFinished to process tickets...`n==============================") -ForegroundColor DarkMagenta
    Write-Host ('{0} of {1} processed tickets' -f $TotalProcessedTicket, $PATicketList.Count) -ForegroundColor Green

    If (Test-Path -Path $Global:TmpFolderPath -PathType Container)
    {
        Write-Host "`nClearing temporary files..." -ForegroundColor Cyan
        Remove-Item -Path $Global:TmpFolderPath -Force -Recurse
    }
    Set-PnPTraceLog -Off
    $ScriptStopwatch.Stop()
    $ScriptElapsedTime = $(Get-Date -Date $($ScriptStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
    Write-Host ('Total script execution time: {0}' -f $ScriptElapsedTime) -ForegroundColor Green # Change to Red if script error
    $ScriptExecutionEndDate = (Get-Date -Format 'dd/MM/yyyy - HH:mm:ss')
    Write-Host ('Script execution end date and time: {0}' -f $ScriptExecutionEndDate) -ForegroundColor Green

    # Write to the generic log file the details of the script execution
    $ExecutionsLogsDetails = [PSCustomObject]@{
        'Total Input Tickets'       = $PATicketList.Count
        'Total Unprocessed Tickets' = !($ScriptError) ? 0 : $($PATicketList.Count - $($Counter - 1))
        'Total Processed Tickets'   = $TotalProcessedTicket
        'Total Skipped Tickets'     = $TotalSkippedTicket
        'Total Resolution Time'     = $ScriptElapsedTime
        'Processing Start Date'     = $ScriptExecutionStartDate
        'Processing End Date'       = $ScriptExecutionEndDate
        'Link to Tickets'           = $LinkToTickets
    }

    # Group the tickets by FlowDisplayName
    $GroupedPATicketList = $PATicketList |
        Where-Object -FilterScript { $_.PATicketResolutionTime -ne $null -and $_.SupportedFlowDisplayName -ne 'N/A' } |
            Group-Object -Property SupportedFlowDisplayName

    # Iterate through each group and calculate the average PATicketResolutionTime
    ForEach ($PAGroup in $GroupedPATicketList)
    {
        $PATicketAverageResolutionSeconds = (
            $PAGroup.Group |
                ForEach-Object { [TimeSpan]::Parse($_.PATicketResolutionTime).TotalSeconds } |
                    Measure-Object -Average).Average
        $PATicketAverageResolutionTime = [TimeSpan]::FromSeconds($PATicketAverageResolutionSeconds).ToString('hh\:mm\:ss\.ff')
        $ExecutionsLogsDetails | Add-Member -MemberType NoteProperty -Name $PAGroup.Name -Value $PAGroup.Count
        $ExecutionsLogsDetails | Add-Member -MemberType NoteProperty -Name "'$($PAGroup.Name)' - Average Resolution Time" -Value $PATicketAverageResolutionTime
    }
    $SupportedFlowsAndRemediations.DisplayName | ForEach-Object {
        If ($ExecutionsLogsDetails.$_ -eq $null)
        {
            $ExecutionsLogsDetails | Add-Member -MemberType NoteProperty -Name $_ -Value 0
            $ExecutionsLogsDetails | Add-Member -MemberType NoteProperty -Name "'$_' - Average Resolution Time" -Value 'N/A'
        }
    }

    $ExecutionsLogPath = "$($LogRootFolder)\$($ScriptName)_ExecutionsLog_$($SupportedFlowsAndRemediations.Count)-SupportedFlows.csv"
    $ExecutionsLogsDetails | Export-Csv -Path $ExecutionsLogPath -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8BOM
    Write-Host ("`nConverting CSV reports to Excel...") -ForegroundColor Cyan
    Convert-CSVToExcel -CSVPath $ExecutionsLogPath | Out-Null
    Convert-CSVToExcel -CSVPath $CSVLogPath | Out-Null
    Stop-Transcript | Out-Null

    If ($ScriptError)
    {
        Exit 1
    }
    Else
    {
        Exit 0
    }
}
#EndRegion Main