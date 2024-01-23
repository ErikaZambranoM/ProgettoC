#Requires -Version 7

<#
TODO:
- Console output for auth required or not
- Add login timeout (check snippet)
    - Add Az Environment parameter
    - Excel conversion?
    - Add -Verbose switch to print details about the script execution
        OR ? - Add -Debug switch to print details about the script execution
    - Add -Confirm switch to show total of run to get details for and ask for confirmation before proceeding
    - Array of FlowRun ID filter parameter
#>

Param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            If ([GUID]::TryParse($_, $([ref][guid]::Empty)))
            {
                Return $true
            }
            Else
            {
                Throw "`nInvalid Flow ID: $_"
            }
        })]
    [String]
    $FlowId,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
            ForEach ($Item in $_)
            {
                If (
                    !([GUID]::TryParse($Item.Replace('Default-', ''), $([ref][guid]::Empty))) -and
                    !([GUID]::TryParse($Item, $([ref][guid]::Empty)))
                )
                {
                    Throw "`nInvalid Environment ID: $Item"
                }
                Return $true
            }
        })]
    [String[]]
    $EnvironmentIDs,

    [Parameter(Mandatory = $true)]
    [ValidateSet('*', 'All', 'Succeeded', 'Failed', 'Cancelled', 'Running)')]
    [AllowEmptyCollection()]
    [String[]]
    $StatusFilter
)

# Tenant ID required to authenticate to Azure AD (required to get the flow error details)
$TenantId = '7cc91888-5aa0-49e5-a836-22cda2eae0fc'

#Region Functions

# Function that returns a DateTime object from a string
Function Get-DateTimeFromString
{
    Param(
        [Parameter(Mandatory = $true)]
        [String]
        $DateTimeString,

        [Parameter(Mandatory = $false)]
        [Switch]
        $ClosestEnd
    )

    Try
    {
        # Declare the date formats to be supported
        $DateFormats = [String[]]@(
            # Year only
            'yyyy'

            # Month and year
            'M/yyyy',
            'MM/yyyy',

            # Day, month and year
            'd/M',
            'dd/M',
            'd/MM',
            'dd/MM',
            'd/M/yyyy',
            'dd/M/yyyy',
            'd/MM/yyyy',
            'dd/MM/yyyy',

            # Day, month, year and hour
            'd/M HH',
            'dd/M HH',
            'd/MM HH',
            'dd/MM HH',
            'd/M/yyyy HH',
            'dd/M/yyyy HH',
            'd/MM/yyyy HH',
            'dd/MM/yyyy HH',
            'd/M/yyyy H',
            'dd/M/yyyy H',
            'd/MM/yyyy H',
            'dd/MM/yyyy H',

            # Day, month, year, hour and minute
            'd/M H:m',
            'dd/M H:m',
            'dd/MM H:m',
            'd/MM H:m',
            'd/M HH:m',
            'dd/M HH:m',
            'dd/MM HH:m',
            'd/MM HH:m',
            'd/M H:mm',
            'dd/M H:mm',
            'dd/MM H:mm',
            'd/MM H:mm',
            'd/M HH:mm',
            'dd/M HH:mm',
            'dd/MM HH:mm',
            'd/MM HH:mm',
            'd/M/yyyy H:m',
            'dd/M/yyyy H:m',
            'd/MM/yyyy H:m',
            'dd/MM/yyyy H:m',
            'd/M/yyyy HH:m',
            'dd/M/yyyy HH:m',
            'd/MM/yyyy HH:m',
            'dd/MM/yyyy HH:m',
            'd/M/yyyy H:mm',
            'dd/M/yyyy H:mm',
            'd/MM/yyyy H:mm',
            'dd/MM/yyyy H:mm',
            'd/M/yyyy HH:mm',
            'dd/M/yyyy HH:mm',
            'd/MM/yyyy HH:mm',
            'dd/MM/yyyy HH:mm',

            # Day, month, year, hour, minute and second
            'd/M/yyyy HH:m:s',
            'dd/M/yyyy HH:m:s',
            'd/MM/yyyy HH:m:s',
            'dd/MM/yyyy HH:m:s',
            'd/M/yyyy H:m:s',
            'dd/M/yyyy H:m:s',
            'd/MM/yyyy H:m:s',
            'dd/MM/yyyy H:m:s',
            'd/M HH:m:s',
            'dd/M HH:m:s',
            'd/MM HH:m:s',
            'dd/MM HH:m:s',
            'd/M HH:mm:s',
            'dd/M HH:mm:s',
            'd/MM HH:mm:s',
            'dd/MM HH:mm:s',
            'd/M H:mm:s',
            'dd/M H:mm:s',
            'd/MM H:mm:s',
            'dd/MM H:mm:s',
            'dd/MM HH:m:ss',
            'dd/MM HH:mm:ss',
            'dd/MM/yyyy HH:mm:ss'
        )

        # Convert the date string to a DateTime object
        $DateTime = [DateTime]::MinValue
        If ([DateTime]::TryParseExact($DateTimeString, $DateFormats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$DateTime))
        {
            $DateTime = $DateTime
        }
        Else
        {
            Throw "Unable to convert unsupported date format for date string '$DateTimeString'"
        }

        # If the closest end switch is specified, then convert the date to the end of the period
        If ($ClosestEnd)
        {
            # Determine the matching format
            ForEach ($Format in $DateFormats)
            {
                If ([DateTime]::TryParseExact($DateTimeString, $Format, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$DateTime))
                {
                    $MatchingFormat = $Format
                    Break
                }
            }

            # Add the appropriate amount of time to the date
            Switch ($MatchingFormat)
            {
                # Year only
                'yyyy'
                {
                    $DateTime = $DateTime.AddYears(1).AddSeconds(-1)
                    Break
                }

                # Month and year
                { $_ -in ('M/yyyy', 'MM/yyyy') }
                {
                    $DateTime = $DateTime.AddMonths(1).AddSeconds(-1)
                    Break
                }

                # Day, month and year
                { $_ -in (
                        'd/M',
                        'dd/M',
                        'd/MM',
                        'dd/MM',
                        'd/M/yyyy',
                        'dd/M/yyyy',
                        'd/MM/yyyy',
                        'dd/MM/yyyy'
                    )
                }
                {
                    $DateTime = $DateTime.AddDays(1).AddSeconds(-1)
                    Break
                }

                # Day, month, year and hour
                { $_ -in (
                        'd/M HH',
                        'dd/M HH',
                        'd/MM HH',
                        'dd/MM HH',
                        'd/M/yyyy HH',
                        'dd/M/yyyy HH',
                        'd/MM/yyyy HH',
                        'dd/MM/yyyy HH',
                        'd/M/yyyy H',
                        'dd/M/yyyy H',
                        'd/MM/yyyy H',
                        'dd/MM/yyyy H'
                    )
                }
                {
                    $DateTime = $DateTime.AddHours(1).AddSeconds(-1)
                    Break
                }

                # Day, month, year, hour and minute
                { $_ -in (
                        'd/M H:m',
                        'dd/M H:m',
                        'dd/MM H:m',
                        'd/MM H:m',
                        'd/M HH:m',
                        'dd/M HH:m',
                        'dd/MM HH:m',
                        'd/MM HH:m',
                        'd/M H:mm',
                        'dd/M H:mm',
                        'dd/MM H:mm',
                        'd/MM H:mm',
                        'd/M HH:mm',
                        'dd/M HH:mm',
                        'dd/MM HH:mm',
                        'd/MM HH:mm',
                        'd/M/yyyy H:m',
                        'dd/M/yyyy H:m',
                        'd/MM/yyyy H:m',
                        'dd/MM/yyyy H:m',
                        'd/M/yyyy HH:m',
                        'dd/M/yyyy HH:m',
                        'd/MM/yyyy HH:m',
                        'dd/MM/yyyy HH:m',
                        'd/M/yyyy H:mm',
                        'dd/M/yyyy H:mm',
                        'd/MM/yyyy H:mm',
                        'dd/MM/yyyy H:mm',
                        'd/M/yyyy HH:mm',
                        'dd/M/yyyy HH:mm',
                        'd/MM/yyyy HH:mm',
                        'dd/MM/yyyy HH:mm'
                    )
                }
                {
                    $DateTime = $DateTime.AddMinutes(1).AddSeconds(-1)
                    Break
                }

                # Day, month, year, hour, minute and second
                { $_ -in (
                        'd/M/yyyy HH:m:s',
                        'dd/M/yyyy HH:m:s',
                        'd/MM/yyyy HH:m:s',
                        'dd/MM/yyyy HH:m:s',
                        'd/M/yyyy H:m:s',
                        'dd/M/yyyy H:m:s',
                        'd/MM/yyyy H:m:s',
                        'dd/MM/yyyy H:m:s',
                        'd/M HH:m:s',
                        'dd/M HH:m:s',
                        'd/MM HH:m:s',
                        'dd/MM HH:m:s',
                        'd/M HH:mm:s',
                        'dd/M HH:mm:s',
                        'd/MM HH:mm:s',
                        'dd/MM HH:mm:s',
                        'd/M H:mm:s',
                        'dd/M H:mm:s',
                        'd/MM H:mm:s',
                        'dd/MM H:mm:s',
                        'dd/MM HH:m:ss',
                        'dd/MM HH:mm:ss',
                        'dd/MM/yyyy HH:mm:ss'
                    )
                }
                {
                    $DateTime = Get-Date $DateTime -Second 59
                    Break
                }

                Default
                {
                    Throw "Unable to convert unsupported date format for date string '$DateTimeString'"
                }
            }
        }

        Return $DateTime

    }
    Catch
    {
        Throw
    }
}

# Function that authenticates to Azure and retrieves an access token for the Flow API
Function Connect-AzAccountAndGetAccessToken
{
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                If ([guid]::TryParse($_, $([ref][guid]::Empty)))
                {
                    Return $true
                }
                Else
                {
                    Throw "`nInvalid Tenant ID: $_"
                }
            })]
        [String]$TenantId
    )

    # Scriptblock to authenticate to Azure and retrieve an access token for the Flow API
    $AzAuthenticationScriptBlock = {
        Param(
            [String]$TenantId
        )

        Try
        {
            #If ((Get-AzContext).Tenant.ID -ne $TenantId) {
            Connect-AzAccount -TenantId $TenantId -Scope CurrentUser -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

            #! test login timeout

            #}
            (Get-AzAccessToken -ResourceUrl 'https://service.flow.microsoft.com/' -TenantId $TenantId -ErrorAction Stop)
        }
        Catch
        {
            Throw
        }
    }

    Try
    {
        # Start the job
        $Job = Start-Job -ScriptBlock $AzAuthenticationScriptBlock -ArgumentList $TenantId

        # Wait for the job to complete
        Wait-Job -Job $Job | Out-Null

        # Retrieve the output
        $AccessToken = Receive-Job -Job $Job

        # Clean up the job
        $InvolvedJobs = Get-Job | Where-Object -FilterScript { $_.Command -eq $AzAuthenticationScriptBlock }
        $InvolvedJobs | Remove-Job -Force

        Return $AccessToken
    }
    Catch
    {
        Throw
    }
}

# Function that returns a Hyperlink clickable in the PowerShell console
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

# Function that shows a processing animation while executing a script block
Function Show-ProcessingAnimation
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, HelpMessage = 'The script block to execute.')]
        [ScriptBlock]
        $ScriptBlock,

        [HashTable]
        $ScriptBlockParams,

        [Parameter(HelpMessage = 'Where to display the animation.')]
        [ValidateSet('Console', 'WindowTitle', 'Both')]
        [string]
        $AnimationPosition = 'Console'
    )

    # Store the original console title to restore it later
    $originalTitle = $Host.UI.RawUI.WindowTitle

    Try
    {
        [Console]::CursorVisible = $false

        $Counter = 0
        $Frames = '|', '/', '-', '\'
        $Job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ScriptBlockParams

        While ($Job.State -eq 'Running')
        {
            $Frame = $Frames[$Counter % $Frames.Length]

            if ($AnimationPosition -eq 'Console' -or $AnimationPosition -eq 'Both')
            {
                Try
                {
                    # Dynamically update cursor position based on current console size
                    $CursorTop = [Math]::Min([Console]::CursorTop, [Console]::BufferHeight - 1)

                    [Console]::SetCursorPosition(0, $CursorTop)
                    Write-Host "$Frame" -NoNewline
                }
                Catch
                {
                    # Even if we fail to set the cursor position, continue the animation
                    Write-Host "$Frame" -NoNewline
                }
            }

            if ($AnimationPosition -eq 'WindowTitle' -or $AnimationPosition -eq 'Both')
            {
                # Append the frame to the current window title
                $Host.UI.RawUI.WindowTitle = "$originalTitle - Processing  $Frame"
            }

            $Counter += 1
            Start-Sleep -Milliseconds 125
        }

        # Receive job results or errors
        $Result = Receive-Job -Job $Job
        Remove-Job -Job $Job

        # Only needed if you use multiline frames in console mode
        if ($AnimationPosition -eq 'Console' -or $AnimationPosition -eq 'Both')
        {
            Write-Host "`r" + ($Frames[0] -replace '[^\s+]', ' ') -NoNewline
        }
    }
    Catch
    {
        Write-Host "An error occurred: $_"
    }
    Finally
    {
        Try
        {
            if ($AnimationPosition -eq 'Console' -or $AnimationPosition -eq 'Both')
            {
                [Console]::SetCursorPosition(0, [Console]::CursorTop)
            }

            if ($AnimationPosition -eq 'WindowTitle' -or $AnimationPosition -eq 'Both')
            {
                # Restore the original console title
                $Host.UI.RawUI.WindowTitle = $originalTitle
            }
        }
        Catch
        {
            Write-Host "An error occurred while resetting: $_"
        }

        [Console]::CursorVisible = $true
    }

    Return $Result
}

#EndRegion Functions

Try
{

    # Check if Az module is installed (Not included in Require statement because of a bug)
    if ($null -eq (Get-Module -Name 'Az' -ListAvailable))
    {
        Throw 'Az module is not installed.'
    }

    # Declare variables
    $FlowRuns = @()
    $FlowRuns_API_Calls = @()
    $RunsRetrievalExecutionTime = $null

    # Prompt user for a date range (if any) to filter the Flow Runs to retrieve
    $StartDateFilterInput = Read-Host -Prompt 'Start date filter'
    $EndDateFilterInput = Read-Host -Prompt 'End date filter'

    # If no date range is provided, set a fixed date
    If (-not $StartDateFilterInput)
    {
        $StartDateFilter = (Get-Date).AddMonths(-3)
    }
    Else
    {
        $StartDateFilter = Get-DateTimeFromString -DateTimeString $StartDateFilterInput
    }

    # If no end date is provided, set the end date to today
    If (-not $EndDateFilterInput)
    {
        $EndDateFilter = $(Get-Date)
    }
    Else
    {
        $EndDateFilter = Get-DateTimeFromString -DateTimeString $EndDateFilterInput -ClosestEnd
    }

    # Check if the date range is valid
    If ($StartDateFilter -gt $EndDateFilter)
    {
        Throw "`nInvalid date range: Start date must be before end date."
    }

    # Declare variables
    $ScriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $ScriptStartTime = Get-Date
    $ScriptResult = 'INTERRUPTED'

    # Authenticate to Azure AD, get the access token, then connect to SharePoint Online (required to get the flow error details)
    $AccessToken = Connect-AzAccountAndGetAccessToken -TenantId $TenantId -ResourceUrl 'https://service.flow.microsoft.com/'

    # Set the headers for the Power Automate API calls
    $PA_API_Call_Headers = @{
        'Authorization' = "Bearer $($AccessToken.Token)"
        'Content-Type'  = 'application/json'
    }

    # Get available Power Automate environments
    Write-Host "`n[$(Get-Date)] Getting available Power Automate environments..." -ForegroundColor Cyan
    $PA_Environments_API_Call = 'https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01'
    $PA_AvailableEnvironments = Invoke-RestMethod -Method GET -Uri $PA_Environments_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 180 -RetryIntervalSec 15 -MaximumRetryCount 3
    $PA_Environments = $PA_AvailableEnvironments.value | Select-Object @{L = 'DisplayName'; E = { $_.properties.displayName } }, @{L = 'ID'; E = { $_.name } }, @{L = 'Default'; E = { $_.properties.isDefault } } | Sort-Object -Property Default -Descending
    If ($PA_Environments.Count -eq 0)
    {
        Throw "`nNo Power Automate environment found."
    }


    # Filter the available environments to the one(s) specified by the user
    If ($EnvironmentIDs)
    {
        # Check if the provided Environment IDs are valid
        ForEach ($EnvironmentID in $EnvironmentIDs)
        {
            If ($EnvironmentID -notin ($PA_Environments.ID))
            {
                Throw "`Environment ID not found: $EnvironmentID"
            }
        }
        $PA_Environments = $PA_Environments | Where-Object -FilterScript { $_.ID -in $EnvironmentIDs }
    }
    If ($PA_Environments.Count -eq 0)
    {
        Throw "`nNo valid environment found with provided Environment IDs:`n$($EnvironmentIDs -join "`n")"
    }

    # Loop through available environments to search for the Flow
    ForEach ($PA_Environment in $PA_Environments)
    {
        # Print details about the environment to the console
        Write-Host ("[{0}] Searching for Flow in environment '{1}'..." -f
            $(Get-Date),
            $($PA_Environment.DisplayName)
        ) -ForegroundColor Cyan

        # Get the Flow object
        $PA_Flow_API_Call = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($FlowId)?api-version=2016-11-01"
        $FlowObject = Invoke-RestMethod -Method GET -Uri $PA_Flow_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 180 -RetryIntervalSec 15 -MaximumRetryCount 3 -SkipHttpErrorCheck

        # If the Flow is found, break the loop
        If ($FlowObject.error.code -ne 'FlowNotFound')
        {
            Write-Host ("{0}[{1}] Flow '{2}' found in environment '{3}'." -f
                "`n",
                $(Get-Date),
                $FlowObject.properties.displayName,
                $($PA_Environment.DisplayName)
            ) -ForegroundColor Green
            Break
        }
    }

    # If the Flow is not found in any of the available environments, throw an error
    If ($FlowObject.error.code -eq 'FlowNotFound')
    {
        If ($EnvironmentIDs)
        {
            $SearchMode = 'filtered'
        }
        Else
        {
            $SearchMode = 'available'
        }
        Throw "`nFlow '$FlowId' not found in any of the $SearchMode environments."
    }

    # If the Flow is found in multiple environments, throw an error (should not happen)
    If ($FlowObject.Count -gt 1)
    {
        Throw "`nFlow '$FlowId' found in multiple environments. Please specify the environment to use (TO BE IMPLEMENTED)."
    }

    # Print details about the Flow to the console
    Write-Host ("{0}[{1}] Getting runs for:{0}Flow ID: {2}{0}Flow Display Name: {3}{0}Date range: from '{4}' to '{5}'{0}Status filter: {6}{0}" -f
        "`n",
        $(Get-Date),
        $($FlowId),
        $($FlowObject.properties.displayName),
        $($StartDateFilter.ToString('dd/MM/yyyy HH:mm:ss')),
        $($EndDateFilter.ToString('dd/MM/yyyy HH:mm:ss')),
        $((-not $StatusFilter -or $StatusFilter -eq '*') ? 'All' : ($StatusFilter -Join ', '))
    ) -ForegroundColor Cyan

    # Set the properties to be retrieved from the Flow Runs
    $PropertiesSelection = @(
        @{L = 'FlowRunName'; E = { $_.name } },
        @{L = 'Status'; E = { $_.properties.status } }
        @{L = 'LocalStartTime'; E = { Get-Date $_.properties.startTime.ToLocalTime() } },
        @{L = 'LocalEndTime'; E = { Get-Date $_.properties.endTime.ToLocalTime() } },
        @{L = 'StartTime'; E = { Get-Date $_.properties.startTime } },
        @{L = 'EndTime'; E = { Get-Date $_.properties.endTime } },
        @{L = 'TriggerOutputLink'; E = { $_.properties.trigger.outputsLink.uri } }
    )

    # Build the API call to get the Flow Runs
    $EncodedStartDateFilter = $StartDateFilter.ToUniversalTime().ToString('o')
    $EncodedEndDateFilter = $EndDateFilter.ToUniversalTime().ToString('o')
    If (
        $StatusFilter -notin ('*', 'All') -and
        $StatusFilter.Count -gt 0
    )
    {
        ForEach ($Status in $StatusFilter)
        {
            $EncodedUriFilter = [System.Web.HttpUtility]::UrlEncode("startTime ge $EncodedStartDateFilter and startTime le $($EncodedEndDateFilter) and status eq '$Status'")
            $FlowRuns_API_Calls += [Ordered]@{
                StatusFilter = $Status
                URI          = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($FlowId)/runs?api-version=2016-11-01&`$filter=$($EncodedUriFilter)"
            }
        }
    }
    Else
    {
        $EncodedUriFilter = [System.Web.HttpUtility]::UrlEncode("startTime ge $EncodedStartDateFilter and startTime le $($EncodedEndDateFilter)")
        $FlowRuns_API_Calls += [Ordered]@{
            StatusFilter = 'All'
            URI          = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($FlowId)/runs?api-version=2016-11-01&`$filter=$($EncodedUriFilter)"
        }
    }

    # Get Flow Runs
    $RunsRetrievalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    ForEach ($FlowRuns_API_Call in $FlowRuns_API_Calls)
    {
        # Scriptblock parameters
        $ScriptBlockParams = @{
            FlowRuns_API_Call   = $FlowRuns_API_Call.URI
            PA_API_Call_Headers = $PA_API_Call_Headers
        }

        $FlowRunsRetrievalScriptblock = Show-ProcessingAnimation -AnimationPosition Both -ScriptBlock {
            Param ($Params)
            # Use $Params['ParamName'] to reference the parameters
            $FlowRuns = @()
            $FlowRuns_API_Call = $Params['FlowRuns_API_Call']
            $PA_API_Call_Headers = $Params['PA_API_Call_Headers']
            Do
            {
                # Make the API call
                $FlowRunsRetrieval = Invoke-RestMethod -Method GET -Uri $FlowRuns_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 180 -RetryIntervalSec 15 -MaximumRetryCount 3

                # Add the returned items to the allRuns array
                $FlowRuns += $FlowRunsRetrieval.value

                # Check for a nextLink
                $FlowRuns_API_Call = $FlowRunsRetrieval.nextLink

            } While ($FlowRuns_API_Call)

            # Name and value of the variable to return as property of the returned object
            Return [Ordered]@{FlowRuns = $FlowRuns }

        } -ScriptBlockParams $ScriptBlockParams
        $FlowRuns += $FlowRunsRetrievalScriptblock.FlowRuns | Select-Object $PropertiesSelection
        Write-Host "[$(Get-Date)] Retrieved $($FlowRuns.Count) runs with status: $($FlowRuns_API_Call.StatusFilter)" -ForegroundColor Green
    }
    $FlowRuns = $FlowRuns | Sort-Object -Property LocalStartTime

    # Print details about the Flow Runs retrieval to the console
    $RunsRetrievalStopwatch.Stop()
    $RunsRetrievalExecutionTime = $(Get-Date -Date $($RunsRetrievalStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
    Write-Host "[$(Get-Date)] Runs' retrieval completed in: $RunsRetrievalExecutionTime" -ForegroundColor Green
    If ($FlowRuns.Count -lt 1)
    {
        Write-Host "[$(Get-Date)] No runs found with specified filters." -ForegroundColor Green
        $ScriptResult = 'SUCCESS'
        Return
    }

    # Get FlowRun trigger and error details (if any) from filtered Flow Runs
    $Counter = 0
    $RunsDetailsRetrievalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "`n[$(Get-Date)] Starting retrieval of $($FlowRuns.Count) runs' trigger details..." -ForegroundColor Cyan
    ForEach ($FlowRun in $FlowRuns)
    {
        $Counter++
        $PercentComplete = [Math]::Round(($Counter / $FlowRuns.Count * 100))
        $ProgressBarParameters = @{
            Activity        = 'FlowRun trigger details and error retrieval...'
            Status          = "FlowRun: $($Counter)/$($FlowRuns.Count) ($($PercentComplete)%)"
            PercentComplete = $PercentComplete
        }
        Write-Progress @ProgressBarParameters

        # Provide details about the FlowRun in the console
        $FlowRunLink = "https://make.powerautomate.com/environments/$($PA_Environment.ID)/flows/$($FlowId)/runs/$($FlowRun.FlowRunName)"
        $ConsoleFlowRunLink = New-PowerShellHyperlink -LinkURL $FlowRunLink -LinkDisplayText $FlowRun.FlowRunName
        Write-Host ('{0}[{1}/{2}]{0}ID: {3}{0}Start time: {4}{0}Result: {5}' -f
            "`n",
            $Counter,
            $FlowRuns.Count,
            $ConsoleFlowRunLink,
            $FlowRun.LocalStartTime,
            $FlowRun.Status
        ) -ForegroundColor Cyan

        # Get FlowRun trigger details, if any
        if ($FlowRun.TriggerOutputLink)
        {
            $DetailedFlowRun = Invoke-RestMethod -Method GET -Uri $($FlowRun.TriggerOutputLink) -TimeoutSec 180 -RetryIntervalSec 15 -MaximumRetryCount 3
            $DetailedFlowRunProperties = ($DetailedFlowRun.body | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue).Name ?? $null

            # Add FlowRun trigger details to the FlowRun object
            ForEach ($Property in $DetailedFlowRunProperties)
            {
                $FlowRun | Add-Member -MemberType NoteProperty -Name "Trigger Property: $($Property)" -Value $(($DetailedFlowRun.body.$Property | Format-List | Out-String).Trim())
            }
            $FlowRun | Add-Member -MemberType NoteProperty -Name 'FlowRunLink' -Value $FlowRunLink -Force
        }

        # Get FlowRun error details (if any)
        If ($FlowRun.Status -notin ('Running', 'Succeeded'))
        {
            $API_URI_FlowRunError = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($FlowId)/runs/$($FlowRun.FlowRunName)/remediation?api-version=2016-11-01" #?`$expand=properties%2Factions,properties%2Fflow&api-version=2016-11-01&include=repetitionCount"
            Do
            {
                Try
                {
                    $CustomError = $null
                    $API_Call_Error = $null
                    $DetailedFlowRunError = Invoke-RestMethod -Method GET -Uri $API_URI_FlowRunError -TimeoutSec 180 -RetryIntervalSec 15 -MaximumRetryCount 3 -Headers $PA_API_Call_Headers
                }
                Catch
                {
                    $API_Call_Error = ($_.ErrorDetails.Message | ConvertFrom-Json).error.code
                    If ($API_Call_Error -eq 'ExpiredAuthenticationToken')
                    {
                        Write-Host "`n[$(Get-Date)] Access token expired. Getting a new one..." -ForegroundColor Yellow
                        $AccessToken = Connect-AzAccountAndGetAccessToken -TenantId $TenantId
                        $PA_API_Call_Headers.Authorization = 'Bearer ' + $AccessToken.Token
                    }
                    ElseIf ($API_Call_Error -eq 'RemediationNotFound')
                    {
                        $API_Call_Error = $null
                        $CustomError = 'Missing error details'
                    }
                    Else
                    {
                        Throw
                    }
                }
            } While ($API_Call_Error)

            # Add FlowRun error details to the FlowRun object
            $ExtractedFlowRunError = $CustomError ?? ($DetailedFlowRunError | Select-Object `
                @{L = 'StatusCode'; E = { $_.operationOutputs.StatusCode ?? 'N/A' } },
                @{L = 'ErrorType'; E = { $_.remediationType } ?? 'N/A' },
                @{L = 'ErrorAction'; E = { $_.errorSubject } ?? 'N/A' },
                @{L = 'ErrorMessage'; E = { $($_.SearchText?.trim('()')) ?? 'N/A' } } |
                    Format-List | Out-String).Trim()
            $FlowRun | Add-Member -MemberType NoteProperty -Name 'FlowError' -Value $ExtractedFlowRunError -Force
        }
        Else
        {
            $FlowRun | Add-Member -MemberType NoteProperty -Name 'FlowError' -Value 'N/A' -Force
        }
    }

    # Print details about the FlowRun trigger details retrieval to the console
    $RunsDetailsRetrievalStopwatch.Stop()
    $RunsDetailsRetrievalExecutionTime = $(Get-Date -Date $($RunsDetailsRetrievalStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
    Write-Host "`n[$(Get-Date)] Runs' details' retrieval completed in: $RunsDetailsRetrievalExecutionTime " -ForegroundColor Green
    $ScriptResult = 'SUCCESS'
}
Catch
{
    $ScriptResult = 'ERROR'
    Throw
}
Finally
{
    Write-Progress -Activity 'FlowRun trigger details retrieval...' -Completed

    # Provide details about the script execution in the console
    Switch ($ScriptResult)
    {
        'SUCCESS'
        {
            $ForegroundColor = 'Green'
            Break
        }
        'INTERRUPTED'
        {
            $ForegroundColor = 'Yellow'
            Break
        }
        'ERROR'
        {
            $ForegroundColor = 'Red'
            Break
        }
        Default
        {
            $ForegroundColor = 'White'
        }
    }
    Write-Host "`n[$(Get-Date)] Script result: $ScriptResult" -ForegroundColor $ForegroundColor
    If ($RunsRetrievalExecutionTime)
    {
        Write-Host "[$(Get-Date)] Initial runs' retrieval completed in: $RunsRetrievalExecutionTime" -ForegroundColor $ForegroundColor
    }
    # Export results (if any) to CSV
    $RetrievedDetailedFlowRuns = ($FlowRuns | Where-Object -FilterScript { $_.FlowError }).Count
    If ($FlowRuns)
    {
        Write-Host "`n[$(Get-Date)] Exporting $RetrievedDetailedFlowRuns retrieved results to CSV..." -ForegroundColor $ForegroundColor
        $CsvPath = "$($PSScriptRoot)\Logs\$($FlowObject.properties.displayName) - Runs - $($ScriptStartTime.ToString('dd_MM_yyyy-HH_mm_ss')).csv"
        If (!(Test-Path -Path $CsvPath))
        {
            New-Item -Path $CsvPath -ItemType File -Force | Out-Null
        }
        $FlowRuns | Select-Object -ExcludeProperty TriggerOutputLink, StartTime, EndTime | Export-Csv -Path $CsvPath -Delimiter ';' -NoTypeInformation
    }

    # Print details about the script execution to the console
    $ScriptStopwatch.Stop()
    $ScriptExecutionTime = $(Get-Date -Date $($ScriptStopwatch.Elapsed.ToString()) -Format 'HH:mm:ss')
    Write-Host "Script total execution time: $ScriptExecutionTime" -ForegroundColor $ForegroundColor
    #Disconnect-AzAccount | Out-Null
}