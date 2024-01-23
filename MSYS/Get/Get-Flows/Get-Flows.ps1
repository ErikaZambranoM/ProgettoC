#Requires -Version 7

<#
    .SYNOPSIS
    Get Power Automate flows from a Power Automate environment.

    .DESCRIPTION
    Get Power Automate flows from a Power Automate environment.

    .REQUIREMENTS
        - Az module
        - $TenantId variable set to Azure the tenant ID

    .PARAMETER EnvironmentID
    Power Automate environment ID. If not provided, the default environment will be used.

    .EXAMPLE
    Get-Flows

    .EXAMPLE
    Get-Flows -EnvironmentID "11111111-aaaa-2222-bbbb-333333333333"
#>

Param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({
            If (
                !([GUID]::TryParse($_.Replace('Default-', ''), $([ref][guid]::Empty))) -and
                !([GUID]::TryParse($_, $([ref][guid]::Empty)))
            )
            {
                Throw "`nInvalid Environment ID: $_"
            }
            Return $true

        })]
    [String]
    $EnvironmentID
)

# Tenant ID required to authenticate to Azure AD (required to get the flow error details)
$TenantId = '7cc91888-5aa0-49e5-a836-22cda2eae0fc'

#Region Functions

# Function that authenticates to Azure and retrieves an access token for the Flow API
Function Connect-AzAccountAndGetAccessToken
{
    Param(
        [String]$ResourceUrl,

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
            [String]$ResourceUrl,
            [String]$TenantId
        )

        Try
        {
            #If ((Get-AzContext).Tenant.ID -ne $TenantId) {
            #Disconnect-AzAccount | Out-Null
            Connect-AzAccount -TenantId $TenantId -Scope CurrentUser -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

            #! test login timeout
            #}
            (Get-AzAccessToken -ResourceUrl $ResourceUrl -TenantId $TenantId -ErrorAction Stop)
        }
        Catch
        {
            Throw
        }
    }

    Try
    {
        # Start the job
        $Job = Start-Job -ScriptBlock $AzAuthenticationScriptBlock -ArgumentList $ResourceUrl, $TenantId

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


# Function that shows a processing animation while executing a script block
Function Show-ProcessingAnimation
{
    Param (
        [Parameter(Mandatory = $true, HelpMessage = 'The script block to execute.')]
        [ScriptBlock]
        $ScriptBlock,

        [HashTable]
        $ScriptBlockParams
    )

    $CursorTop = [Console]::CursorTop

    Try
    {
        [Console]::CursorVisible = $false

        $Counter = 0
        $Frames = '|', '/', '-', '\'
        $Job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ScriptBlockParams

        While ($Job.State -eq 'Running')
        {
            $Frame = $Frames[$Counter % $Frames.Length]

            Write-Host "$Frame" -NoNewline
            [Console]::SetCursorPosition(0, $CursorTop)

            $Counter += 1
            Start-Sleep -Milliseconds 125
        }

        # Receive job results or errors
        $Result = Receive-Job -Job $Job
        Remove-Job -Job $Job

        # Only needed if you use multiline frames
        Write-Host ($Frames[0] -replace '[^\s+]', ' ') -NoNewline
    }
    Finally
    {
        [Console]::SetCursorPosition(0, $CursorTop)
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

    # Get the access token for the Power Automate API
    $PA_AccessToken = (Connect-AzAccountAndGetAccessToken -ResourceUrl 'https://service.flow.microsoft.com/' -TenantId $TenantId).Token

    # Set the headers for the Power Automate API calls
    $PA_API_Call_Headers = @{
        'Authorization' = "Bearer $($PA_AccessToken)"
        'Content-Type'  = 'application/json'
    }

    # Get available Power Automate environments
    Write-Host "`n[$(Get-Date)] Getting available Power Automate environments..." -ForegroundColor Cyan
    $PA_Environments_API_Call = 'https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01'
    $PA_AvailableEnvironments = Invoke-RestMethod -Method GET -Uri $PA_Environments_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 120 -RetryIntervalSec 5 -MaximumRetryCount 2
    $PA_Environments = $PA_AvailableEnvironments.value |
        Select-Object @{L = 'DisplayName'; E = { $_.properties.displayName } },
        @{L = 'ID'; E = { $_.name } },
        @{L = 'Default'; E = { $_.properties.isDefault } },
        @{L = 'SolutionsEnvironmentUrl'; E = { $_.properties.linkedEnvironmentMetadata.instanceApiUrl } } |
            Sort-Object -Property Default -Descending

    If ($PA_Environments.Count -eq 0)
    {
        Throw "`nNo Power Automate environment found."
    }

    # Filter the available environments to the one(s) specified by the user
    If ($EnvironmentID)
    {
        # Check if the provided Environment IDs are valid
        If ($EnvironmentID -notin ($PA_Environments.ID))
        {
            Throw "`Environment ID not found: $EnvironmentID"
        }
        $PA_Environment = $PA_Environments | Where-Object -FilterScript { $_.ID -in $EnvironmentID }
    }
    Else
    {
        # If no Environment ID is provided, use the default environment
        $PA_Environment = $PA_Environments | Where-Object -FilterScript { $_.Default -eq $true }
    }
    If (-not $PA_Environment)
    {
        Throw "`nNo valid environment found with provided Environment ID:`n'$($EnvironmentID)'"
    }



    # Get the access token for Dynamics 365 API (where the solution components are stored)
    $D365_AccessToken = (Connect-AzAccountAndGetAccessToken -ResourceUrl $($PA_Environment.SolutionsEnvironmentUrl) -TenantId $TenantId).Token

    # Set the headers for the Dynamics 365 API calls
    $D365_API_Call_Headers = @{
        'Authorization' = "Bearer $($D365_AccessToken)"
        'Content-Type'  = 'application/json'
    }

    # Scriptblock parameters
    $ScriptBlockParams = @{
        PA_Flows_API_Call             = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows?`$filter=*&`$expand=properties%2Ftriggers&api-version=2016-11-01"
        PA_API_Call_Headers           = $PA_API_Call_Headers
        #PA_SolutionsFlows_API_Call    = "$($PA_Environment.SolutionsEnvironmentUrl)/api/data/v9.2/msdyn_solutioncomponentsummaries?api-version=9.1&`$filter=msdyn_componentlogicalname%20eq%20'workflow'"
        PA_SolutionsFlows_API_Call    = "$($PA_Environment.SolutionsEnvironmentUrl)/api/data/v9.2/workflows?`$filter=category%20eq%205&api-version=9.2"
        PA_Solutions_API_Call_Headers = $D365_API_Call_Headers
    }

    # Get the Flows
    Write-Host "`n[$(Get-Date)] Getting all flows..." -ForegroundColor Cyan
    $FlowRunsRetrievalScriptblock = Show-ProcessingAnimation -ScriptBlock {
        Param ($Params)
        # Use $Params['ParamName'] to reference the parameters
        $PA_Flows_API_Call = $Params['PA_Flows_API_Call']
        $PA_API_Call_Headers = $Params['PA_API_Call_Headers']
        $PA_SolutionsFlows_API_Call = $Params['PA_SolutionsFlows_API_Call']
        $PA_Solutions_API_Call_Headers = $Params['PA_Solutions_API_Call_Headers']
        $RawFlowList = Invoke-RestMethod -Method GET -Uri $PA_Flows_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 120 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck
        $RawSolutionsFlowList = Invoke-RestMethod -Method GET -Uri $PA_SolutionsFlows_API_Call -Headers $PA_Solutions_API_Call_Headers -TimeoutSec 120 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck

        # Name and value of the variable to return as property of the returned object
        Return [Ordered]@{
            RawFlowList          = $RawFlowList
            RawSolutionsFlowList = $RawSolutionsFlowList
        }

    } -ScriptBlockParams $ScriptBlockParams

    # Set Flows' properties to be retrieved from the environment
    $FlowsPropertiesSelection = @(
        @{L = 'FlowId'; E = { $_.name } },
        @{L = 'DisplayName'; E = { $_.properties.displayName } },
        @{L = 'FlowUrl'; E = { "https://make.powerautomate.com/environments/$($PA_Environment.ID)/flows/$($_.name)/details" } },
        @{L = 'Status'; E = { $_.properties.state } },
        @{L = 'CreatedTime'; E = { $_.properties.createdTime } },
        @{L = 'ModifiedTime'; E = { $_.properties.lastModifiedTime } },
        @{L = 'SolutionId'; E = { 'N/A' } },
        @{L = 'TriggerConnectorType'; E = {
                $_.properties.definitionSummary.triggers.api.properties.displayName ?? # Automatic
                $_.properties.definitionSummary.triggers.kind ?? # Manual
                $_.properties.definitionSummary.triggers.type # Recurrence
            }
        },
        @{L = 'TriggerAction'; E = {
                $_.properties.definitionSummary.triggers.swaggerOperationId ?? # Automatic and Manual
                $_.properties.definitionSummary.triggers.type # Recurrence
            }
        },
        @{L = 'TriggerUri'; E = { $null } },
        @{L = 'Description'; E = { $_.properties.definitionSummary.description } }
    )
    [Array]$FlowList = $FlowRunsRetrievalScriptblock.RawFlowList.value | Select-Object -Property $FlowsPropertiesSelection

    # Set Solutions' Flows' properties to be retrieved from the environment
    $SolutionsFlowsPropertiesSelection = @(
        @{L = 'FlowId'; E = { $_.workflowidunique } },
        @{L = 'DisplayName'; E = { $_.name } },
        @{L = 'FlowUrl'; E = { "https://make.powerautomate.com/environments/$($PA_Environment.ID)/flows/$($_.workflowidunique)/details" } }
        @{L = 'Status'; E = {
                switch ($_.statecode)
                {
                    0 { 'Stopped'; break }
                    1 { 'Started'; break }
                    Default { 'Unknown'; break }
                }
            }
        },
        @{L = 'CreatedTime'; E = { $_.createdon } },
        @{L = 'ModifiedTime'; E = { $_.modifiedon } },
        @{L = 'ModifiedBy'; E = { ( (Get-AzADUser -ObjectId $_._modifiedby_value -ErrorAction SilentlyContinue) ?? $_._modifiedby_value) } },
        @{L = 'SolutionId'; E = { $_.solutionid } },
        @{L = 'TriggerConnectorType'; E = {
                $(($_.clientdata | ConvertFrom-Json -Depth 100 -AsHashtable).properties.definition.triggers.Values.inputs.host.connectionName) ?? # Automatic
                $(($_.clientdata | ConvertFrom-Json -Depth 100 -AsHashtable).properties.definition.triggers.Values.kind) ?? # Manual
                $(($_.clientdata | ConvertFrom-Json -Depth 100 -AsHashtable).properties.definition.triggers.Values.type) # Recurrence
            }
        },
        @{L = 'TriggerAction'; E = {
                $(($_.clientdata | ConvertFrom-Json -Depth 100 -AsHashtable).properties.definition.triggers.Values.inputs.method) ?? # Manual
                $(($_.clientdata | ConvertFrom-Json -Depth 100 -AsHashtable).properties.definition.triggers.Values.inputs.host.operationId ) ?? # Automatic
                $(($_.clientdata | ConvertFrom-Json -Depth 100 -AsHashtable).properties.definition.triggers.Values.type ) # Recurrence
            }
        },
        @{L = 'TriggerUri'; E = { $null } },
        @{L = 'Description'; E = { $_.description } }
    )
    [Array]$SolutionsFlowList = $FlowRunsRetrievalScriptblock.RawSolutionsFlowList.value | Select-Object -Property $SolutionsFlowsPropertiesSelection

    # Merge the Flows and Solutions' Flows
    $AllFlowList = $FlowList + $SolutionsFlowList

    # Get the trigger URI for the flows that have an HTTP trigger
    Write-Host "`n[$(Get-Date)] Getting the trigger URI for the flows that have an HTTP trigger..." -ForegroundColor Cyan
    $FlowCounter = 0
    foreach ($Flow in $AllFlowList)
    {
        # Progress bar
        $FlowCounter++
        $Parameters = @{
            Activity        = 'Getting URI for flows with an HTTP trigger'
            Status          = "Flow: $($Flow.DisplayName)"
            PercentComplete = ($FlowCounter / $AllFlowList.Count * 100)
        }
        Write-Progress @Parameters

        if ($Flow.TriggerConnectorType -eq 'Http')
        {
            $TriggerUri = (Invoke-RestMethod -Uri "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($Flow.FlowId)/triggers/manual/listCallbackUrl?&api-version=2016-11-01" -Headers $PA_API_Call_Headers -TimeoutSec 120 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck -Method Post).Response.Value
            $Flow.TriggerUri = $TriggerUri
        }
        elseif ($Flow.TriggerConnectorType -eq 'Button')
        {
            $TriggerUri = (Invoke-RestMethod -Uri "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($Flow.FlowId)?api-version=2016-11-01" -Headers $PA_API_Call_Headers -TimeoutSec 120 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck -Method Get).properties.flowTriggerUri
            $Flow.TriggerUri = $TriggerUri
        }
        else
        {
            $Flow.TriggerUri = 'N/A'
        }
    }
}
Catch
{
    Throw
}
Finally
{
    Write-Progress -Activity 'Getting the trigger URI for the flows that have an HTTP trigger' -Completed

    If ($AllFlowList)
    {
        Write-Host "`n[$(Get-Date)] Exporting retrieved flows to CSV..." -ForegroundColor Green
        $CsvPath = "$($PSScriptRoot)\Logs\$((Get-Date).ToString('dd_MM_yyyy-HH_mm_ss'))-$($PA_Environment.DisplayName).csv"
        If (!(Test-Path -Path $CsvPath))
        {
            New-Item -Path $CsvPath -ItemType File -Force | Out-Null
        }
        $AllFlowList | Select-Object -ExcludeProperty TriggerOutputLink, StartTime, EndTime | Export-Csv -Path $CsvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8

        # Manually add the BOM
        $CSV_Content = Get-Content -Path $CsvPath -Raw
        $CSV_Content = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetPreamble()) + $CSV_Content
        Set-Content -Path $CsvPath -Value $CSV_Content -Encoding UTF8
        Write-Host "`n[$(Get-Date)] CSV file exported to:`n$CsvPath" -ForegroundColor Green
    }
    Else
    {
        Write-Host "`n[$(Get-Date)] No flows retrieved." -ForegroundColor Red
    }
    Disconnect-AzAccount | Out-Null
}