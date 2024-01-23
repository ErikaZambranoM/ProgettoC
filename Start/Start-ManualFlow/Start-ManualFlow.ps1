param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$FlowId,

    [Parameter(Mandatory = $false)]
    [string]$EnvironmentID
)

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
            If ((Get-AzContext).Tenant.ID -ne $TenantId)
            {
                Connect-AzAccount -TenantId $TenantId -Scope CurrentUser -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

                #! test login timeout
            }
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

# Get the access token
$TenantId = '7cc91888-5aa0-49e5-a836-22cda2eae0fc'
$AccessToken = Connect-AzAccountAndGetAccessToken -TenantId $TenantId -ResourceUrl 'https://service.flow.microsoft.com/'

# Set the headers for the Power Automate API calls
$PA_API_Call_Headers = @{
    'Authorization' = "Bearer $($AccessToken.Token)"
    'Content-Type'  = 'application/json'
}

# Get available Power Automate environments
Write-Host "`n[$(Get-Date)] Getting available Power Automate environments..." -ForegroundColor Cyan
$PA_Environments_API_Call = 'https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01'
$Parameters = @{
    Method            = 'GET'
    Uri               = $PA_Environments_API_Call
    Headers           = $PA_API_Call_Headers
    TimeoutSec        = 120
    RetryIntervalSec  = 5
    MaximumRetryCount = 2
}
$PA_AvailableEnvironments = Invoke-RestMethod @Parameters
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

# Retrieve Flow properties (Trigger URL and input parameters)
$Parameters = @{
    Method            = 'GET'
    Uri               = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($FlowId)?`$expand=properties.connectionreferences.apidefinition,properties.definitionSummary.operations.apiOperation&api-version=2016-11-01"
    Headers           = $PA_API_Call_Headers
    TimeoutSec        = 120
    RetryIntervalSec  = 5
    MaximumRetryCount = 2
}
$FlowProperties = Invoke-RestMethod @Parameters -SkipHttpErrorCheck
If ($FlowProperties.error)
{
    Throw ($FlowProperties.error | Format-List | Out-String)
}
Write-Host "`n[$(Get-Date)] Flow found: $($FlowProperties.properties.displayName)" -ForegroundColor Cyan
If (-not $FlowProperties.properties.definition.triggers.manual)
{
    Throw "`nFlow is not manually triggerable."
}
$FlowTriggerUri = $FlowProperties.properties.flowTriggerUri
$FlowRequiredInputParams = $FlowProperties.properties.definition.triggers.manual.inputs.schema.required
$FlowInputParams = $FlowProperties.properties.definition.triggers.manual.inputs.schema.properties

# Foreach parameter, ask the user to provide a value and add it to the body. Specify the type of the parameter,if it is required and possible accepted values.
$Body = @{}
$FlowInputParams.psobject.Properties.name | ForEach-Object {
    Write-Host ''
    $ParamInternalName = $_
    $ParamDisplaylName = $FlowInputParams.$ParamInternalName.title
    $ParamType = $FlowInputParams.$ParamInternalName.type
    $ParamRequired = $FlowRequiredInputParams -contains $ParamInternalName
    $EnumValues = $FlowInputParams.$ParamInternalName.enum
    $ParamValue = Read-Host -Prompt ("Please provide a value for the parameter '{1}'{0}Type: {2}{0}Required: {3}{0}{4}{0}{1}" -f
        "`n",
        $ParamDisplaylName,
        $ParamType,
        $($ParamRequired.ToString().ToLower()),
        ($EnumValues -eq $null ? '' : "Accepted values: $($EnumValues -Join ', ')")
    )
    $Body.Add($ParamInternalName, $ParamValue)
}

#! add confirmation before triggering the flow with the provided parameters (json body)

# Trigger the Flow
$Parameters = @{
    Method            = 'POST'
    Uri               = $FlowTriggerUri
    Headers           = $PA_API_Call_Headers
    Body              = $Body | ConvertTo-Json -Depth 100
    TimeoutSec        = 120
    RetryIntervalSec  = 5
    MaximumRetryCount = 2
}
Invoke-RestMethod @Parameters
