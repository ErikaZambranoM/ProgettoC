#Requires -Version 7
#todo: check solutions in url
[CmdletBinding(DefaultParameterSetName = 'FlowUrlAndJSON')]
Param (

    [Parameter(Mandatory = $true, ParameterSetName = 'FlowUrlAndZip')]
    [Parameter(Mandatory = $true, ParameterSetName = 'FlowUrlAndJSON')]
    [Uri]
    $FlowUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'FlowDetailsAndZip')]
    [Parameter(Mandatory = $true, ParameterSetName = 'FlowDetailsAndJSON')]
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

    [Parameter(Mandatory = $false, ParameterSetName = 'FlowDetailsAndZip')]
    [Parameter(Mandatory = $false, ParameterSetName = 'FlowDetailsAndJSON')]
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

    [Parameter(Mandatory = $false, ParameterSetName = 'FlowUrlAndJSON')]
    [Parameter(Mandatory = $false, ParameterSetName = 'FlowDetailsAndJSON')]
    [switch]
    $AsJSON,

    [Parameter(Mandatory = $true, ParameterSetName = 'FlowUrlAndZip')]
    [Parameter(Mandatory = $true, ParameterSetName = 'FlowDetailsAndZip')]
    [switch]
    $AsZip
)

begin
{
    # Tenant ID required to authenticate to Azure AD (required to get the flow error details)
    $TenantId = '7cc91888-5aa0-49e5-a836-22cda2eae0fc'

    # If no output format is specified, ask the user before proceeding
    if (-not $PSBoundParameters.ContainsKey('AsJson') -and -not $PSBoundParameters.ContainsKey('AsZip'))
    {
        # Ask for output format
        $Title = 'No output format specified!'
        $Info = 'Choose the output format for the Flow definition:'

        $AsZipChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Zip', (
            'AsZip{0}Use the Business Application Platform API to export the Flow as a zip file in the default download folder.{0}{0}' -f
            "`n"
        )

        $AsJsonChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&JSON', (
            'AsJSON (Default){0}Save the Flow definition as a JSON string into the clipboard.{0}It also get stored in variable $FlowJsonDefinition to be accessed for further analysis within the same terminal session.{0}{0}' -f
            "`n"
        )

        $CancelChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Cancel', (
            'Cancel{0}Terminate the process without any change{0}{0}' -f
            "`n"
        )

        $Options = [System.Management.Automation.Host.ChoiceDescription[]] @($AsZipChoice, $AsJsonChoice, $CancelChoice)
        [int]$DefaultChoice = 1
        $ChoicePrompt = $Host.UI.PromptForChoice($Title, $Info, $Options, $DefaultChoice)

        Switch ($ChoicePrompt)
        {
            0
            {
                $AsZip = $true
                Write-Host 'Zip option chosen.' -ForegroundColor DarkCyan
            }

            1
            {
                $AsJSON = $true
                Write-Host 'JSON option chosen.' -ForegroundColor DarkCyan
            }
            2
            {
                Write-Host "`nProcess canceled!" -ForegroundColor Yellow
                Exit
            }
        }
    }

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
                #If ((Get-AzContext).Tenant.ID -ne $TenantId)
                #{
                #Connect-AzAccount -TenantId $TenantId -Scope CurrentUser -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

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

    # Function that gets the details of a Power Automate flow from its URL
    function Get-FlowDetailsFromUrl
    {
        <#
        .SYNOPSIS
        Extracts and validates the environment ID, solution ID (if present), and flow ID from a Power Automate flow URL.

        .DESCRIPTION
        This function takes a Power Automate flow URL, validates it, and extracts the environment ID, solution ID (if present), and flow ID.

        .PARAMETER Url
        The Power Automate flow URL from which the IDs need to be extracted and validated.

        .EXAMPLE
        PS> Get-FlowDetailsFromUrl -FlowUrl "https://make.powerautomate.com/environments/888880f-6484-4675-b4c8-e52c7a164797/flows/8888d530-853d-4b39-8614-11957c590775/runs"
    #>

        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$FlowUrl
        )

        Begin
        {
            # Regex pattern for a generically valid URL
            $UrlPattern = '^(https?):\/\/[^\s\/$.?#].[^\s]*$'

            # Regex pattern to match environment, optional solution, and flow IDs
            $IdPattern = 'environments\/([0-9a-fA-F\-]+)\/(?:solutions\/([0-9a-fA-F\-]+)\/)?flows\/([0-9a-fA-F\-]+)'
        }

        Process
        {
            if ($FlowUrl -notmatch $UrlPattern)
            {
                Throw 'Invalid URL format.'
                return
            }

            if ($FlowUrl -match $IdPattern)
            {
                $EnvironmentId = $Matches[1]
                $SolutionId = if ($Matches[2]) { $Matches[2] } else { 'N/A' }
                $FlowId = $Matches[3]

                if (-not ([GUID]::TryParse($EnvironmentId.Replace('Default-', ''), [ref][guid]::Empty)) -and
                    -not ([GUID]::TryParse($EnvironmentId, [ref][guid]::Empty)))
                {
                    Throw 'Invalid Environment ID.'
                    return
                }

                if ($SolutionId -ne 'N/A' -and -not ([GUID]::TryParse($SolutionId, [ref][guid]::Empty)))
                {
                    Throw 'Invalid Solution ID.'
                    return
                }

                if (-not ([GUID]::TryParse($FlowId, [ref][guid]::Empty)))
                {
                    Throw 'Invalid Flow ID.'
                    return
                }

                $Result = [PSCustomObject]@{
                    EnvironmentId = $EnvironmentId
                    SolutionId    = $SolutionId
                    FlowId        = $FlowId
                }
                return $Result
            }
            else
            {
                Throw 'URL does not contain valid environment, solution, and flow ID patterns.'
            }
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

    # Function that returns the Flow actions from a Flow JSON definition as an array of PSCustomObject
    function Get-FlowActions
    {
        param (
            [Parameter(Mandatory = $true)]
            [String]$FlowJsonDefinition
        )

        begin
        {
            function Find-AndAddActions
            {
                param (
                    [Object]$ParsedJsonObject,
                    [System.Collections.Generic.List[Object]]$AllActionsObject,
                    [String]$CurrentPath = 'Root'
                )

                if ($ParsedJsonObject -is [System.Management.Automation.PSCustomObject])
                {
                    foreach ($Property in $ParsedJsonObject.PSObject.Properties)
                    {
                        if ($Property.Name -eq 'actions')
                        {
                            foreach ($Action in $Property.Value.PSObject.Properties)
                            {
                                $ActionPath = if ($CurrentPath -eq 'Root') { "Root/$($Action.Name -replace '_', ' ')" } else { "$($CurrentPath -replace '_', ' ')/$($Action.Name -replace '_', ' ')" }
                                $Action.Value | Add-Member -MemberType NoteProperty -Name 'InternalName' -Value $Action.Name -Force
                                $Action.Value | Add-Member -MemberType NoteProperty -Name 'DisplayName' -Value $($Action.Name -replace '_', ' ') -Force
                                $Action.Value | Add-Member -MemberType NoteProperty -Name 'Path' -Value $CurrentPath -Force
                                if ('' -eq $Action.value.runAfter) { $Action.value.runAfter = 'N/A' }

                                $OrderedProperties = @('DisplayName', 'InternalName', 'Path', 'runAfter', 'type')
                                $OtherProperties = $OrderedProperties + ($Action.Value.PSObject.Properties.Name | Where-Object { $_ -notin $OrderedProperties -and $_ -ne 'metadata' })
                                $ActionProperties = $Action.value | Select-Object -Property $OtherProperties
                                $AllActionsObject.Add($ActionProperties)

                                if ($Action.Value.PSObject.Properties.Name -contains 'actions')
                                {
                                    Find-AndAddActions -ParsedJsonObject $Action.Value -AllActionsObject $AllActionsObject -CurrentPath $ActionPath
                                }
                            }
                        }
                        elseif ($Property.Value -is [System.Management.Automation.PSCustomObject])
                        {
                            $NewPath = if ($CurrentPath -eq 'Root') { "Root/$($Property.Name -replace '_', ' ')" } else { "$($CurrentPath -replace '_', ' ')/$($Property.Name -replace '_', ' ')" }
                            Find-AndAddActions -ParsedJsonObject $Property.Value -AllActionsObject $AllActionsObject -CurrentPath $NewPath
                        }
                    }
                }
            }
        }

        process
        {
            try
            {
                $ParsedJson = $FlowJsonDefinition | ConvertFrom-Json -Depth 100
                $FlowActionsList = New-Object System.Collections.Generic.List[Object]

                if ($ParsedJson -is [System.Collections.IEnumerable])
                {
                    $ParsedJson = $ParsedJson | Select-Object -First 1
                }

                Find-AndAddActions -ParsedJsonObject $ParsedJson -AllActionsObject $FlowActionsList

                return $FlowActionsList
            }
            catch
            {
                throw
            }
        }
    }

    # Function that searches for a specific word inside flow actions
    function Global:Search-InFlowActions
    {
        <#
            .SYNOPSIS
            Searches for a specific word within the properties of a list of flow actions.

            .DESCRIPTION
            The function takes a list of flow actions and a search word as input. It then recursively searches for the search word within the properties of each flow action. The function avoids duplicate results by keeping track of the actions that have already been processed.

            .PARAMETER FlowActions
            A list of flow actions to search in as returned from function Get-FlowActions. This parameter is mandatory.

            .PARAMETER SearchWord
            The word to search for within the properties of the flow actions. This parameter is mandatory.

            .EXAMPLE
            Search-InFlowActions -FlowActions $Flow_Actions_Array -SearchWord 'body'

            This example searches for the word 'body' within the properties of the flow actions in the $Flow_Actions_Array list.
        #>

        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.Collections.Generic.List[PSCustomObject]]$FlowActions,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$SearchWord
        )

        process
        {
            # Function to Escape Regex Characters
            function Escape-Regex
            {
                param (
                    [string]$String
                )

                $String -replace '([\\.*+?|(){}[\]^$])', '\$1'
            }

            function Search-Recursive
            {
                param (
                    [object]$Object,
                    [string]$Word,
                    [string]$CurrentPath = '',
                    [Hashtable]$ProcessedActions,
                    [Hashtable]$AccumulatedResults,
                    [string]$ActionPath
                )

                $ExcludedProperties = @('DisplayName', 'InternalName', 'Path', 'runAfter', 'type')

                if ($Object -is [string])
                {
                    $EscapedWord = Escape-Regex -String $Word
                    if ($Object -match $EscapedWord)
                    {
                        $FormattedPath = if ($CurrentPath.StartsWith($ActionPath)) { $CurrentPath.Substring($ActionPath.Length) } else { $CurrentPath }
                        $FormattedPath = $FormattedPath -replace '^/', '' # Remove leading slash
                        if (-not $AccumulatedResults.ContainsKey($ActionPath))
                        {
                            $AccumulatedResults[$ActionPath] = @{}
                        }
                        if (-not $AccumulatedResults[$ActionPath].ContainsKey($FormattedPath))
                        {
                            $AccumulatedResults[$ActionPath][$FormattedPath] = New-Object System.Collections.ArrayList
                        }
                        $AccumulatedResults[$ActionPath][$FormattedPath].Add("$($Object)") | Out-Null
                    }

                }
                elseif ($Object -is [System.Management.Automation.PSCustomObject])
                {
                    $ActionIdentifier = if ($Object.Path -and $Object.DisplayName) { $Object.Path + $Object.DisplayName } else { [Guid]::NewGuid().ToString() }

                    if ($ProcessedActions.ContainsKey($ActionIdentifier))
                    {
                        return
                    }
                    $ProcessedActions[$ActionIdentifier] = $true

                    # Define the action path
                    if ($Object.PSObject.Properties.Name -contains 'Path')
                    {
                        $ActionPath = $Object.Path + '/' + $Object.DisplayName
                    }

                    # Recursively search in all properties
                    foreach ($Property in $Object.PSObject.Properties)
                    {
                        if ($Property.Name -notin $ExcludedProperties)
                        {
                            $NewPath = $ActionPath + '/' + $Property.Name
                            if ($Property.Value -is [System.Object[]])
                            {
                                $Value = $Property.Value | Select-Object -First 1
                            }
                            else
                            {
                                $Value = $Property.Value
                            }
                            Search-Recursive -Object $Value -Word $Word -CurrentPath $NewPath -ProcessedActions $ProcessedActions -AccumulatedResults $AccumulatedResults -ActionPath $ActionPath
                        }
                    }
                }
            }

            $ProcessedActions = @{}
            $AccumulatedResults = @{}
            foreach ($action in $FlowActions)
            {
                Search-Recursive -Object $action -Word $SearchWord -ProcessedActions $ProcessedActions -AccumulatedResults $AccumulatedResults -ActionPath ''
            }

            # Display results
            Write-Host ''
            Write-Host "Results for: $($SearchWord)" -BackgroundColor DarkGreen -NoNewline
            Write-Host "`n"

            foreach ($ActionPath in $AccumulatedResults.Keys)
            {
                Write-Host ('Found {0} occurrence{1} in action: ' -f
                    $($AccumulatedResults[$ActionPath].Count),
                    $(($AccumulatedResults[$ActionPath].Count -gt 1) ? 's' : '')
                ) -ForegroundColor Green -NoNewline
                Write-Host "$ActionPath" -ForegroundColor Magenta
                foreach ($PropertyPath in $AccumulatedResults[$ActionPath].Keys)
                {
                    $Matched_String = ($AccumulatedResults[$ActionPath][$PropertyPath]) -join ', '
                    $DisplayPath = if ($PropertyPath -eq '') { 'inputs' } else { "inputs/$PropertyPath" }
                    Write-Host "$($DisplayPath): " -ForegroundColor DarkGray -NoNewline
                    Write-Host "$Matched_String " -ForegroundColor Cyan
                }
                Write-Host ''
            }
        }
    }

    # Function that gets the default download folder for the current user
    function Get-DownloadFolder
    {
        <#
    .SYNOPSIS
        Gets the default download folder for the current user.

    .DESCRIPTION
        This script queries the Windows Registry to find the path of the default download directory for the current user.

    .EXAMPLE
        .\GetDownloadFolder.ps1
    #>
        try
        {
            # Registry key path for the Shell Folders
            $KeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'

            # Get the value of the Downloads folder
            $DownloadFolder = (Get-ItemProperty -Path $KeyPath).'{374DE290-123F-4565-9164-39C4925E467B}'

            # Expand any environment variables in the path
            $ResolvedPath = [Environment]::ExpandEnvironmentVariables($DownloadFolder)

            return $ResolvedPath
        }
        catch
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
            if ($LinkURL -match '^https?://')
            {
                $PowerShellHyperlink = ("`e]8;;{0}`e\{1}`e]8;;`e\" -f $LinkURL, $LinkDisplayText)
            }
            else
            {
                $PowerShellHyperlink = ("`e]8;;file:///{0}`e\{1}`e]8;;`e\" -f $LinkURL, $LinkDisplayText)
            }
            Return $PowerShellHyperlink
        }
        Catch
        {
            Throw
        }
    }

    #EndRegion Functions

}

process
{
    Try
    {
        # Get the Flow details from the URL
        if ($FlowUrl)
        {
            $FlowDetails = Get-FlowDetailsFromUrl -FlowUrl $FlowUrl
            $FlowId = $FlowDetails.FlowId
            $EnvironmentIDs = $FlowDetails.EnvironmentId
        }

        # Check if Az module is installed (Not included in Require statement because of a bug)
        if ($null -eq (Get-Module -Name 'Az' -ListAvailable))
        {
            Throw 'Az module is not installed.'
        }

        #Region Get Environment and Flow
        # Authenticate to Azure AD, get the access token, then connect to SharePoint Online (required to get the flow error details)
        $PA_AccessToken = Connect-AzAccountAndGetAccessToken -TenantId $TenantId -ResourceUrl 'https://service.flow.microsoft.com/'

        # Set the headers for the Power Automate API calls
        $PA_API_Call_Headers = @{
            'Authorization' = "Bearer $($PA_AccessToken.Token)"
            'Content-Type'  = 'application/json'
        }

        # Get available Power Automate environments
        Write-Host "`n[$(Get-Date)] Getting available Power Automate environments..." -ForegroundColor Cyan
        $PA_Environments_API_Call = 'https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01'
        $PA_AvailableEnvironments = Invoke-RestMethod -Method GET -Uri $PA_Environments_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 15 -RetryIntervalSec 5 -MaximumRetryCount 2
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
            $FlowObject = Invoke-RestMethod -Method GET -Uri $PA_Flow_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 15 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck

            # If the Flow is found, break the loop
            If ($FlowObject.error.code -ne 'FlowNotFound')
            {
                $FlowDisplayName = $FlowObject.properties.displayName
                Write-Host ("[{1}] Flow '{2}' found in environment '{3}'." -f
                    "`n",
                    $(Get-Date),
                    $FlowDisplayName,
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
        #EndRegion Get Environment and Flow

        #Region AsZip
        if ($AsZip)
        {

            # Get Flow properties
            $FlowCreatorObject = (Get-AzADUser -ObjectId $FlowObject.properties.creator.objectId -ErrorAction SilentlyContinue) ?? $FlowObject.properties.creator.objectId
            $FlowCreator = "$($FlowCreatorObject.DisplayName) ($($FlowCreatorObject.Mail))"
            $FlowDescription = $FlowObject.properties.definition.description
            $IsSolutionFlow = $null -ne $FlowObject.properties.workflowEntityId

            if ($IsSolutionFlow)
            {
                Write-Host "`n[$(Get-Date)] Flow is inside a solution. Changing target for the retrieval..." -ForegroundColor Cyan

                # Get the access token for Dynamics 365 API (where the solution components are stored)
                $D365_AccessToken = (Connect-AzAccountAndGetAccessToken -ResourceUrl $($PA_Environment.SolutionsEnvironmentUrl) -TenantId $TenantId).Token

                # Set the headers for the Dynamics 365 API calls
                $D365_API_Call_Headers = @{
                    'Authorization' = "Bearer $($D365_AccessToken)"
                    'Content-Type'  = 'application/json'
                }
                $PA_SolutionsFlow_API_Call = "$($PA_Environment.SolutionsEnvironmentUrl)/api/data/v9.2/workflows?`$filter=category%20eq%205%20and%20workflowidunique%20eq%20'$FlowId'&api-version=9.2"
                $SolutionsFlowObject = Invoke-RestMethod -Method GET -Uri $PA_SolutionsFlow_API_Call -Headers $D365_API_Call_Headers -TimeoutSec 15 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck

                if (-not $SolutionsFlowObject.value)
                {
                    Throw "`nSolution flow '$FlowId' not found."
                }
            }

            Write-Host "`n[$(Get-Date)] Requesting Flow as zip file..." -ForegroundColor Cyan
            $DownloadFolderPath = Get-DownloadFolder
            $DownloadPath = Join-Path -Path $DownloadFolderPath -ChildPath "$($FlowDisplayName)_$(Get-Date -Format 'dd_MM_yyyy-HH_mm_ss').zip"

            # Get Access Token for the Business App Platform API
            $BAP_AccessToken = Connect-AzAccountAndGetAccessToken -TenantId $TenantId -ResourceUrl 'https://api.bap.microsoft.com/'

            # Build the API call to export the Flow as a zip file
            $BAP_API_Call_Headers = @{
                'Authorization' = "Bearer $($BAP_AccessToken.Token)"
                'Content-Type'  = 'application/json'
            }
            $BAP_Flow_API_Call = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/$($PA_Environment.ID)/exportPackage?api-version=2016-11-01"
            $ResourceId = ('/providers/Microsoft.Flow/flows/{0}' -f ($IsSolutionFlow ? $SolutionsFlowObject.value.resourceid : $FlowId))
            $ResourceIdEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($ResourceId.ToUpper()))
            $BAP_Body = '{
            "includedResourceIds": ["' + $ResourceId + '"],
            "details": {
            "displayName": "' + $FlowDisplayName + '",
            "description": "' + $FlowDescription + '",
            "creator": "' + $FlowCreator + '",
            "sourceEnvironment": "' + $PA_Environment.ID + '",
            },
            "resources": {
                "' + $ResourceIdEncoded + '": {
                    "id": "' + $ResourceId + '",
                    "creationType": "Existing, New, Update",
                    "suggestedCreationType": "New",
                    "dependsOn": [],
                    "details": {
                        "displayName": "' + $FlowDisplayName + '",
                    },
                    "name": "' + $FlowId + '",
                    "type": "Microsoft.Flow/flows",
                    "configurableBy": "User"
                }
            }
        }'

            # Export the Flow as a zip file
            $FlowZipRequest = Invoke-RestMethod -Method POST -Uri $BAP_Flow_API_Call -Headers $BAP_API_Call_Headers -Body $BAP_Body -TimeoutSec 15 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck
            if ($FlowZipRequest.status -eq 'Succeeded')
            {
                Invoke-WebRequest -Uri $($FlowZipRequest.packageLink.value) -OutFile $DownloadPath -UseBasicParsing
                $DownloadPathLink = New-PowerShellHyperlink -LinkURL $DownloadPath -LinkDisplayText $DownloadPath
                Write-Host ("`nFlow exported as zip file to: {0}" -f $($DownloadPathLink) ) -ForegroundColor Green
            }
            else
            {
                Throw ('Flow export failed: {0}' -f ($null -ne $FlowZipRequest.message ? $FlowZipRequest.message : ("Code: $($FlowZipRequest.errors.code), Messsage: $($FlowZipRequest.errors.message)")))
            }
        }
        #EndRegion AsZip

        #Region AsJSON
        if ($AsJSON)
        {
            # Get the Flow definition
            $PA_JSON_Flow_API_Call = "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$($PA_Environment.ID)/flows/$($FlowId)/exportToARMTemplate?api-version=2016-11-01"
            $PA_Flow_JSON_Object = Invoke-RestMethod -Method POST -Uri $PA_JSON_Flow_API_Call -Headers $PA_API_Call_Headers -TimeoutSec 15 -RetryIntervalSec 5 -MaximumRetryCount 2 -SkipHttpErrorCheck

            # Convert the Flow definition to JSON and copy it to the clipboard
            $Flow_JSON_Definition = $PA_Flow_JSON_Object.template.resources.properties.definition | ConvertTo-Json -Depth 100
            $Flow_JSON_Definition | Set-Clipboard

            # Store the Flow definition in a global variable (to be used by other functions after the script has finished)
            New-Variable -Name 'FlowJsonDefinition' -Value $Flow_JSON_Definition -Force -Option ReadOnly -Scope Global
            Write-Host "`nFlow definition copied to clipboard and stored in variable `$FlowJsonDefinition" -ForegroundColor Green

            # Get an array of all Flow's actions and store it in a global variable (to be used by other functions after the script has finished)
            $Flow_Actions_Array = Get-FlowActions -FlowJsonDefinition $Flow_JSON_Definition
            New-Variable -Name 'FlowActions' -Value $Flow_Actions_Array -Force -Option ReadOnly -Scope Global
            Write-Host "An array of all Flow's actions has been stored in variable `$FlowActions" -ForegroundColor Green
            Write-Host "`nYou can search inside flow actions with function Search-InFlowActions as shown in below example:" -ForegroundColor Green
            Write-Host "Search-InFlowActions -FlowActions `$FlowActions -SearchWord 'upload area'`n" -ForegroundColor White
        }
        #EndRegion AsJSON
    }
    Catch
    {
        Throw
    }
    Finally
    {
        #Disconnect-AzAccount | Out-Null
    }
}