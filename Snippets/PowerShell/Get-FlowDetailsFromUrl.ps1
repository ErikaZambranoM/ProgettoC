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

$GoodFlowDetails = Get-FlowDetailsFromUrl -FlowUrl 'https://make.powerautomate.com/environments/88888888-a57c-e8c1-8397-85514fffda53/flows/888897d6-8f7d-4dbf-b903-a3e47d30e807'
$GoodSolutionFlowDetails = Get-FlowDetailsFromUrl -FlowUrl 'https://make.powerautomate.com/environments/88888888-a57c-e8c1-8397-85514fffda53/solutions/88886f28-f167-4942-ac9d-be22ba367c40/flows/888897d6-8f7d-4dbf-b903-a3e47d30e807/details?utm_source=solution_explorer'
$WrongFlowDetails = Get-FlowDetailsFromUrl -FlowUrl 'https://make.powerautomate.com/environments/8463f-a57c-e8c1-8397-85514/flows/3c8044-30bc-4e4c-87a4-c41d8?v3=false'