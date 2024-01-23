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

$TenantId = 'your_tenant_id_here'
$AccessToken = Connect-AzAccountAndGetAccessToken -TenantId $TenantId -ResourceUrl 'https://service.flow.microsoft.com/'


<#
# Create a timer job that will wait for 2 minutes
$timerJob = Start-Job -ScriptBlock {
    Start-Sleep -Seconds 120
    Write-Output "TimeOut"
}

# Run the Connect-AzAccount command
$connectJob = Start-Job -ScriptBlock {
    Connect-AzAccount
}

# Wait for either the Connect-AzAccount command to complete or for the timer to reach zero
$waitResult = Wait-Job -Any $timerJob, $connectJob

# Check which job completed first
if ($waitResult -eq $timerJob) {
    # Timer reached zero, throw an error
    Stop-Job $connectJob
    Remove-Job $connectJob
    Throw "Connect-AzAccount timed out after waiting for 2 minutes."
} else {
    # Connect-AzAccount completed, stop the timer
    Stop-Job $timerJob
    Remove-Job $timerJob
    Receive-Job $connectJob
}

# Cleanup
Remove-Job $timerJob
Remove-Job $connectJob

# Your code continues here

#>