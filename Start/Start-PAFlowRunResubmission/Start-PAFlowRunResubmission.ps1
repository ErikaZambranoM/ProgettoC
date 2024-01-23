# Requires CSV with following columns: PATicketID;FlowID;RunID;TriggerName;ResubmissionType

Function Convert-PSCustomObjectToList {
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $InputPSCustomObject
    )

    # Get properties of the input PSCustomObject
    $Properties = $InputPSCustomObject.PSObject.Properties | Select-Object Name, Value | Select-Object -ExpandProperty Name

    # Loop through properties and expand nested PSCustomObject properties when needed
    ForEach ($Property In $Properties) {
        $Value = $InputPSCustomObject.$Property

        If ($Value -is [PSCustomObject]) {
            Write-Output ($Value | Format-List | Out-String).Trim()
        }
        Else {
            Write-Output ('{0}: {1}' -f
                $Property,
                $Value
            )
        }
    }
}

Function Invoke-PAHTTPFlow {
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

    Try {
        # Create a new HTTP request to trigger the flow
        $Headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
        $Headers.Add('Content-Type', 'application/json')
        $Headers.Add('CharSet', 'charset=UTF-8')
        $Headers.Add('Accept', 'application/json')
        $EncodedBody = [System.Text.Encoding]::UTF8.GetBytes($Body)

        # Invoke the HTTP request
        $Response = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $EncodedBody

        If (-not $Response) {
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

$AMSResubmitFlowUri = 'https://prod-209.westeurope.logic.azure.com:443/workflows/8a0ba5d97be94e65a619e92acb032496/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jWKnOGirP7PL8RVu6HSW_yau0yJo5lZ3FKpbUFWs2a4'

$RunsToResubmit = Import-Csv -Path 'C:\Users\ST-442\Downloads\INC0868563.csv' -Delimiter ';'
$Counter = 0

ForEach ($Run in $RunsToResubmit) {
    $Counter++
    # Body of the flow resubmission request
    #$ResubmissionType = 'Automatic'
    $AMSResubmitFlowBody = '{
		"PATicketID": "' + $Run.PATicketID + '",
		"FlowID": "' + $Run.FlowID + '",
		"RunID": "' + $Run.RunID + '",
		"TriggerName": "' + $Run.TriggerName + '",
		"ResubmissionType": "Automatic"
	}'

    # Trigger the flow 'AMS - Resubmit flow' to resubmit the flow from Run with given parameters
    Write-Host ('{0}Running automatic resubmission {1}/{2}...' -f "`n", $Counter, $RunsToResubmit.Count) -ForegroundColor Green
    $FlowResubmissionResponse = Invoke-PAHTTPFlow -Uri $AMSResubmitFlowUri -Body $AMSResubmitFlowBody -Method POST
    Write-Host (($FlowResubmissionResponse | Convert-PSCustomObjectToList | Out-String).TrimEnd()) -ForegroundColor Cyan
}