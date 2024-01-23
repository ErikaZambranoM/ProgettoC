# Modificare il Body prima di eseguire lo script (Preso dal trigger della Run su Power Automate)
param (
    [Parameter(Mandatory = $true)][String]$Uri # Preso dal trigger su Power Automate
)

try {
    ################## BODY ##################
    $body = '{}'
    ##########################################

    $method = 'POST'

    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    # Definizione Header in base all'uso Power Automate o Azure Function
    if ($Uri.ToLower().Contains('azurewebsites.net')) { $headers.Add('Content-Type', 'text/plain; charset=utf-8') }
    else { $headers.Add('Content-Type', 'application/json; charset=utf-8') }
    $headers.Add('Accept', 'application/json')

    #$Headers.Add("x-ms-workflow-name", '09d6794e-0e93-4e1a-8d7f-96a60083cf3a')
    #$Headers.Add("x-ms-client-keywords", "testFlow")
    $encodedBody = [System.Text.Encoding]::UTF8.GetBytes($body)
    Invoke-RestMethod -Uri $Uri -Method $method -Headers $headers -Body $encodedBody | Out-Null

    Write-Host 'FLOW TRIGGERED SUCCESSFULLY.' -ForegroundColor Green
}
catch { Throw }

# Client Document Review: https://prod-215.westeurope.logic.azure.com:443/workflows/39090c8a35ab46d1a698faa454452b6c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jox9caWSB_0y4Rh9JJixM5OOlSx065mNuaoGiZ2K8VQ
# Trn TCM -> Client: https://prod-191.westeurope.logic.azure.com:443/workflows/cf5ca2fd88ae4b928f0c0707de8727fd/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_xpvZDTs7E0LMbtFI1uxzqLB4ePP3ozQ7Ezb6a0zbRQ
# Trn VDM -> Client: https://prod-189.westeurope.logic.azure.com:443/workflows/1bf6906ec072441d9c87c6aa2b96d7db/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JCi97qOlItVF70-M-IlwTGCy896RJhFsf7dvSa38R4w
# Trn Client -> TCM: https://prod-06.westeurope.logic.azure.com:443/workflows/61b006ccac844a588abbffa1f2fa1821/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fprKvBNC1Vla25qTLhipTkuXd41P9YRLE8_v7ago-YI
