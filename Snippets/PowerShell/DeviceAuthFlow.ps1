<#
!Found on the web.  I don't remember where, but it's not mine.
This snippet assumes a valid refresh token.  To see how to get one of those, check out:
https://www.thelazyadministrator.com/2019/07/22/connect-and-navigate-the-microsoft-graph-api-with-powershell/#3_Authentication_and_Authorization_Different_Methods_to_Connect
#>

$clientId = '1950a258-227b-4e31-a9cf-717495945fc2'  # This is the standard client ID for Windows Azure PowerShell
$redirectUrl = [System.Uri]'urn:ietf:wg:oauth:2.0:oob' # This is the standard Redirect URI for Windows Azure PowerShell
$tenant = 'tecnimont.onmicrosoft.com'              # TODO - your tenant name goes here
$resource = 'https://graph.microsoft.com/';
$serviceRootURL = "https://graph.microsoft.com//$tenant"
$authUrl = "https://login.microsoftonline.com/$tenant";
$postParams = @{ resource = "$resource"; client_id = "$clientId" }

$response = Invoke-RestMethod -Method POST -Uri "$authurl/oauth2/devicecode" -Body $postParams
Write-Host $response.message
#I got tired of manually copying the code, so I did string manipulation and stored the code in a variable and added to the clipboard automatically
$code = ($response.message -split 'code ' | Select-Object -Last 1) -split ' to authenticate.'
Set-Clipboard -Value $code
Start-Process 'https://microsoft.com/devicelogin' # must complete before the rest of the snippet will work

# Get the initial token
$tokenParams = @{
    grant_type    = 'client_credentials'
    client_id     = $clientId
    client_secret = $Secret
}
$tokenResponse = Invoke-RestMethod -Method POST -Uri "$authurl/oauth2/token" -Body $tokenParams

# Use the Refresh Token
$refreshToken = $tokenResponse.refresh_token
$refreshTokenParams = @{
    grant_type    = 'refresh_token'
    client_id     = "$clientId"
    refresh_token = $refreshToken
}
$tokenResponse = Invoke-RestMethod -Method POST -Uri "$authurl/oauth2/token" -Body $refreshTokenParams
$tokenResponse | Select-Object * | Format-List
Connect-AzAccount -AccessToken $tokenResponse.access_token