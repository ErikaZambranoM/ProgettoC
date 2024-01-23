# Parameter help description
Param (
    [Parameter(Mandatory = $true, HelpMessage = 'The URL to decode')]
    [String]$URL
)

If ($URL -like '*_layouts/15/AccessDenied.aspx?Source=*') {
    Do {
        Write-Host "A SharePoint online 'Access Denied' URL has been recognized.`nDo you want to decode the link to the target object? " -BackgroundColor Yellow -NoNewline
        Write-Host ''
        $CleanAccessDeniedLink = Read-Host -Prompt '(Y/N) - Type Y for Yes or N for No'
        Switch ($CleanAccessDeniedLink) {
            'Y' {
                $URL = $URL -replace '.*Source=', ''
            }
            'N' {
                Break
            }
            Default {
                Write-Host 'Invalid input. Please type Y for Yes or N for No.' -ForegroundColor Red
                Break
            }
        }
    }
    Until ($CleanAccessDeniedLink -eq 'Y' -or $CleanAccessDeniedLink -eq 'N')
}

$DecodedURL = [system.uri]::UnescapeDataString($URL)
$DecodedURL | Set-Clipboard

Write-Host $DecodedURL -ForegroundColor Green