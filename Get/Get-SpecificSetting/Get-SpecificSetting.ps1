<#
    ToDo:
        ! - Fix hardcoded values

#>

# Function to Search for a Value in a SharePoint List
Function Search-In-SharePointList {
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param
    (
        [Parameter(Mandatory = $true, HelpMessage = 'Site URL')]
        [string] $SiteUrl,

        [Parameter(Mandatory = $true, HelpMessage = 'Search List')]
        [ValidateSet('Configuration List', 'Settings')]
        [string] $SearchList,

        [Parameter(Mandatory = $true, HelpMessage = 'Search Value')]
        [string] $SearchValue
    )

    Begin {
        # Connect to SharePoint Site
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
    }

    Process {
        Try {
            # Get items from the 'Settings' list where the 'Title' matches the search value
            $items = Get-PnPListItem -List $SearchList | ForEach-Object { #! ConfigurationList / Setting
                [PSCustomObject]@{
                    Title = $_['Title']
                    Value = $_['Value'] #VDM: VD_ConfigValue DD:Value
                }
            }
            $items = $items | Where-Object { $_.Title -eq $SearchValue }

            # Report findings
            if ($items.Count -eq 0) {
                Write-Host "No items found in '$SiteUrl' with the title '$SearchValue' in list '$SearchList'" -ForegroundColor Yellow
            }
            else {
                Write-Host "`nItems found in '$SiteUrl' with the title '$SearchValue' in list '$SearchList':`n$(($items | Format-List | Out-String).Trim())" -ForegroundColor Green
            }
        }
        Catch {
            Throw "An error occurred: $_"
        }
    }

    End {
        # Disconnect from SharePoint Site
        #Disconnect-PnPOnline
    }
}

# Main script starts here
$CSVPath = (Read-Host -Prompt 'Site URL / CSV Path').Trim('"')
$SearchValue = (Read-Host -Prompt "Enter the search value for 'Title'").Trim()

if ($CSVPath.ToLower().Contains('.csv')) {
    $csv = Import-Csv -Path $CSVPath -Delimiter ';'
    $currentIndex = 0

    $csv | ForEach-Object {
        $currentIndex++
        Search-In-SharePointList -SiteUrl $_.SiteUrl -SearchValue $SearchValue -SearchList "Settings"
    }
}
elseif ($CSVPath -ne '') {
    $SiteUrl = $CSVPath
    Search-In-SharePointList -SiteUrl $SiteUrl -SearchValue $SearchValue
}
else {
    Write-Host 'Invalid input. Please provide either a Site URL or a path to a CSV file.'
}
