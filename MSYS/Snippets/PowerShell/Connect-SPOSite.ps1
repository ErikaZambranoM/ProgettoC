Function Connect-SPOSite
{
    <#
    .SYNOPSIS
        Connects to a SharePoint Online Site or Sub Site.

    .DESCRIPTION
        This function connects to a SharePoint Online Site or Sub Site and returns the connection object.
        If a connection to the specified Site already exists, the function returns the existing connection object.

    .PARAMETER SiteUrl
        Mandatory parameter. Specifies the URL of the SharePoint Online site or subsite.

    .EXAMPLE
        PS C:\> Connect-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/contoso"
        This example connects to the "https://contoso.sharepoint.com/sites/contoso" site.

    .OUTPUTS
        The function returns an object with the following properties:
            - SiteUrl: The URL of the SharePoint Online site or subsite.
            - Connection: The connection object to the SharePoint Online site or subsite as returned by the Connect-PnPOnline cmdlet.
#>

    Param(
        # SharePoint Online Site URL
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                # Match a SharePoint Main Site or Sub Site URL
                If ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$')
                {
                    $True
                }
                Else
                {
                    Throw "`n'$($_)' is not a valid SharePoint Online site or subsite URL."
                }
            })]
        [String]
        $SiteUrl
    )

    Try
    {

        # Initialize Global:SPOConnections array if not already initialized
        If (-not $Script:SPOConnections)
        {
            $Script:SPOConnections = @()
        }
        Else
        {
            # Check if SPOConnection to specified Site already exists
            $SPOConnection = ($Script:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $SiteUrl }).Connection
        }

        # Create SPOConnection to specified Site if not already established
        If (-not $SPOConnection)
        {
            # Create SPOConnection to SiteURL
            Write-Host "Creating connection to '$($SiteUrl)'..." -ForegroundColor Cyan
            $SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

            # Add SPOConnection to the list of connections
            $Script:SPOConnections += [PSCustomObject]@{
                SiteUrl    = $SiteUrl
                Connection = $SPOConnection
            }
        }
        else
        {
            Write-Host "Using existing connection to '$($SiteUrl)'..." -ForegroundColor Cyan
        }

        Return $SPOConnection
    }
    Catch
    {
        Throw
    }
}