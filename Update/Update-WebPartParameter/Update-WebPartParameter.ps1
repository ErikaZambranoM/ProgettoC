$SiteUrl = 'https://tecnimont.sharepoint.com/sites/DDWave2'
$PageName = 'Home'
$ParameterInternalName = 'flowUrl'
$ParameterValue = 'ahttps://prod-35.westeurope.logic.azure.com:443/workflows/10127479080b4a608d53b3434c2b9472/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=AWUkMY55qa8X9gOv_kch8u2LJcTxryvxoEPpOw5JhJE'
$WebPartButtonLabel = 'Create Transmittal TEST'
$WebpartTitle = 'DD Transmittal To Customer'
#$Comment = "Updated parameter '$ParameterInternalName' of the webpart '$WebpartTitle' ($WebPartButtonLabel)."

#! check existence of the parameter and page

Try
{
    # Connect to the SharePoint site
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue

    # Get the page
    $Page = Get-PnPClientSidePage -Identity $PageName
    $PageRelativeUrl = $Page.PagesLibrary.RootFolder.ServerRelativeUrl + '/' + $PageName + '.aspx'

    # Get the page file
    $PageFile = Get-PnPFile -Url $PageRelativeUrl -AsListItem

    # Check if the page is checked out
    if ($null -ne $PageFile['CheckoutUser'])
    {
        $CheckedOutBy = $PageFile['CheckoutUser']
        Throw ('Page is currently checked out by: {0} ({1})' -f $CheckedOutBy.LookupValue, $CheckedOutBy.Email)
    }
    else
    {
        Write-Host 'Checking out the page.' -ForegroundColor Yellow
        Set-PnPFileCheckedOut -Url $PageRelativeUrl -ErrorAction Stop -WarningAction Stop
    }

    # Find the webpart by Title and ButtonLabel
    $WebPart = $Page.Controls |
        Select-Object -Property *, @{ Name = 'ParsedProperties'; Expression = { $($_.PropertiesJson | ConvertFrom-Json -AsHashtable -Depth 100) } } |
            Where-Object {
                $_.Title -eq $WebpartTitle -and
                $_.ParsedProperties['buttonLabel'] -eq $WebPartButtonLabel
            }

    # Check if the webpart was found and if it is unique
    if ($null -eq $Webpart)
    {
        Throw 'Webpart not found.'
    }
    if ($Webpart.Count -gt 1)
    {
        Throw "$($Webpart.Count) webparts found."
    }

    # Check if parameter needs to be updated
    if ($Webpart.ParsedProperties[$ParameterInternalName] -ne $ParameterValue)
    {
        # Change the parameter value inside retrieved properties
        $Webpart.ParsedProperties[$ParameterInternalName] = $ParameterValue

        # Convert the properties back to JSON and update the webpart
        $UpdatedPropertiesJSON = ConvertTo-Json $($Webpart.ParsedProperties) -Depth 100
        Set-PnPPageWebPart -Page $PageName -Identity $($Webpart.InstanceId) -PropertiesJson $UpdatedPropertiesJSON -ErrorAction Stop -WarningAction Stop
        Write-Host 'Parameter updated.' -ForegroundColor Green
    }
    else
    {
        Write-Host 'Parameter already has the requested value.' -ForegroundColor Yellow
    }

    # Check in the page
    $Parameters = @{
        Url           = $PageRelativeUrl
        CheckinType   = 'MinorCheckIn'
        Comment       = $Comment
        ErrorAction   = 'Stop'
        WarningAction = 'Stop'
    }
    Set-PnPFileCheckedIn @Parameters
    Write-Host 'Page checked in.' -ForegroundColor Green
}
Catch
{
    Throw
}