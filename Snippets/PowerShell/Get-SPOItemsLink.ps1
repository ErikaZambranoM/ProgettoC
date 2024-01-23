# Function that returns the URL of a filtered List on one or more provided SharePoint list items
Function Get-SPOItemsLink {
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $SPOConnection,

        [Parameter(Mandatory = $true)]
        [String]
        $ListName,

        [Parameter(Mandatory = $true)]
        [Int[]]
        $ItemIDs
    )

    Try {
        $ListObject = Get-PnPList -Identity $ListName -Includes ParentWeb -Connection $SPOConnection
        $ListFilter = ('FilterField{0}1=ID&FilterValue{0}1={1}&FilterType1=Counter' -f
            (($ItemIDs.Count -gt 1) ? 's' : $null),
            $($ItemIDs -Join '%3B%23')
        )
        $FilteredListItemsURL = ('{0}{1}{2}?{3}' -f
            $SPOConnection.Url,
            $ListObject.DefaultViewUrl.Substring(0, $ListObject.DefaultViewUrl.LastIndexOf('/')).Replace($ListObject.ParentWeb.ServerRelativeUrl, ''),
            '/AllItems.aspx',
            $ListFilter
        )

        Return $FilteredListItemsURL
    }
    Catch {
        Throw
    }
}