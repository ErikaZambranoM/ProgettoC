<#

TODO:

Add control on Doc Library
Add log
Add try catch
Check Displayed name vs Internal Name


CHECK FOR DL
#>

#Set Parameters for the script
param (
    [parameter(Mandatory = $true)]
    [String]$SiteURL,

    [parameter(Mandatory = $true)]
    [String]$ListName, #displayname

    [parameter(Mandatory = $true)]
    [String]$NewListName #internalName
)

#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -UseWebLogin

#Get the List
$List = Get-PnPList -Identity $ListName -Includes RootFolder

If ($List.BaseType -ne 'DocumentLibrary')
{
    $NewListURL = 'Lists/' + $NewListName
}
Else
{
    $NewListURL = $NewListName
}

#Set new list name
$List.Rootfolder.MoveTo($NewListURL)
Invoke-PnPQuery