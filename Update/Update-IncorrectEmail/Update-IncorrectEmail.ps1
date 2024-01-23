function Update-IncorrectEmail
{
    param (
        [string]$SiteUrl,
        [string]$ListName,
        [string]$ColumnName,
        [string]$IncorrectEmail,
        [string]$CorrectEmail
    )

    try
    {
        # Connect to SharePoint site
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue

        # Get items from the specified list
        $items = Get-PnPListItem -List $ListName

        $updatedItemCount = 0

        foreach ($item in $items)
        {
            $emailColumnValue = $item[$ColumnName]
            $updated = $false

            if ($emailColumnValue -like "*$IncorrectEmail*")
            {
                $updatedEmailValue = $emailColumnValue -replace $IncorrectEmail, $CorrectEmail
                $item[$ColumnName] = $updatedEmailValue
                Set-PnPListItem -List $ListName -Identity $item.Id -Values @{ $ColumnName = $updatedEmailValue } | Out-Null
                $updatedItemCount++
                $updated = $true
                Write-Host ("Processed item ID: {0}, Updated: {1}" -f $item.Id, $updated) -ForegroundColor Green
            }

        }

        Write-Host ("Total items updated: {0}" -f $updatedItemCount) -ForegroundColor Green
    }
    catch
    {
        Write-Host ("An error occurred: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
    finally
    {
        Disconnect-PnPOnline
    }
}

# Example usage
Update-IncorrectEmail -SiteUrl "https://tecnimont.sharepoint.com/sites/vdm_4191" -ListName "DDDisciplines" -ColumnName "CCRecipients" -IncorrectEmail "FathallahA.Adda@resident.tecnimont.it" -CorrectEmail "A.Adda@resident.tecnimont.it"
