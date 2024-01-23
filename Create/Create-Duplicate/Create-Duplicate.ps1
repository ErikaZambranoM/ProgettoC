<#
    ToDo:
    - compare all duplicates columns
    - 34: ($_.Key -eq "VD_DisciplinesTCM")
#>

Function Create-ListItemDuplicate {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$ListName,
        [Parameter(Mandatory = $true)][int]$ItemId
    )

    Begin {
        Write-Host 'Starting the duplication process...' -ForegroundColor Cyan
    }

    Process {
        Try {
            # Connect to the SharePoint site
            Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

            [Array]$ListColumnsTypes = Get-PnPField -List $ListName | Select-Object -Property InternalName, Title, TypeAsString, ReadOnlyField | Where-Object -FilterScript { $_.ReadOnlyField -eq $false -and $_.Title -ne 'Name' }

            # Get the item to be duplicated
            $item = Get-PnPListItem -List $ListName -Id $ItemId

            # Create a new item with the same field values
            $values = @{}
            $item.FieldValues.GetEnumerator() | ForEach-Object {
                # Skip read-only or hidden fields
                if ($_.Key -in $ListColumnsTypes.InternalName -and !(-not $_.Value)) {
                    if ($_.Value -eq $true) { $values[$_.Key] = $true }
                    elseif ($_.Value -eq $false) { $values[$_.Key] = $false }
                    elseif ($_.Key -eq 'VD_DisciplinesTCM') { $values[$_.Key] = $_.Value.LookupId -join ', ' }
                    elseif ([String]$_.Value -eq 'Microsoft.SharePoint.Client.FieldLookupValue') { $values[$_.Key] = $_.Value.LookupId }
                    else { $values[$_.Key] = $_.Value }
                }
            }

            $values
            Pause

            # Add the new item to the list
            $newItem = Add-PnPListItem -List $ListName -Values $values
            Write-Host "[SUCCESS] - ID: $($newItem.ID) - Item duplication completed successfully!" -ForegroundColor Green
        }
        Catch { Throw }
    }

    End {
        Write-Host 'Duplication process ended.' -ForegroundColor Cyan
    }
}

Create-ListItemDuplicate