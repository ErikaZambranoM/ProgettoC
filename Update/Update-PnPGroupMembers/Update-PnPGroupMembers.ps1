param (
    [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$SiteUrl
)

try {
    # Caricamento Group Name o CSV Path
    $CSVPath = (Read-Host -Prompt 'Group Name o CSV Path').Trim('"')
    if ($CSVPath.ToLower().Contains('.csv')) {
        $csv = Import-Csv -Path $CSVPath -Delimiter ';'
        $validCols = @('GroupName', 'Mail')
        $validCounter = 0
        ($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
            if ($_ -in $validCols) { $validCounter++ }
        }
        if ($validCounter -lt $validCols.Count) {
            Write-Host "Missing mandatory columns: $($validCols -join ', ')" -ForegroundColor Red
            Exit
        }
    }
    elseif ($CSVPath -ne '') {
        $mail = Read-Host -Prompt 'Mail address'
        $csv = [PSCustomObject] @{
            GroupName = $CSVPath
            Mail      = $mail
            Count     = 1
        }
    }
    else { Exit }

    $menu = $host.UI.PromptForChoice('Execution mode:', '',
			([System.Management.Automation.Host.ChoiceDescription[]] @('&Add', '&Remove')), 0
    )

    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

    $count = 0
    Write-Host 'Job starting...' -ForegroundColor Cyan
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Update' -Status "$($count+1)/$($csv.Count)" -PercentComplete (($count++ / $csv.Count) * 100) }

        try {
            Switch ($menu) {
                0 {
                    Add-PnPGroupMember -Group $row.GroupName -LoginName $row.Mail | Out-Null
                    Write-Host "[SUCCESS] - Group: $($row.GroupName) - Mail: $($row.Mail) - ADDED" -ForegroundColor Green
                }
                1 {
                    Remove-PnPGroupMember -Group $row.GroupName -LoginName $row.Mail | Out-Null
                    Write-Host "[SUCCESS] - Group: $($row.GroupName) - Mail: $($row.Mail) - REMOVED" -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host "[ERROR] - Group: $($row.GroupName) - Mail: $($row.Mail) - FAILED - $($_)" -ForegroundColor Red
            Exit
        }
    }
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Update' -Completed }
    Write-Host 'Job done.' -ForegroundColor Cyan
}
catch { throw }