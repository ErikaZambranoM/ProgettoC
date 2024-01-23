$TransmittalFoldersPath = 'C:\Temp\NuovoImport-A2201\TRM_TCM to Client'
Get-ChildItem -Path $TransmittalFoldersPath -Directory | ForEach-Object {
    Get-ChildItem -Path $_.FullName -Recurse -File |
        Where-Object -FilterScript {
            !$_.FullName.ToLOwer().Contains('attachments')
        } |
            Move-Item -Destination "$($TransmittalFoldersPath)\$($_.Name)"
            Write-Host "Moving $($_.Name)";
        }