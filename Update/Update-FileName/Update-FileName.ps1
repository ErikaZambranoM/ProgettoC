$srcFolder = 'C:\Temp\NuovoImport-A2201\Client to TCM'

Get-ChildItem -Path $srcFolder -File -Recurse | ForEach-Object {
    If ('P' -eq $($_.BaseName)[-1]) {
        $NewName = $($_.BaseName).Substring(0, $_.BaseName.length - 1) + $($_.Extension)
        $Newname
        Rename-Item -Path $_.FullName -NewName $NewName
    }
}