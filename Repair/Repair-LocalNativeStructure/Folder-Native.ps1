# Creazione cartella tcmDocNum_native e spostamento tutti i file all'interno

param (
    [parameter(Mandatory = $true)][String]$Folder
)

$Folder = $Folder.Trim('"')

$wrongExtension = ('.docx', '.dwg', '.xlsm', '.xls', '.pid')

Write-Host 'Inizio operazione...' -ForegroundColor Cyan
Get-ChildItem -Path $Folder -File -Depth 1 | Where-Object -FilterScript { !($_.Name.Contains($_.FullName.Split('\')[-2])) } | ForEach-Object {
    $nameSplit = $_.BaseName.Split('_')[0]
    $folderName = "$($_.DirectoryName)\$($nameSplit)_native"
    if (!(Test-Path -Path $folderName) ) {
        mkdir $folderName | Out-Null
        Write-Host "Cartella $($folderName) creata." -ForegroundColor Green
    }
    if ($_.Extension.ToLower() -eq '.zip' -and $_.BaseName.ToLower().Contains('_native')) {
        #if ($_.Extension.ToLower() -eq ".zip" -and $_.BaseName.ToLower().Contains($nameSplit.ToLower())) {
        try {
            7z x "$($_.FullName)" -o"$($folderName)" | Out-Null
            Remove-Item -Path $_.FullName | Out-Null
            Write-Host "Archivo $($_.Name) estratto in $($nameSplit)_native." -ForegroundColor Green
        }
        catch {
            Write-Host "Estrazione archivio $($_.Name) fallita." -ForegroundColor Red
            Pause
        }
    }
    elseif ($_.Name.ToLower().Contains('_crs.pdf') -or ($_.Extension.ToLower() -in $wrongExtension) -or ( !($_.BaseName.Contains('_CRS')) -and $_.Extension -eq '.xlsx' )) {
        Move-Item -Path $_.FullName -Destination $folderName
        Write-Host "File $($_.Name) spostato in $($nameSplit)_native." -ForegroundColor Green
    }
}
Write-Host 'Operazione completata.' -ForegroundColor Cyan