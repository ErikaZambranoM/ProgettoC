# csv: TCM_DN, Rev, ClientCode, LastTransmittal

param (
    [Parameter(Mandatory = $True)][String]$Folder
)

function IsMainContent {
    param (
        [Parameter(Mandatory = $True)][String]$fileName,
        [Parameter(Mandatory = $True)][System.Object]$doc
    )

    $goodNames = (
        "$($doc.TCM_DN).pdf",
        "$($doc.ClientCode).pdf",
        "$($doc.TCM_DN)_$($doc.Rev).pdf",
        "$($doc.ClientCode)_$($doc.Rev).pdf",
        "$($doc.TCM_DN)-$($doc.Rev).pdf",
        "$($doc.ClientCode)-$($doc.Rev).pdf",
        "$($doc.TCM_DN)_IS$($doc.Rev).pdf",
        "$($doc.ClientCode)_IS$($doc.Rev).pdf",
        "$($doc.TCM_DN)-IS$($doc.Rev).pdf",
        "$($doc.ClientCode)-IS$($doc.Rev).pdf"
    )

    foreach ($name in $goodNames) {
        if ($fileName -eq $name) { return $true }
    }
    return $false
}

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = Read-Host -Prompt 'CSV Path o TCM Document Number'
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
elseif ($CSVPath -ne '') {
    $rev = Read-Host -Prompt 'Issue Index'
    $CC = Read-Host -Prompt 'Client Code'
    $trn = Read-Host -Prompt 'Last Transmittal'
    $csv = New-Object -TypeName PSCustomObject @{
        TCM_DN          = $CSVPath
        Rev             = $rev
        ClientCode      = $CC
        LastTransmittal = $trn
        Count           = 1
    }
}

$rowCounter = 0
Write-Host 'Inizio correzione...' -ForegroundColor Cyan
ForEach ($row in $csv) {
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Controllo' -Status "$($row.TCM_DN) - $($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

    $trnPath = "$($Folder)\$($row.LastTransmittal)"
    Get-ChildItem -Path $trnPath -File | Where-Object -FilterScript { !$_.BaseName.ToLower().Contains('_crs') -and $_.BaseName.ToLower().Contains($row.TCM_DN.ToLower()) } | ForEach-Object {
        if (!(IsMainContent $_.Name $row)) {
            Write-Host "File $($_.BaseName) trovato."
            $nativeFolder = "$($_.Directory)\$($row.TCM_DN)_native"
            if (!(Test-Path -Path $nativeFolder)) {
                New-Item -Path $nativeFolder -ItemType Directory | Out-Null
            }
            Move-Item -Path $_.FullName -Destination $nativeFolder -Force
        }
    }
}
Write-Progress -Activity 'Controllo' -Completed
