# Questo script converte gli Excel in tutte le cartelle dei transmittal in xlsx.

param (
    [Parameter(Mandatory = $True)][String]$Folder
)

Function ConvertTo-ExcelXlsx {
    <#
        .SYNOPSIS
            Converts an Excel xls to a xlsx using -ComObject
    #>
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory = $true, ValueFromPipeline)]
        [string]$Path,
        [parameter(Mandatory = $false)]
        [switch]$Force
    )
    Process {
        if (-Not ($Path | Test-Path) ) {
            throw 'File not found'
        }
        if (-Not ($Path | Test-Path -PathType Leaf) ) {
            throw 'Folder paths are not allowed'
        }

        $xlFixedFormat = 51 # Constant for XLSX Workbook
        $xlsFile = Get-Item -Path $Path
        $xlsxPath = '{0}x' -f $xlsFile.FullName

        if ($xlsFile.Extension -ne '.xls') {
            throw 'Expected .xls extension'
        }

        if (Test-Path -Path $xlsxPath) {
            if ($Force) {
                try {
                    Remove-Item $xlsxPath -Force
                }
                catch {
                    throw '{0} already exists and cannot be removed. The file may be locked by another application.' -f $xlsxPath
                }
                Write-Verbose $('Removed {0}' -f $xlsxPath)
            }
            else {
                throw '{0} already exists!' -f $xlsxPath
            }
        }

        try {
            $Excel = New-Object -ComObject 'Excel.Application'
        }
        catch {
            throw 'Could not create Excel.Application ComObject. Please verify that Excel is installed.'
        }

        $Excel.Visible = $false
        $null = $Excel.Workbooks.Open($xlsFile.FullName)
        $Excel.ActiveWorkbook.SaveAs($xlsxPath, $xlFixedFormat)
        $Excel.ActiveWorkbook.Close()
        $Excel.Quit()
    }
}

$files = Get-ChildItem -Path $Folder -File -Recurse | Where-Object -FilterScript { $_.BaseName.ToLower().Contains('_crs') -and $_.Extension -eq '.xls' }

$fileCounter = 0
Write-Host 'Start conversion...' -ForegroundColor Cyan
$files | ForEach-Object {
    if ($files.Count -gt 1) { Write-Progress -Activity 'Converting' -Status "$($fileCounter)/$($files.Count) - $($_.Name)" -PercentComplete (($fileCounter++ / $files.Count) * 100) }

    try {
        ConvertTo-ExcelXlsx $_.FullName -Force
        Write-Host "[SUCCESS] File $($_.BaseName) converted successfully." -ForegroundColor Green
        Remove-Item $_.FullName -Force
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Exit
    }
    Start-Sleep -Seconds 1
}
Write-Progress -Activity 'Converting' -Completed
Write-Host 'Conversion completed.' -ForegroundColor Cyan