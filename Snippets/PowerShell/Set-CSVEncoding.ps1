# Manually add the BOM
$CSV_Content = Get-Content -Path $CsvPath -Raw
$CSV_Content = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetPreamble()) + $CSV_Content
Set-Content -Path $CsvPath -Value $CSV_Content -Encoding UTF8