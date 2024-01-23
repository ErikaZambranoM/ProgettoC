Add-Type -AssemblyName System.Windows.Forms

$Hashes = @()
$Files2CompareArray = @()

$FileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = "$env:USERPROFILE/Downloads"
    Title            = 'Seleziona i file 2 file da comparare'
    Multiselect      = $true
}

Write-Host 'Seleziona i file 2 file da comparare' -ForegroundColor Yellow
Do {

    $DialogResult = $FileDialog.ShowDialog()
    if ($DialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Host 'Operazione Annullata' -ForegroundColor Red
        Exit
    }
    $Files2CompareArray += $FileDialog.FileNames

    if ($Files2CompareArray.Count -gt 2) {
        Write-Host 'Selezionare solo 2 documenti da comparare' -ForegroundColor Red
        $Files2CompareArray = @()
    }
} Until ($Files2CompareArray.Count -eq 2)


ForEach ($File in $Files2CompareArray) {
    $hashes += (Get-FileHash $file).Hash
}

# Create a Windows Form object
$Form = New-Object System.Windows.Forms.Form

# Set the form as topmost
$Form.TopMost = $true
$Form.Add_Shown({ $Form.Activate() })

# Compare the hashes
if ($hashes[0] -eq $hashes[1]) {
    Write-Host ('File identici:{0}{1}' -f "`n", ($Files2CompareArray -join "`n")) -ForegroundColor Green
    [System.Windows.Forms.MessageBox]::Show($Form, 'UGUALI', 'Risultato HASH', 0, 64) | Out-Null
}
Else {
    Write-Host ('File diversi:{0}{1}' -f "`n", ($Files2CompareArray -join "`n")) -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show($Form, 'DIVERSI', 'Risultato HASH', 0, 16) | Out-Null
}
$Form.Dispose()