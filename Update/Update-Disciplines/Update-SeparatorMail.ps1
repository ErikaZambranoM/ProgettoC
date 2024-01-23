#Questo scritp aggiunge il separator ";" se non c'è

param (
    [Parameter(Mandatory = $true)][string]$Sito,
    [Parameter(Mandatory = $true)][string]$Lista,
    [Parameter(Mandatory = $true)][string]$Field
)
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $codiceSito

    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;Level;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Level: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}
Connect-PnPOnline -Url $Sito -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction Continue

Write-Log "Caricamento lista $($Lista)"
$ListEmails = Get-PnPListItem -List $Lista -PageSize 5000 | ForEach-Object {
    [PSCustomObject]@{
        ID             = $_['ID']
        DepartmentCode = $_['DepartmentCode']
        $Field         = $_[$Field]
    }
}
Write-Log "Lista $($Lista) caricata"
$counterItem = 0
foreach ($item in $ListEmails) {
    Write-Progress -Activity 'Aggiornamento' -Status "$($counterItem+1)/$($ListEmails.Length)" -PercentComplete (($counterItem++ / $ListEmails.Length) * 100)
    try {
        $current = $item.$Field.Split(';')
        for ($i = 0; $i -lt $current.Length; $i++) {
            if ($current[$i].Split('@').Length -gt 2 ) {
                $current[$i] = $current[$i].ToLower().Replace('.it', '.it;').Replace('.com', '.com;').Trim(';')

            }
        }
        $newString = $current -join ';'
        if ($item.$Field -ne $newString) {
            Set-PnPListItem -List $Lista -Identity $item.ID -Values @{ $Field = $newString } | Out-Null
            Write-Log "[SUCCESS] $($Field) - DepartmentCode: $($item.DepartmentCode) "
        }
    }
    catch {
        Write-Log "[ERROR] ID $($item.ID) - DepartmentCode: $($item.DepartmentCode) "
    }
}
Write-Progress -Activity 'Aggiornamento' -Completed
Write-Log "Aggiornamento lista $($Lista) completato"