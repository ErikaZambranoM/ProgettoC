param (
    [Parameter(Mandatory = $true)][String]$codeProject,
    [Parameter(Mandatory = $true)][string]$RevisioToModify,
    [Parameter(Mandatory = $true)][string]$RevisionUpdate
)
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $ProjectCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID;Action;Key;Value;OldValue'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Previous: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}
# Funzione SystemUpdate
$system ? ( $updateType = 'SystemUpdate' ) : ( $updateType = 'Update' ) | Out-Null

$site = "https://tecnimont.sharepoint.com/sites/vdm_$($codeProject)"
$conn = Connect-PnPOnline -Url $site -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
$VDL = "Vendor Documents List"
$VDLItem = Get-PnPListItem -List $VDL -Connection $conn -PageSize 5000
$Rev = $VDLItem.FieldValues | Where-Object -FilterScript { $_.VD_RevisionNumber -eq $RevisioToModify }
try {
    foreach ($item in $Rev) {
        if ($item.VD_RevisionNumber -eq $RevisioToModify) {
            Set-PnPListItem -List $VDL -Identity $item.ID -Values @{ VD_RevisionNumber = $RevisionUpdate } -Connection $conn -UpdateType $updateType | Out-Null
            Write-Log "[SUCCESS] - List: $($VDL) - ID: $($item.ID) - VD_RevisionNumberOLD: $($RevisioToModify) - VD_RevisionNumberNEW: $($RevisionUpdate) - UPDATED"
        }
    }
}
catch {
    throw
}