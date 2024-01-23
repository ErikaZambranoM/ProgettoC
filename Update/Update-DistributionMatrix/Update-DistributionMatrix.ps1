#Questo script consente di aggiornare qualunque attributo su qualunque sito e su qualunque lista. Va lanciato senza modificare perché è parametrizzato.
#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
	[parameter(Mandatory = $true)][string]$ProjectCode, #URL del sito
	[Switch]$Client, #DistributionMatrix lato Client (default: TCM)
	[Switch]$System #System Update (opzionale)
)

#Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $ProjectCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;ListName;Code;TeamsTagsIds;Action'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - Code: ', ';').Replace(' - TeamsTagsIds: ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

# Funzione SystemUpdate
$system ? ( $updateType = 'SystemUpdate' ) : ( $updateType = 'Update' ) | Out-Null

$siteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"

#Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop

$teamTags = Get-PnPListItem -List 'TeamTags' -PageSize 5000 | ForEach-Object {
	[PSCustomObject]@{
		ID    = $_['ID']
		Title = $_['Title']
		Tag   = $_['TagId']
	}
}

if ($Client) { $listName = 'DistributionMatrixClientTransmittal' }
else { $listName = 'DistributionMatrix' }

$matrix = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
	[PSCustomObject]@{
		ID             = $_['ID']
		Code           = $_['DocumentCode']
		DisciplineList = $_['DisciplineList']
		TeamTags       = $_['TeamsTagsIds']
	}
}

Write-Log 'Inizio correzione...'
$itemCounter = 0
ForEach ($item in $matrix) {
	Write-Progress -Activity 'Aggiornamento' -Status "$($itemCounter+1)/$($matrix.Length) - $($item.Code)" -PercentComplete (($itemCounter++ / $matrix.Length) * 100)
	$newValue = ''
	$item.DisciplineList | ForEach-Object {
		$dName = $_
		$teamId = $teamTags | Where-Object -FilterScript { $_.Title -eq $dName }
		$newValue += "$($teamId.ID),"
	}
	$newValue = $newValue.Substring(0, $newValue.Length - 1)
	try {
		Set-PnPListItem -List $listName -Identity $item.ID -Values @{TeamsTagsIds = $newValue } -UpdateType $updateType | Out-Null
		Write-Log "[SUCCESS] - List: $($ListName) - Code: $($item.Code) - TeamsTagsIds: $($newValue) - UPDATED"
	}
	catch { Write-Log "[ERROR] - List: $($ListName) - Code: $($item.Code) - TeamTagsIds: $($item.TeamTags) - FAILED" }
	#if ($itemCounter -eq 2) { exit }
}
Write-Progress -Activity 'Aggiornamento' -Completed
Write-Log 'Correzione completata.'