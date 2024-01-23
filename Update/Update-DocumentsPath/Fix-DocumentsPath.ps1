param(
	[parameter(Mandatory = $true)][string]$SiteUrl #URL del sito
)

#Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $siteCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;ListName;ID;Action;Key;Value;OldValue'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) {
		Write-Host $Message -ForegroundColor Green
	}
	elseif ($Message.Contains('[ERROR]')) {
		Write-Host $Message -ForegroundColor Red
	}
	elseif ($Message.Contains('[WARNING]')) {
		Write-Host $Message -ForegroundColor Yellow
	}
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Previous: ', ';').Replace(' - ', ';').Replace(': ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

$listName = 'DocumentList'

#Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop
$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

#
Write-Log "Caricamento '$($listName)'..."
$itemList = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
	[PSCustomObject]@{
		ID      = $_['ID']
		TCM_DN  = $_['Title']
		Rev     = $_['IssueIndex']
		DocPath = $_['DocumentsPath']
	}
}
Write-Log 'Caricamento lista completato.'

$csv = $itemList | Where-Object -FilterScript { $_.DocPath[-2] -eq '-' }

$itemCounter = 0
Write-Log 'Inizio correzione...'
ForEach ($item in $csv) {
	Write-Progress -Activity 'Aggiornamento' -Status "$($item.TCM_DN) - $($item.Rev)" -PercentComplete (($itemCounter++ / $csv.Count) * 100)

	Write-Host "Current: $($item.DocPath)" -ForegroundColor Green

	Start-Process $item.DocPath

	$newPath = $item.DocPath.Replace("$($item.TCM_DN)-$($item.Rev)", "$($item.TCM_DN)/$($item.Rev)")
	Write-Host "New: $($newPath)" -ForegroundColor Yellow

	Pause

	try {
		Set-PnPListItem -List $listName -Identity $item.ID -Values @{
			DocumentsPath = $newPath
		} -UpdateType SystemUpdate | Out-Null
		$msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - UPDATED"
	}
	catch { $msg = "[ERROR] - List: $($listName) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - FAILED" }
	Write-Log $msg
}
Write-Progress -Activity 'Aggiornamento' -Completed