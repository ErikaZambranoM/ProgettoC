#Questo script consente di aggiornare qualunque attributo su qualunque sito e su qualunque lista. Va lanciato senza modificare perché è parametrizzato.
param(
	[parameter(Mandatory = $true)][string]$SiteUrl, #URL del sito
	[Switch]$System #System Update (opzionale)
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

# Get all list items
Write-Host "Caricamento '$($listName)'..." -ForegroundColor Cyan
$listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
	[PSCustomObject]@{
		ID       = $_['ID']
		TCM_DN   = $_['Title']
		Rev      = $_['IssueIndex']
		DeptCode = $_['DepartmentCode']
	}
}

Write-Host 'Inizio controllo...' -ForegroundColor Magenta
$itemCounter = 0
foreach ($item in $listItems) {
	Write-Progress -Activity 'Correzione' -Status "$($counter)/$($listItems.Count),  $($item.TCM_DN) Rev.$($item.Rev)" -PercentComplete (($itemCounter++ / $listItems.Count) * 100)

	if ($item.TCM_DN[5] -ne $item.DeptCode) {
		try {
			Set-PnPListItem -List $ListName -Identity $item.ID -Values @{ 'DepartmentCode' = $($item.TCM_DN[5]) } | Out-Null
			$msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - UPDATED"
		}
		catch { $msg = "[ERROR] - List: $($listName) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - FAILED" }
		Write-Log -Message $msg
	}
}
Write-Progress -Activity 'Correzione' -Completed