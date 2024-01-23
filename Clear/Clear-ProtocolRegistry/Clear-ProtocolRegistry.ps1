param(
	[parameter(Mandatory = $true)][string]$SiteUrl #URL del sito
)

#Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $siteCode
	)

	$Path = "$($PSScriptRoot)\logs\$($Code)-$(Get-Date -Format 'yyyy_MM_dd').csv";

	if (!(Test-Path -Path $Path)) {
		$newLog = New-Item $Path -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;ListName;TCM_DN;Rev;Action;Value'
	}

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) {	Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}

	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
	$Message = $Message.Replace(' - List: ', ';').Replace(' - TCM_DN: ', ';')
	Add-Content $Path "$FormattedDate;$Message"
}

if ($SiteUrl.ToLower().Contains('digitaldocumentsc')) { $listName = 'ClientProtocolRegistry' }
else { $listName = 'ProtocolRegistry' }

#Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

Write-Log "Caricamento '$($listName)'..."
$itemList = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
	[PSCustomObject]@{
		ID              = $_['ID']
		Title           = $_['Title']
		ProtocolCounter = $_['ProtocolCounter']
	}
}
Write-Log 'Caricamento lista completato.'

#$csv = $itemList | Where-Object -FilterScript { $_.Title.Contains("WTRAN") } | Sort-Object Title | Select-Object -First ($itemList.Length-1000)
try {
	$csv = $itemList | Sort-Object ProtocolCounter | Select-Object -First ($itemList.Length - 1000)
}
catch {
	Write-Host 'Pulizia non necessaria.' -ForegroundColor Green
	Exit
}

$rowCounter = 0
Write-Log 'Inizio operazione...'
ForEach ($row in $csv) {
	Write-Progress -Activity 'Pulizia' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100)

	try {
		Remove-PnPListItem -List $listName -Identity $row.ID -Force | Out-Null
		$msg = "[SUCCESS] - List: $($listName) - Title: $($row.Title) - DELETED"
	}
	catch { $msg = "[WARNING] - List: $($listName) - Title: $($row.Title) - FAILED" }
	Write-Log -Message $msg
}
Write-Progress -Activity 'Pulizia' -Completed
Write-Log 'Operazione completata.'