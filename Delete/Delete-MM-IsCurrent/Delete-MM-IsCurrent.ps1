# Il CSV deve contenere le colonne TCM_DN
# Questo script elimina gli IsCurrent duplicati importati per errore da MileMate.

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

	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
	$Message = $Message.Replace(' - List: ', ';').Replace(' - TCM_DN: ', ';').Replace(' - Rev: ', ';').Replace(' - ID: ', ';').Replace(' - Path: ', ';').Replace(' - ', ';').Replace(': ', ';').Replace("'", ';')
	Add-Content $Path "$FormattedDate;$Message"
}

#Bypass ExecutionPolicy e recupera la data per il log
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

$listName = 'DocumentList'

$CSVPath = Read-Host -Prompt 'CSV Path o TCM Document Number'
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
else {
	$csv = New-Object -TypeName PSCustomObject -Property @{
		TCM_DN = $CSVPath
		Count  = 1
	}
}

#Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop

$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

Write-Host "Caricamento '$($listName)'..." -ForegroundColor Cyan
$itemList = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
	[PSCustomObject]@{
		ID              = $_['ID']
		TCM_DN          = $_['Title']
		Rev             = $_['IssueIndex']
		Status          = $_['DocumentStatus']
		IsCurrent       = $_['IsCurrent']
		LastTransmittal = $_['LastTransmittal']
	}
}
Write-Host 'Caricamento lista completato.' -ForegroundColor Cyan

$rowCounter = 0
Write-Host 'Inizio pulizia...' -ForegroundColor Cyan
ForEach ($row in $csv) {
	if ($csv.Count -gt 1) {
		Write-Progress -Activity 'Pulizia' -Status "$($row.TCM_DN)" -PercentComplete (($rowCounter++ / $csv.Count) * 100)
	}

	$items = $itemList | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.IsCurrent -eq $true }

	if ($null -eq $items) { $msg = "[WARNING] - List: $($listName) - TCM_DN: $($row.TCM_DN) - NOT FOUND" }
	elseif ($items.Length -eq 1) { $msg = "[WARNING] - List: $($listName) - TCM_DN: $($row.TCM_DN) - NOT DUPLICATED" }
	else {
		ForEach ($item in $items) {
			Write-Host "ID: $($item.ID) - TCM_DN: $($row.TCM_DN) - Rev: $($item.Rev) - Status: $($item.Status) - Is Current: $($item.IsCurrent) - Last Tran.: $($item.LastTransmittal)" -ForegroundColor DarkYellow
		}

		$delCheck = $items | Where-Object -FilterScript { $_.Status -eq 'Placeholder' }

		if ($delCheck.Length -eq 1) { $delItem = $delCheck }
		else {
			$delItem = $items[1]
			<#$menu = $host.UI.PromptForChoice("Quale eliminare?", "",
				([System.Management.Automation.Host.ChoiceDescription[]] @("&ID: $($items[0].ID)", "I&D: $($items[1].ID)", "&Salta")), 1
			)

			Switch ($menu) {
				0 { $delItem = $items[0] }
				1 { $delItem = $items[1] }
				2 { Continue }
			}#>
		}

		try {
			Remove-PnPListItem -List $listName -Identity $delItem.ID -Recycle -Force | Out-Null
			$msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($delItem.TCM_DN) - Rev: $($delItem.Rev) - DELETED - ID: '$($delItem.ID)'"
		}
		catch { $msg = "[ERROR] - List: $($listName) - TCM_DN: $($delItem.TCM_DN) - Rev: $($delItem.Rev) - FAILED - ID: '$($delItem.ID)'" }
	}
	Write-Log -Message $msg
}