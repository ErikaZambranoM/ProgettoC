# Questo script consente di aggiornare qualunque attributo ovunque.
param(
	[parameter(Mandatory = $true)][String]$ProjectCode # URL sito
)

# Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $ProjectCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;ListName;Doc;EDMSCode;Action'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - EDMS: ', ';').Replace(' - Doc: ', ';').Replace(' - ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

$TCMUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
$ClientUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsc"
$VDMUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"

$DDL = 'DocumentList'
$CDL = 'Client Document List'
$VDL = 'Vendor Documents List'

# Caricamento CSV / Documento
$CSVPath = (Read-Host -Prompt 'CSV Path o EDMSClientTRCode').Trim('"')
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
elseif ($CSVPath -ne '') {
	$date = Read-Host -Prompt 'Last Client Transmittal Date'
	$csv = [PSCustomObject] @{
		EDMSCode = $CSVPath
		Date     = $date
		Count    = 1
	}
}
else { Exit }

# Connessione al sito
$TCMConn = Connect-PnPOnline -Url $TCMUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
$clientConn = Connect-PnPOnline -Url $ClientUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
$VDMConn = Connect-PnPOnline -Url $VDMUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

Write-Log "Caricamento '$($CDL)'..."
$listItems = Get-PnPListItem -List $CDL -PageSize 5000 -Connection $clientConn | ForEach-Object {
	[PSCustomObject]@{
		ID                = $_['ID']
		TCM_DN            = $_['Title']
		Rev               = $_['IssueIndex']
		SrcID             = $_['IDDocumentList']
		SrcEnv            = $_['DD_SourceEnvironment']
		EDMSCode          = $_['EDMSClientTRCode']
		LastClientTrnDate = $_['LastClientTransmittalDate']
	}
}
Write-Log 'Caricamento lista completato.'

$rowCounter = 0
Write-Log 'Inizio correzione...'
ForEach ($row in $csv) {
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count) - $($row.EDMSCode)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

	[Array]$items = $listItems | Where-Object -FilterScript { $_.EDMSCode -eq $row.EDMSCode }

	if ($null -eq $items) { Write-Log "[ERROR] - List: $($CDL) - Doc: NOT FOUND - EDMSCode: $($row.EDMSCode) - FAILED" }
	else {
		ForEach ($item in $items) {
			# Aggiornamento lato Client
			try {
				Set-PnPListItem -List $CDL -Identity $item.ID -Values @{
					LastClientTransmittalDate = $row.Date
				} -UpdateType SystemUpdate -Connection $clientConn | Out-Null
				$msg = "[SUCCESS] - List: $($CDL) - Doc: $($item.TCM_DN)/$($item.Rev) - EDMSCode: $($row.EDMSCode) - UPDATED"
			}
			catch { $msg = "[ERROR] - List: $($CDL) - Doc: $($item.TCM_DN)/$($item.Rev) - EDMSCode: $($row.EDMSCode) - FAILED - $($_)" }
			Write-Log -Message $msg

			if ($item.SrcEnv -eq 'DigitalDocuments') {
				$srcList = $DDL
				$srcConn = $TCMConn
			}
			else {
				$srcList = $VDL
				$srcConn = $VDMConn
			}

			# Aggiornamento Source
			try {
				Set-PnPListItem -List $srcList -Identity $item.SrcID -Values @{
					LastClientTransmittalDate = $row.Date
				} -UpdateType SystemUpdate -Connection $srcConn | Out-Null
				$msg = "[SUCCESS] - List: $($srcList) - Doc: $($item.TCM_DN)/$($item.Rev) - EDMSCode: $($row.EDMSCode) - UPDATED"
			}
			catch { $msg = "[ERROR] - List: $($srcList) - Doc: $($item.TCM_DN)/$($item.Rev) - EDMSCode: $($row.EDMSCode) - FAILED - $($_)" }
			Write-Log -Message $msg
		}
	}
}
Write-Progress -Activity 'Aggiornamento' -Completed
Write-Log 'Operazione completata.'