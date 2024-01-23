# Questo script consente di aggiornare qualunque attributo ovunque.
param(
	[parameter(Mandatory = $true)][String]$SiteUrl # URL sito
)

# Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $siteCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;ListName;ID/Doc;Action;Key;Value;OldValue'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Previous: ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

$filePath = (Read-Host -Prompt 'File Path').Trim('"')

# Caricamento ID / CSV / Documento
$CSVPath = (Read-Host -Prompt 'CSV Path').Trim('"')
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
else { Exit }

if ($SiteUrl.ToLower().Contains('digitaldocumentsc')) { $listName = 'Client Document List' }
else { $listName = 'DocumentList' }

# Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

Write-Log "Caricamento '$($ListName)'..."
$listItems = Get-PnPListItem -List $ListName -PageSize 5000 -Connection $mainConn | ForEach-Object {
	[PSCustomObject]@{
		ID                  = $_['ID']
		TCM_DN              = $_['Title']
		Rev                 = $_['IssueIndex']
		ClientCode          = $_['ClientCode']
		RFI                 = $_['ReasonForIssue']
		DocClass            = $_['DocumentClass']
		Src_ID              = $_['IDDocumentList']
		CommentDueDate      = $_['CommentDueDate']
		LastTransmittal     = $_['LastTransmittal']
		LastTransmittalDate = $_['LastTransmittalDate']
		DocPath             = $_['DocumentsPath']
	}
}
Write-Log 'Caricamento lista completato.'

$rowCounter = 0
Write-Log 'Inizio correzione...'
ForEach ($row in $csv) {
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)-$($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

	$item = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }
	if ($null -eq $item) { $msg = "[ERROR] - List: $($ListName) - ID: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
	elseif ($item.Length -gt 1) { $msg = "[WARNING] - List: $($ListName) - ID: $($row.TCM_DN)/$($row.Rev) - DUPLICATED" }
	else {
		Write-Log "Doc: $($row.TCM_DN)/$($row.Rev)"
		try {
			Set-PnPListItem -List $listName -Identity $item.ID -Values @{
				ApprovalResult = $row.ApprovalResult
			} -UpdateType SystemUpdate | Out-Null
			$msg = "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($row.Rev) - UPDATED"
		}
		catch { $msg = "[ERROR] - List: $($listName) - Doc: $($row.TCM_DN)/$($row.Rev) - FAILED" }
		Write-Log $msg


		$pathSplit = $item.DocPath.Split('/')

		# Upload del file
		try {
			if ($SiteUrl.ToLower().Contains('digitaldocumentsc')) {
				$attFolder = ($pathSplit[5..$pathSplit.Length] -join '/') + '/Attachments'

				# Creazione cartella FromClient
				if (-not (Get-PnPFolder -Url "$($item.DocPath)/Attachments" -ErrorAction SilentlyContinue)) {
					Resolve-PnPFolder -SiteRelativePath $attFolder | Out-Null
					Write-Log "[SUCCESS] - List: $($pathSplit[5]) - Folder: Attachments - CREATED"
				}

				Add-PnPFile -Path $filePath -Folder $attFolder -Values @{
					IssueIndex               = $item.Rev
					ClientCode               = $item.ClientCode
					ReasonForIssue           = $item.RFI
					Transmittal_x0020_Number = $item.LastTransmittal
					TransmittalDate          = $item.LastTransmittalDate
					CommentDueDate           = $item.CommentDueDate
				} | Out-Null
				$msg = "[SUCCESS] - List: $($pathSplit[5]) - File - UPLOADED"
			}
			else {
				$attFolder = ($pathSplit[5..$pathSplit.Length] -join '/') + "/$($item.TCM_DN) - Attachments"

				# Creazione cartella FromClient
				if (-not (Get-PnPFolder -Url "$($item.DocPath)/$($item.TCM_DN) - Attachments" -ErrorAction SilentlyContinue)) {
					Resolve-PnPFolder -SiteRelativePath $attFolder | Out-Null
					Write-Log "[SUCCESS] - List: $($pathSplit[5]) - Folder: Attachments - CREATED"
				}

				Add-PnPFile -Path $filePath -Folder $attFolder -Values @{
					DocumentType   = 'Attachment'
					DocumentNumber = $item.TCM_DN
					IssueIndex     = $item.Rev
					ClientCode     = $item.ClientCode
					ReasonForIssue = $item.RFI
					DocumentClass  = $item.DocClass

				} | Out-Null
				$msg = "[SUCCESS] - List: $($pathSplit[5]) - File - UPLOADED"
			}
		}
		catch { $msg = "[ERROR] - List: $($pathSplit[5]) - File - FAILED ($_)" }
	}
	Write-Log -Message $msg
}
Write-Progress -Activity 'Aggiornamento' -Completed
Write-Log 'Operazione completata.'