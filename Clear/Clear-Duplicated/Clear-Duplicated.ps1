# Il CSV deve contenere le colonne TCM_DN e Rev

param(
	[parameter(Mandatory = $true)][string]$SiteCode # Codice sito
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
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}

	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
	$Message = $Message.Replace(' - List: ', ';').Replace(' - Library: ', ';').Replace(' - TCM_DN: ', ';').Replace(' - Rev: ', ';').Replace(' - ID: ', ';').Replace(' - Path: ', ';').Replace(' - ', ';').Replace(': ', ';').Replace("'", ';')
	Add-Content $Path "$FormattedDate;$Message"
}

function Save-Warning {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TCMDocNum,
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Rev,
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ClientCode,
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Trn,
		[string]$CTrn = '',
		[String]$Code = $siteCode
	)

	$Path = "$($PSScriptRoot)\logs\Warning-$($Code)-$(Get-Date -Format 'yyyy_MM_dd').csv";

	if (!(Test-Path -Path $Path)) {
		$newLog = New-Item $Path -Force -ItemType File
		Add-Content $newLog 'TCM_DN;Rev;ClientCode;LastTransmittal;LastClientTransmittal'
	}

	Add-Content $Path "$($TCMDocNum);$($Rev);$($ClientCode);$($Trn);$($CTrn)"
	Write-Host "$($TCMDocNum) - $($Rev) salvato in $($Path)." -ForegroundColor Cyan
}

try {
	$DD = "https://tecnimont.sharepoint.com/sites/$($SiteCode)DigitalDocuments"
	$CD = "https://tecnimont.sharepoint.com/sites/$($SiteCode)DigitalDocumentsC"
	$VD = "https://tecnimont.sharepoint.com/sites/vdm_$($SiteCode)"

	$CSVPath = Read-Host -Prompt 'CSV Path o TCM Document Number'
	if ($CSVPath.ToLower().Contains('.csv')) {
		$csv = Import-Csv -Path $CSVPath -Delimiter ';'
		# Validazione colonne
		$validCols = @('TCM_DN', 'Rev')
		$validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
			if ($_ -in $validCols) { $validCounter++ }
		}
		if ($validCounter -lt $validCols.Count) {
			Write-Host "Colonne obbligatorie mancanti: $($validCols -join ', ')" -ForegroundColor Red
			Exit
		}
	}
	elseif ($CSVPath -ne '') {
		$Rev = Read-Host -Prompt 'Issue Index'
		$csv = [PSCustomObject]@{
			TCM_DN = $CSVPath
			Rev    = $Rev
			Count  = 1
		}
	}
	else {
		Write-Host 'MODE: ALL LIST' -ForegroundColor Red
		Pause
	}

	# Connessione al sito
	$CDConn = Connect-PnPOnline -Url $CD -UseWebLogin -ValidateConnection -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue
	$DDConn = Connect-PnPOnline -Url $DD -UseWebLogin -ValidateConnection -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue
	$VDConn = Connect-PnPOnline -Url $VD -UseWebLogin -ValidateConnection -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue
	$listsArray = Get-PnPList -Connection $CDConn

	$listName = 'Client Document List'

	Write-Log "Caricamento '$($listName)'..."
	$itemList = Get-PnPListItem -List $listName -PageSize 5000 -Connection $CDConn | ForEach-Object {
		[PSCustomObject]@{
			ID            = $_['ID']
			TCM_DN        = $_['Title']
			Rev           = $_['IssueIndex']
			ClientCode    = $_['ClientCode']
			LastTrn       = $_['LastTransmittal']
			LastClientTrn = $_['LastClientTransmittal']
			SrcID         = $_['IDDocumentList']
			Source        = $_['DD_SourceEnvironment']
			DocPath       = $_['DocumentsPath']
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Filtro per tutta la lista
	if ($CSVPath -eq '') { $csv = $itemList }

	$rowCounter = 0
	Write-Log 'Inizio pulizia...'
	ForEach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

		$items = $itemList | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

		if ($null -eq $items) { Write-Log "[WARNING] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - NOT FOUND" }
		elseif ($items.Length -eq 1) { Write-Log "[WARNING] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - NOT DUPLICATED" }
		else {
			Write-Host "TCM DN: $($row.TCM_DN) - Rev: $($row.Rev)" -ForegroundColor Blue

			# Stampa l'elenco dei duplicati
			ForEach ($item in $items) {
				Write-Host "ID: $($item.ID) - Client Code: $($item.ClientCode) - LastTrn: $($item.LastTrn) - LastClientTrn: $($item.LastClientTrn)" -ForegroundColor Magenta
			}

			Start-Process "$($CD)/Lists/ClientDocumentList/CustomView.aspx?FilterField1=Title&FilterValue1=$($row.TCM_DN)&FilterType1=Text"

			$pathSplit = $items[0].DocPath.Split('/')
			$folderRelPath = $pathSplit[5..($pathSplit.Length)] -join '/'
			$DLRelPath = '/' + ($pathSplit[3..5] -join '/')
			$DL = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }

			$fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $folderRelPath -ItemType File -Recursive -Connection $CDConn | Select-Object Name, ServerRelativeUrl

			# Cerca documento nella libreria sorgente
			if ($items[0].Source -eq 'DigitalDocuments') {
				$srcItem = Get-PnPListItem -List 'DocumentList' -Id $items[0].SrcID -Connection $DDConn | ForEach-Object {
					[PSCustomObject]@{
						LastTrn              = $_['LastTransmittal']
						LastTrnDate          = $_['LastTransmittalDate']
						CommentDueDate       = $_['CommentDueDate']
						LastClientTrn        = $_['LastClientTransmittal']
						LastClientTrnDate    = $_['LastClientTransmittalDate']
						ActualDate           = $_['ActualDate']
						ApprovalUser         = $_['ApprovalUser']
						UserApprovalComments = $_['UserApprovalComments']
						ApprovalResult       = $_['ApprovalResult']
						Url                  = "$($DD)/Lists/DocumentList/AllRevisionsView.aspx?FilterField1=Title&FilterValue1=$($row.TCM_DN)&FilterType1=Text"
					}
				}
			}
			elseif ($items[0].Source -eq 'VendorDocuments') {
				$srcItem = Get-PnPListItem -List 'Vendor Documents List' -Id $items[0].SrcID -Connection $VDConn | ForEach-Object {
					[PSCustomObject]@{
						LastTrn              = $_['LastTransmittal']
						LastTrnDate          = $_['LastTransmittalDate']
						CommentDueDate       = $_['CommentDueDate']
						LastClientTrn        = $_['LastClientTransmittal']
						LastClientTrnDate    = $_['LastClientTransmittalDate']
						ActualDate           = $_['LastClientTransmittalDate']
						UserApprovalComments = $_['UserApprovalComments']
						ApprovalResult       = $_['ApprovalResult']
						Url                  = "$($VD)/Lists/VendorDocumentsList/AllRevisions.aspx?FilterField1=VD%5FDocumentNumber&FilterValue1=$($row.TCM_DN)&FilterType1=Text"
					}
				}
			}

			Start-Process $srcItem.Url

			$menu = $host.UI.PromptForChoice('Come procedere?', '',
			([System.Management.Automation.Host.ChoiceDescription[]] @("&Elimina ID: $($items[0].ID)", '&Salta')), 0
			)

			Switch ($menu) {
				0 {
					# Elimina il record dalla Lista
					try {
						Remove-PnPListItem -List $listName -Identity $items[0].ID -Recycle -Force -Connection $CDConn | Out-Null
						Write-Log "[SUCCESS] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - DELETED - ID: '$($items[0].ID)'"
					}
					catch { Write-Log "[ERROR] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - FAILED - ID: '$($items[0].ID)' - $($_)" }

					# Corregge il secondo
					try {
						Set-PnPListItem -List $listName -Identity $items[1].ID -Values @{
							LastTransmittal           = $srcItem.LastTrn
							LastTransmittalDate       = $srcItem.LastTrnDate
							CommentDueDate            = $srcItem.CommentDueDate
							LastClientTransmittal     = $srcItem.LastClientTrn
							LastClientTransmittalDate = $srcItem.LastClientTrnDate
							ApprovalResult            = $srcItem.ApprovalResult
							ApprovalUser              = $srcItem.ApprovalUser
							UserApprovalComments      = $srcItem.UserApprovalComments
							ActualDate                = $srcItem.ActualDate
						} -UpdateType SystemUpdate -Connection $CDConn | Out-Null
						$msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - UPDATED - ID: '$($items[1].ID)'"
					}
					catch {
						try {
							Set-PnPListItem -List $listName -Identity $items[1].ID -Values @{
								LastTransmittal           = $srcItem.LastTrn
								LastTransmittalDate       = $srcItem.LastTrnDate
								LastClientTransmittal     = $srcItem.LastClientTrn
								LastClientTransmittalDate = $srcItem.LastClientTrnDate
								ApprovalResult            = $srcItem.ApprovalResult
							} -UpdateType SystemUpdate -Connection $CDConn | Out-Null
							$msg = "[SUCCESS] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - UPDATED - ID: '$($items[1].ID)'"
						}
						catch { $msg = "[ERROR] - List: $($listName) - TCM_DN: $($row.TCM_DN) - Rev: $($row.Rev) - FAILED - ID: '$($items[1].ID)'" }
					}
					Write-Log -Message $msg

					ForEach ($file in $fileList) {
						$fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem -Connection $CDConn
						If ($null -ne $fileAtt.FieldValues.Transmittal_x0020_Number) {
							try {
								Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
									IDClientDocumentList     = $items[1].ID
									Transmittal_x0020_Number = $srcItem.LastTrn
									TransmittalDate          = $srcItem.LastTrnDate
								} -UpdateType SystemUpdate -Connection $CDConn | Out-Null
								$msg = "[SUCCESS] - Library: $($DL.Title) - File: $($file.Name) - UPDATED - Last Transmittal"
							}
							catch { $msg = "[ERROR] - Library: $($DL.Title) - File: $($file.Name) - FAILED - Last Transmittal" }
							Write-Log -Message $msg
						}
						If ($null -ne $srcItem.LastClientTrn -and !($file.ServerRelativeUrl.Contains('Originals'))) {
							try {
								Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
									LastClientTransmittal     = $srcItem.LastClientTrn
									LastClientTransmittalDate = $srcItem.LastClientTrnDate
									ApprovalResult            = $srcItem.ApprovalResult
									ApprovalUser              = $srcItem.ApprovalUser
									UserApprovalComments      = $srcItem.UserApprovalComments
									ActualDate                = $srcItem.ActualDate
								} -Connection $CDConn | Out-Null
								$msg = "[SUCCESS] - Library: $($DL.Title) - File: $($file.Name) - UPDATED - Last Client Transmittal"
							}
							catch {
								try {
									Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
										LastClientTransmittal     = $srcItem.LastClientTrn
										LastClientTransmittalDate = $srcItem.LastClientTrnDate
										ApprovalResult            = $srcItem.ApprovalResult
									} -Connection $CDConn | Out-Null
									$msg = "[SUCCESS] - Library: $($DL.Title) - File: $($file.Name) - UPDATED - Last Client Transmittal"
								}
								catch { $msg = "[ERROR] - Library: $($DL.Title) - File: $($file.Name) - FAILED - Last Client Transmittal" }
							}
							Write-Log -Message $msg
						}
					}
				}
				1 {
					ForEach ($item in $items) { Save-Warning $item.TCM_DN $item.Rev $item.ClientCode $item.LastTrn $item.LastClientTrn }
				}
			}
		}
	}
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Completed }
	Write-Log 'Operazione completata.'
}
catch { throw }