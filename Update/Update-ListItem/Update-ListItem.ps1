<#
	Questo script consente di aggiornare qualunque attributo ovunque.
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
	[parameter(Mandatory = $true)][String]$SiteUrl, # URL sito
	[parameter(Mandatory = $true)][String]$ListName, # Nome lista / DL
	[parameter(Mandatory = $true)][String]$FieldName, # Nome attributo
	[Switch]$Files, # Aggiornamento sui file (opzionale)
	[Switch]$BypassCheck, # Aggiorna i dati dei documenti (opzionale)
	[Switch]$SkipSame, # Salta i valori già uguali (opzionale)
	[Switch]$System # System Update (opzionale)
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
	$Message = $Message.Replace(' - List: ', ';').Replace(' - ID: ', ';').Replace(' - Doc: ', ';').Replace(' - Folder: ', ';').Replace(' - File: ', ';').Replace('/', ';').Replace(' - ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

try {
	# Funzione SystemUpdate
	$system ? ( $updateType = 'SystemUpdate' ) : ( $updateType = 'Update' ) | Out-Null
	$listItems = $null

	# Caricamento ID / CSV / Documento
	$CSVPath = (Read-Host -Prompt 'ID / CSV Path / TCM Document Number').Trim('"')
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
	elseif ($CSVPath -match '^[\d]+$') {
		$csv = [PSCustomObject]@{
			ID    = $CSVPath
			Count = 1
		}
	}
	elseif ($CSVPath -ne '') {
		$rev = Read-Host -Prompt 'Issue Index'
		$csv = [PSCustomObject] @{
			TCM_DN = $CSVPath
			Rev    = $rev
			Count  = 1
		}
	}
	else { Exit }

	# Connessione al sito
	Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
	$siteCode = (Get-PnPWeb).Title.Split(' ')[0]
	$listsArray = Get-PnPList

	# Verifica esistenza attributo
	try {
		$internalName = (Get-PnPField -List $ListName -Identity $FieldName).InternalName
		if ($internalName -notin ($csv | Get-Member -MemberType NoteProperty).Name) {
			$csv | Add-Member -NotePropertyName $internalName -NotePropertyValue $null
		}
	}
	catch {
		# Se non trova l'Internal Name della colonna, interrompe lo script
		Write-Host "[ERROR] Attributo '$FieldName' non trovato." -ForegroundColor Red
		Exit
	}

	$rowCounter = 0
	Write-Log 'Inizio correzione...'
	foreach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)-$($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

		if ($null -ne $row.ID -and !($CSVPath.ToLower().Contains(".csv"))) {
			try {
				$row.$internalName = ((Get-PnPListItem -List $ListName -Id $row.ID -Fields $internalName).FieldValues).Values
				Write-Log "CURRENT $($FieldName): '$($row.$internalName)'"
				if ($row.$internalName -ne '') { Set-Clipboard $row.$internalName }

				# Input nuovo valore
				$newValue = Read-Host -Prompt 'Nuovo Valore'
			}
			catch {
				# Se non trova l'ID sulla lista, interrompe lo script
				Write-Host "[WARNING] Elemento '$($row.ID)' non trovato. Potrebbe essere stato eliminato." -ForegroundColor Yellow
				Exit
			}

			# Aggiornamento Attributo
			try {
				Set-PnPListItem -List $ListName -Identity $row.ID -Values @{
					$internalName = $newValue ? $newValue : $null
				} -UpdateType $updateType | Out-Null
				Write-Log "[SUCCESS] - List: $($ListName) - ID: $($row.ID) - UPDATED`n$($FieldName): $($newValue)`nPrevious: $($row.$internalName)"
			}
			catch { Write-Log "[ERROR] - List: $($ListName) - ID: $($row.ID) - FAILED`n$($_)" }
		}
		else {
			if ($null -eq $listItems) {
				Write-Log "Caricamento '$($ListName)'..."
				$listItems = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
					if ($SiteUrl.ToLower().Contains('vdm_')) {
						[PSCustomObject]@{
							ID            = $_['ID']
							TCM_DN        = $_['VD_DocumentNumber']
							Rev           = $_['VD_RevisionNumber']
							$internalName = $_[$internalName]
						}
					}
					else {
						[PSCustomObject]@{
							ID            = $_['ID']
							TCM_DN        = $_['Title']
							Rev           = $_['IssueIndex']
							$internalName = $_[$internalName]
							DocPath       = $_['DocumentsPath']
						}
					}
				}
				Write-Log 'Caricamento lista completato.'
			}

			if ($row.Rev -eq '') { $found = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN } }
			else { $found = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev } }

			if ($null -eq $found) { Write-Log "[ERROR] - List: $($ListName) - Doc: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
			else {
				foreach ($item in $found) {
					if ($null -eq $row.$internalName) {
						Write-Log "CURRENT $($FieldName): '$($item.$internalName)'"
						if ($null -ne $item.$internalName -and $item.$internalName -ne '') { Set-Clipboard $item.$internalName }
						$newValue = Read-Host -Prompt 'Nuovo Valore'
					}
					else { $newValue = $row.$internalName }

					# Se il valore dell'Item è uguale al nuovo, skippa
					if ($item.$internalName -eq $row.$internalName -and $SkipSame) { continue }

					# Aggiornamento Attributo
					try {
						Set-PnPListItem -List $ListName -Identity $item.ID -Values @{
							$internalName = $newValue ? $newValue : $null
						} -UpdateType $updateType | Out-Null
						Write-Log "[SUCCESS] - List: $($ListName) - Doc: $($item.TCM_DN)/$($item.Rev) - UPDATED`n$($FieldName): $($newValue)`nPrevious: $($item.$internalName)"
					}
					catch { Write-Log "[ERROR] - List: $($ListName) - Doc: $($item.TCM_DN)/$($item.Rev) - FAILED`n$($_)" }

					if ($Files) {
						$pathSplit = $item.DocPath.Split('/')
						$DLRelPath = '/' + ($pathSplit[3..5] -join '/')
						$DL = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }
						$oldPath = ($pathSplit[5..$pathSplit.Length] -join '/')
	
						# Aggiorno la Revisione sui file
						$fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $oldPath -ItemType File -Recursive | Select-Object Name, ServerRelativeUrl
	
						if ($null -eq $fileList) { Write-Log "[WARNING] - List: $($DL.Title) - Folder: $($item.TCM_DN)/$($item.Rev) - EMPTY OR NOT FOUND" }
						else {
							foreach ($file in $fileList) {
								$fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem
								If ($null -ne $fileAtt.FieldValues.$internalName -or $BypassCheck) {
									try {
										Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
											$internalName = $newValue ? $newValue : $null
										} -UpdateType $updateType | Out-Null
										Write-Log "[SUCCESS] - List: $($DL.Title) - File: $($file.Name) - UPDATED`n$($FieldName): $($newValue)`nPrevious: $($item.$internalName)"
									}
									catch { Write-Log "[ERROR] - List: $($DL.Title) - File: $($file.Name) - FAILED`n$($_)" }
								}
							}
						}
					}
				}
			}
		}
		Write-Host ''
	}
	Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed } }