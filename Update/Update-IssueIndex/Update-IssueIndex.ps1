<#
	AGGIORNAMENTO ISSUE INDEX SINGOLO O MASSIVO SU DD O DDC.
	VA FATTO GIRARE UNA VOLTA PER SITO

	La nuova revisione da impostare viene richiesta sulla shell, quindi la Rev nuova sarà uguale per tutti i documenti nel CSV

	ToDo:
		- Add NewRev to CSV
#>
#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
	[parameter(Mandatory = $true)][string]$SiteUrl, #URL del sito
	[Switch]$System #System Update (opzionale)
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
		Add-Content $newLog "Timestamp;Type;ListName;ID;Action;Key;Value;OldValue"
	}
	$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

	if ($Message.Contains("[SUCCESS]")) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains("[ERROR]")) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains("[WARNING]")) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(" - List: ", ";").Replace(" - ID: ", ";").Replace(" - Previous: ", ";")
	Add-Content $logPath "$FormattedDate;$Message"
}

try {
	# Funzione SystemUpdate
	$System ? ( $updateType = "SystemUpdate" ) : ( $updateType = "Update" ) | Out-Null

	# Indentifica il nome della Lista
	if ($SiteUrl.ToLower().Contains("digitaldocumentsc")) { $listName = "Client Document List" }
	elseif ($SiteUrl.ToLower().Contains("digitaldocuments")) { $listName = "DocumentList" }

	# Caricamento CSV/Documento/Tutta la lista
	$CSVPath = Read-Host -Prompt "CSV Path o TCM Document Number"
	if ($CSVPath.ToLower().Contains(".csv")) {
		If (Test-Path -Path $CSVPath) { $csv = Import-Csv -Path $CSVPath -Delimiter ";" }
		Else {
			Write-Host "File '$($CSVPath)' non trovato." -ForegroundColor Red
			Exit
		}
		$CSVColumns = ($csv | Get-Member -MemberType NoteProperty).Name
		If ( $CSVColumns -notcontains 'TCM_DN' -and $CSVColumns -notcontains 'Rev') {
			Write-Host "File '$($CSVPath)' non valido. Colonne necessarie: TCM_DN, Rev" -ForegroundColor Red
			Exit
		}
	}
	elseif ($CSVPath -ne "") {
		$Rev = Read-Host -Prompt "Current Issue Index"
		$csv = [PSCustomObject] @{
			TCM_DN = $CSVPath
			Rev    = $Rev
			Count  = 1
		}
	}
	else {
		Write-Host "MODE: ALL LIST" -ForegroundColor Red
		# Aggiungere qui le variabili da compilare per
		$Rev = Read-Host -Prompt "Filter Issue Index"
		#$Status = Read-Host -Prompt "Filter Status"
	}

	$newRev = Read-Host -Prompt "New Issue Index"

	# Connessione al sito
	Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
	$siteCode = (Get-PnPWeb).Title.Split(" ")[0]
	$listsArray = Get-PnPList

	# Legge tutta la document library
	Write-Log "Caricamento '$($ListName)'..."
	$listItems = Get-PnPListItem -List $ListName -PageSize 5000 | ForEach-Object {
		[PSCustomObject ]@{
			ID                 = $_["ID"]
			TCM_DN             = $_["Title"]
			Rev                = $_["IssueIndex"]
			Status             = $_["DocumentStatus"]
			OwnerDocumentClass = $_["OwnerDocumentClass"]
			LastTrn            = $_["LastTransmittal"]
			DocPath            = $_["DocumentsPath"]
		}
	}
	Write-Log "Caricamento lista completato."

	# Filtro per tutta la lista
	if ($CSVPath -eq "") {
		$csv = $listItems | Where-Object -FilterScript { $_.Rev -eq $Rev }
		#$csv = $listItems | Where-Object -FilterScript { $_.Rev -eq $Rev -and $_.Status -eq $Status }
	}

	$rowCounter = 0
	Write-Log "Inizio aggiornamento..."
	ForEach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity "Aggiornamento" -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

		Write-Host "Doc: $($row.TCM_DN)/$($row.Rev)" -ForegroundColor Blue
		$item = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and ($_.Rev -eq $row.Rev -or $_.Rev -eq $newRev) }

		if ($null -eq $item) { Write-Log "[ERROR] - List: $($ListName) - $($row.TCM_DN) - $($row.Rev) - NOT FOUND" }
		elseif ($item.Length -gt 1) { Write-Log "[WARNING] - List: $($ListName) - $($row.TCM_DN) - $($row.Rev) - DUPLICATED" }
		else {
			$pathSplit = $item.DocPath.Split("/")
			$DLRelPath = "/" + ($pathSplit[3..5] -join "/")
			$DL = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }
			$oldPath = ($pathSplit[5..$pathSplit.Length] -join "/")
			$pathSplit[-1] = $newRev
			$newPath = ($pathSplit -join "/")

			# Aggiorno l'item sulla DL
			try {
				Set-PnPListItem -List $ListName -Identity $item.ID -Values @{
					"IssueIndex"    = $newRev
					"DocumentsPath" = $newPath
				} -UpdateType $updateType | Out-Null
				Write-Log "[SUCCESS] - List: $($ListName) - ID: $($item.ID) - $($item.TCM_DN) - $($newRev) - UPDATE - Previous '$($item.Rev)'"
			}
			catch { Write-Log "[ERROR] - List: $($ListName) - ID: $($item.ID) - $($item.TCM_DN) - $($item.Rev) - FAILED" }

			# Aggiorno la Revisione sui file
			$fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $oldPath -ItemType File -Recursive | Select-Object Name, ServerRelativeUrl

			if ($null -eq $fileList) { Write-Log "[WARNING] - List: $($DL.Title) - Files - $($item.TCM_DN) - $($item.Rev) - EMPTY OR NOT FOUND" }
			else {
				foreach ($file in $fileList) {
					$fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem
					If ($null -ne $fileAtt.FieldValues.IssueIndex) {
						try {
							Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
								'IssueIndex' = $newRev
							} -UpdateType $updateType | Out-Null
							Write-Log "[SUCCESS] - List: $($DL.Title) - FileName: $($file.Name) - $($newRev) - UPDATED - Previous: '$($item.Rev)'"
						}
						catch { Write-Log "[ERROR] - List: $($DL.Title) - FileName: $($file.Name) - $($item.Rev) - FAILED" }
					}
				}
			}

			# Rinomino la cartella nel DocPath
			try {
				Rename-PnPFolder -Folder $oldPath -TargetFolderName $newRev | Out-Null
				Write-Log "[SUCCESS] - List: $($DL.Title) - Folder: $($newRev) - RENAMED - Previous: '$($item.Rev)'"
			}
			catch {
				# Verifica se la cartella contiene file o meno
				if ($null -eq $fileList) {
					try {
						$newRelPath = ($pathSplit[5..$pathSplit.Length] -join "/")
						Resolve-PnPFolder -SiteRelativePath $newRelPath | Out-Null
						Write-Log "[SUCCESS] - List: $($DL.Title) - Folder: $($newRev) - CREATED"
					}
					catch { Write-Log "[ERROR] - List: $($DL.Title) - Folder: $($item.Rev) - FAILED" }
				}
				# Se non riesce a rinominare la cartella perché ne esiste già un'altra con i file dentro, va in Pausa e chiede di risolvere manualmente
				else {
					$errorPath = ($pathSplit[0..($pathSplit.Length - 2)] -join "/")
					Write-Host "[ERROR] - Check manually $($errorPath)" -ForegroundColor Red
					Write-Host "URL copiato nella Clipboard." -ForegroundColor Cyan
					Set-Clipboard -Value $errorPath
					Pause
					continue
				}
			}
		}
		Write-Host ''
	}
	if ($csv.Count -gt 1) { Write-Progress -Activity "Aggiornamento" -Completed }
	Write-Log "Aggiornamento completato."
}
catch { Throw }