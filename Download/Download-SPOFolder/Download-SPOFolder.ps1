# Questo script scarica le cartelle da SharePoint a locale.

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
	[parameter(Mandatory = $true)][string]$Link, #URL del sito
	[parameter(Mandatory = $true)][string]$DownloadPath #Cartella di download
)

function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
	)

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else { Write-Host $Message -ForegroundColor Cyan }
}

Try {
	# Ricava le variabili che gli servono dal parametro $Link
	$linkArr = $Link.Replace('%20', ' ').Split('/')
	$DLRelPath = '/' + ($linkArr[3..5] -join '/')
	$siteURL = ($linkArr[0..4] -join '/')
	$folderServerRelURL = '/' + ($linkArr[3..($linkArr.Length)] -join '/')
	$DownloadPath = $DownloadPath.Trim('"')

	$CSVPath = (Read-Host -Prompt 'CSV Path o Transmittal Number').Trim('"')
	if ($CSVPath.ToLower().Contains('.csv')) {
		$csv = Import-Csv $CSVPath -Delimiter ';'
		# Validazione colonne
		$validCols = @('TrnName')
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
		$csv = [PSCustomObject]@{
			TrnName = $CSVPath
			Count   = 1
		}
	}
	else {
		Write-Host 'MODE: ALL LIST' -ForegroundColor Red
		Pause
	}

	# Connect to PnP Online
	Connect-PnPOnline -Url $siteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
	try {
		$DL = Get-PnPList | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }
		if ($null -eq $DL) { throw }
	}
	catch {
		Write-Log '[ERROR] Document Library non riconosciuta.'
		Exit
	}

	# Legge tutta la document library
	Write-Log "Caricamento Document Library '$($DL.Title)'..."
	$listItems = Get-PnPListItem -List $DL.Title -PageSize 5000 | ForEach-Object {
		[PSCustomObject] @{
			ID           = $_['ID']
			ObjectName   = $_['FileLeafRef']
			RelativePath = $_['FileRef']
			ObjectType   = $_['FSObjType'] #1 folder, 0 file
		}
	}
	Write-Log 'Document Library caricata con successo.'

	if ($CSVPath -eq '') { $csv = $listItems | Where-Object -FilterScript { $_.ObjectType -eq 0 -and $_.RelativePath.Contains($folderServerRelURL) } }

	$rowCounter = 0
	Write-Log 'Inizio download...'
	ForEach ($row in $csv) {
		$foundFolder = $folderServerRelURL + '/' + $row.TrnName
		if ($csv.Count -gt 1) {
			$perc = [Math]::Round((($rowCounter / $csv.Count) * 100), 2)
			Write-Progress -Activity 'Download' -Status "$($perc)%" -PercentComplete (($rowCounter++ / $csv.Count) * 100)
		}

		$TempFolderList = $listItems | Where-Object -FilterScript { $_.ObjectType -eq 0 -and $_.RelativePath.Contains($foundFolder) }

		if ($null -eq $TempFolderList) { Write-Log "[WARNING] Folder $($row.TrnName) not found or empty." }
		else {
			$TempFolderList | ForEach-Object {
				#Calcolo cartella locale
				$localFolder = $DownloadPath + $_.RelativePath.Replace($folderServerRelURL, '').Replace('/', '\').Replace($_.ObjectName, '')
				#Se l'elemento è già stato scaricato, vai all'iterazione successiva
				if (Test-Path -Path ("$($localFolder)\$($_.ObjectName)")) { continue }
				#Controllo per creazione cartella locale
				if (!(Test-Path -Path $localFolder)) { New-Item -ItemType Directory -Path $localFolder | Out-Null }
				#Scarica il file in locale
				try {
					Get-PnPFile -ServerRelativeUrl $_.RelativePath -Path $localFolder -Filename $_.ObjectName -AsFile -Force
					Write-Log "[SUCCESS] Downloaded $($localFolder.Replace($DownloadPath,''))$($_.ObjectName)"
				}
				catch { Write-Log "[ERROR] Failed download $($localFolder.Replace($DownloadPath,''))$($_.ObjectName) - $($_)" }
			}
		}
	}
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Download' -Completed }
	Write-Log 'Download completato.'
}
Catch { Throw }