<#
Il CSV deve contenere le colonne TCM_DN e Rev
Funziona sia lato TCM che lato Client
#>
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
	$Message = $Message.Replace(' - List: ', ';').Replace(' - Library: ', ';').Replace(' - Doc: ', ';').Replace(' - ID: ', ';').Replace(' - Folder: ', ';')
	Add-Content $Path "$FormattedDate;$Message"
}

function Remove-Folder {
	param (
		[Parameter(Mandatory=$true)][String]$Name,
		[Parameter(Mandatory=$true)][String]$Path,
		[Parameter(Mandatory=$true)][System.Object]$Document
	)

	$DLName = $Document.DocPath.Split("/")[5]

	try {
		Remove-PnPFolder -Name $Name -Folder $Path -Recycle -Force | Out-Null
		Write-Log "[SUCCESS] - Library: $($DLName) - Doc: $($Document.TCM_DN)/$($Document.Rev) - Folder: $($Name) - DELETED"
	}
	catch {
		Write-Log "[WARNING] - Library: $($DLName) - Doc: $($Document.TCM_DN)/$($Document.Rev) - Folder: $($Name) - FAILED"
		Write-Host "$($_)" -ForegroundColor Red
	}
}

try {
	# Selettore lista principale
	if ($SiteUrl.ToLower().Contains('digitaldocumentsc')) { $listName = 'Client Document List' }
	else { $listName = 'DocumentList' }

	# Dati di input
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
	else {
		$Rev = Read-Host -Prompt 'Issue Index'
		$csv = [PSCustomObject] @{
			TCM_DN = $CSVPath
			Rev    = $Rev
			Count  = 1
		}
	}

	# Modalità di cancellazione
	Write-Log 'Lista opzioni (true/false)'
	$ListRecord = Read-Host -Prompt 'Elemento Lista'
	$Folder = Read-Host -Prompt 'Cartella Rev'

	if ([System.Convert]::ToBoolean($folder) -eq $false) {
		if ($listName -eq 'DocumentList') {
			$CMTD = Read-Host -Prompt 'Cartella CMTD'
			$Native = Read-Host -Prompt 'Cartella Native'
		}
		else {
			$Originals = Read-Host -Prompt 'Cartella Originals'
		}
	}

	#Connessione al sito
	Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
	$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

	# Caricamento lista principale
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

	$rowCounter = 0
	Write-Log 'Inizio operazione...'
	ForEach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Status "$($row.TCM_DN) - $($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }
		Write-Host "Doc: $($row.TCM_DN)/$($row.Rev)" -ForegroundColor Blue
		
		# Filtro documento/revisione
		$items = $itemList | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

		if ($null -eq $items) { Write-Log "[WARNING] - List: $($listName) - Doc: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
		else {
			$itemsCounter = 0

			# Gestione di eventuali duplicati
			ForEach ($item in $items) {

				$itemsCounter++

				if ([System.Convert]::ToBoolean($ListRecord)) {
					# Elimina il record dalla CDL
					try {
						Remove-PnPListItem -List $listName -Identity $item.ID -Recycle -Force | Out-Null
						Write-Log "[SUCCESS] - List: $($listName) - Doc: $($item.TCM_DN)/$($item.Rev) - ID: $($item.ID) - DELETED"
					}
					catch {
						Write-Log "[ERROR] - List: $($listName) - Doc: $($item.TCM_DN)/$($item.Rev) - ID: $($item.ID) - FAILED"
						Throw
					}
				}

				# Se i record sono duplicati, elimina DocumentsPath solo la prima volta.
				if ($itemsCounter -eq 1) {

					# Elimina tutta la cartella della Revisione
					if ([System.Convert]::ToBoolean($Folder)) { Remove-Folder -Name $item.Rev -Path $item.DocPath.Trim("/$($item.Rev)") -Document $item }
					else {
						# Lato TCM, elimina CMTD
						if ([System.Convert]::ToBoolean($CMTD)) { Remove-Folder -Name "CMTD" -Path $item.DocPath -Document $item }

						# Lato TCM, elimina Native
						if ([System.Convert]::ToBoolean($Native)) { Remove-Folder -Name $item.TCM_DN -Path $item.DocPath -Document $item }

						# Lato Client, elimina Original
						if ([System.Convert]::ToBoolean($Originals)) { Remove-Folder -Name "Originals" -Path $item.DocPath -Document $item }
					}
				}
			}
		}
	}
	Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Pulizia' -Completed } }