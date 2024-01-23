param(
	[parameter(Mandatory = $true)][string]$ProjectCode #URL del sito
)

#Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $ProjectCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;List;Content;Action'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - Document: ', ';').Replace(' - Folder: ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

$siteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
$listName = 'DocumentList'

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
elseif ($CSVPath -ne '') {
	$Rev = Read-Host -Prompt 'Issue Index'
	$csv = [PSCustomObject] @{
		TCM_DN = $CSVPath
		Rev    = $Rev
		Count  = 1
	}
}
else { Write-Host 'MODE: ALL LIST' -ForegroundColor Red }

#Connessione al sito
Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
$listArray = Get-PnPList

Write-Log "Caricamento '$($listName)'..."
$listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
	[PSCustomObject] @{
		ID          = $_['ID']
		TCM_DN      = $_['Title']
		Rev         = $_['IssueIndex']
		DocPath     = $_['DocumentsPath']
		DocMigrated = $_['DocumentsMigrated']
	}
}
Write-Log 'Caricamento lista completato.'

# Per modalità tutta lista: filtra quelli con DocumentsMigrated FALSE
if ($CSVPath -eq '') { $csv = $listItems | Where-Object -FilterScript { $_.DocMigrated -eq $false -and $_.Rev -ne $null -and $_.Rev -ne '' } }

$rowCounter = 0
Write-Log 'Inizio correzione...'
ForEach ($row in $csv) {
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)/$($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

	# Filtra l'item sulla lista
	$item = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

	if ($null -eq $item) { Write-Log "[ERROR] - List: $($ListName) - Document: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
	elseif ($item.Count -gt 1) {
		Write-Log "[WARNING] - List: $($ListName) - Document: $($row.TCM_DN)/$($row.Rev) - DUPLICATED"
		Write-Log "Aperta $($listName) filtrata. Verificare prima di continuare."
		Start-Process "$($siteUrl)/lists/$($listName)/AllRevisionsView.aspx?FilterField1=Title&FilterValue1=$($row.TCM_DN)&FilterType1=Text"
		Pause
	}
	else {
		Write-Host "Documento: $($item.TCM_DN)/$($item.Rev)" -ForegroundColor Blue
		$pathSplit = $item.DocPath.Split('/')
		$DLRelPath = '/' + ($pathSplit[3..5] -join '/')
		$DL = $listArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }
		$relDocPath = $pathSplit[5..$pathSplit.Length] -join '/'
		$newDocFolder = ($pathSplit[5..7] -join '/') + "/$($item.TCM_DN)"
		$newPath = $item.DocPath.Replace("$($item.TCM_DN)-$($item.Rev)", "$($item.TCM_DN)/$($item.Rev)")

		try {
			# Creazione nuova cartella documento
			Resolve-PnPFolder -SiteRelativePath $newDocFolder | Out-Null

			# Aggiornamento record Document Library
			Set-PnPListItem -List $listName -Identity $item.ID -Values @{
				DocumentsMigrated = $true
				DocumentsPath     = $newPath
			} -UpdateType SystemUpdate | Out-Null
			Write-Log "[SUCCESS] - List: $($listName) - Document: $($item.TCM_DN)/$($item.Rev) - UPDATED"

			if ($item.DocPath -ne $newPath) {
				# Spostamento cartella nella nuova posizione
				Move-PnPFolder -Folder $relDocPath -TargetFolder $newDocFolder | Out-Null
				Write-Log "[SUCCESS] - List: $($DL.Title) - Folder: $($pathSplit[-1]) - MOVED"

				# Rinomina della cartella
				Rename-PnPFolder -Folder "$($newDocFolder)/$($pathSplit[-1])" -TargetFolderName $item.Rev | Out-Null
				Write-Log "[SUCCESS] - List: $($DL.Title) - Folder: $($item.Rev) - RENAMED"
			}
		}
		catch {
			Write-Log "[ERROR] - List: $($listName) - Document: $($item.TCM_DN)/$($item.Rev) - FAILED"
			Write-Log 'Aperto Document Path. Verificare prima di continuare.'
			Start-Process ($newPath.Replace("$($item.TCM_DN)/$($item.Rev)", "$($item.TCM_DN)"))
			Pause
		}
	}
}
if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed }
Write-Log 'Aggiornamento completato.'