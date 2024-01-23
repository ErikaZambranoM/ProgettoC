<#
	Aggiunta record alla Distribution Matrix (VDM) da CSV.
	Colonne del CSV:
	- Title
	- DisciplineOwnerTCM
	- DisciplinesTCM
	- DisciplinesTCM RoleX
	- DisciplinesTCM RoleO
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

param(
	[parameter(Mandatory = $true)][string]$SiteUrl, #URL del sito
	[Switch]$System
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
		Add-Content $newLog 'Timestamp;Type;ListName;Code;Action'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - Code: ', ';').Replace(' - ', ';').Replace(': ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

# Funzione che restituisce array di ID delle Discipline
function Find-ID {
	param (
		[Parameter(Mandatory = $true)][String]$Items,
		[Array]$List = $discipline,
		[string]$Splitter = $separator
	)

	$array = @()
	$Items.Split($Splitter) | ForEach-Object {
		$currItem = $_.Trim()
		$found = $discipline | Where-Object -FilterScript { $_.Title -eq $currItem }
		if ($null -ne $found) { $array += $found.ID }
		else {
			Write-Host "Discipline '$($_)' not found." -ForegroundColor Red
			Exit
		}
	}
	return $array
}

try {
	# Funzione SystemUpdate
	$System ? ( $updateType = 'SystemUpdate' ) : ( $updateType = 'Update' ) | Out-Null

	$listName = 'Distribution Matrix'
	$separator = ','
	$dList = 'Disciplines'

	# Caricamento CSV/Documento/Tutta la lista
	$CSVPath = (Read-Host -Prompt 'CSV Path o Code').Trim('"')
	if ($CSVPath.ToLower().Contains('.csv')) {
		$csv = Import-Csv -Path $CSVPath -Delimiter ';'
		# Validazione colonne
		$validCols = @('Title', 'DisciplineOwnerTCM', 'DisciplinesTCM')
		$validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
			if ($_ -in $validCols) { $validCounter++ }
		}
		if ($validCounter -lt $validCols.Count) {
			Write-Host "Mandatory column(s) missing: $($validCols -join ', ')" -ForegroundColor Red
			Exit
		}
	}
	else {
		$dOwnerTCM = Read-Host -Prompt 'Discipline Owner TCM'
		$dTCM = Read-Host -Prompt 'Disciplines TCM'
		$csv = [PSCustomObject]@{
			Title              = $CSVPath
			DisciplineOwnerTCM = $dOwnerTCM
			DisciplinesTCM     = $dTCM
			Count              = 1
		}
	}

	# Connessione al sito e calcolo del site Code
	Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
	$siteCode = (Get-PnPWeb).Title.Split(' ')[0]

	# Caricamento lista Distribution Matrix
	Write-Log "Caricamento '$($listName)'..."
	$listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID                 = $_['ID']
			Title              = $_['Title']
			DisciplineOwnerTCM = $_['VD_DisciplineOwnerTCM']
			DisciplinesTCM     = $_['VD_DisciplinesTCM']
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Caricamento lista Disciplines
	Write-Log "Caricamento '$($dList)'..."
	$discipline = Get-PnPListItem -List $dList -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID    = $_['ID']
			Title = $_['Title']
		}
	}
	Write-Log 'Caricamento lista completato.'

	$rowCounter = 0
	Write-Log 'Inizio operazioni...'
	ForEach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity 'Codice' -Status "$($rowCounter+1)/$($csv.Count) - $($row.Title)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

		# Filtro DisciplineOwnerTCM
		$dOwner = Find-ID -Items $row.DisciplineOwnerTCM

		# Filtro DisciplinesTCM
		if ($row.DisciplinesTCM) { $dTCM = Find-ID -Items $row.DisciplinesTCM }
		else { $dTCM = $null }

		# Filtro DisciplinesTCM Role X
		if ($row.'DisciplinesTCM RoleX') { $dTCMx = Find-ID -Items $row.'DisciplinesTCM RoleX' }
		else { $dTCMx = $null }

		# Filtro DisciplinesTCM Role O
		if ($row.'DisciplinesTCM RoleO') { $dTCMo = Find-ID -Items $row.'DisciplinesTCM RoleO' }
		else { $dTCMo = $null }

		# Filtro record sulla Distribution Matrix
		$item = $listItems | Where-Object -FilterScript { $_.Title -eq $row.Title }

		try {
			# Se NON ESISTE, lo crea
			if ($null -eq $item) {
				Add-PnPListItem -List $listName -Values @{
					Title                   = $row.Title
					VD_DisciplineOwnerTCM   = $dOwner
					VD_DisciplinesTCM       = $dTCM ? $dTCM : $null
					VD_DisciplinesTCM_RoleX = $dTCMx ? $dTCMx : $null
					VD_DisciplinesTCM_RoleO = $dTCMo ? $dTCMo : $null
				} | Out-Null
				Write-Log "[SUCCESS] - List: $($listName) - Code: $($row.Title) - ADDED"
			}
			# Se ESISTE, lo aggiorna
			else {
				Set-PnPListItem -List $listName -Identity $item.ID -Values @{
					VD_DisciplineOwnerTCM   = $dOwner
					VD_DisciplinesTCM       = $dTCM ? $dTCM : $null
					VD_DisciplinesTCM_RoleX = $dTCMx ? $dTCMx : $null
					VD_DisciplinesTCM_RoleO = $dTCMo ? $dTCMo : $null
				} -UpdateType $updateType | Out-Null
				Write-Log "[SUCCESS] - List: $($listName) - Code: $($row.Title) - UPDATED"
			}
		}
		catch {
			Write-Log "[ERROR] - List: $($listName) - Code: $($row.Title) - FAILED"
			Throw
		}
	}
	Write-Log 'Operazione completata.'
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity 'Codice' -Completed } }