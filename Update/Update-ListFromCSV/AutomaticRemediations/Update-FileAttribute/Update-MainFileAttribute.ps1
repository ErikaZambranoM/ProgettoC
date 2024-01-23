
## VERIFICARE QUALI FILE VENGONO AGGIORNATI (NON CRS)
# COMMENTARE / RINOMINARE SCRIPT
# For each Row processed by Update-SPListFromCSV.ps1, run this script to sync matching attributes

#Parametro opzionale per passare il codice del sito dalla shell
#Lo switch $Test permette di usare i siti di Test
param(
	[parameter(Mandatory = $true)][string]$SiteCode,
	[parameter(Mandatory = $true)][ValidateSet('DD', 'CDL', 'VDL')][String]$SiteType,
	[Switch]$Test
)

function Write-Log {
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Message,

		[string]$Path = $logPath
	)

	if (!(Test-Path -Path $Path)) {
		New-Item $Path -Force -ItemType File
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

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
	}

	"$FormattedDate $Message" | Out-File -FilePath $Path -Append
}

#Bypass ExecutionPolicy e recupera la data per il log
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
$ExecutionDate = Get-Date -Format 'yyyy_MM_dd-HH_mm_ss'

#Se il codice del sito non è passato come parametro, viene richiesto al momento dell'esecuzione
if (!$siteCode) { $siteCode = Read-Host -Prompt 'Codice sito' }

#Codifica degli URL dei siti
switch ($SiteType) {
	'DD' {
		if (!$Test) {
			$siteURL = 'https://tecnimont.sharepoint.com/sites/' + $siteCode + 'DigitalDocuments'
		}
		else { $siteURL = 'https://tecnimont.sharepoint.com/sites/DDWave2' }
		$listName = '/Lists/DocumentList'
	}
	'CDL' {
		if (!$Test) {
			$siteUrl = 'https://tecnimont.sharepoint.com/sites/' + $siteCode + 'DigitalDocumentsC'
		}
		else { $siteUrl = 'https://tecnimont.sharepoint.com/sites/DDWave2c' }
		$listName = '/Lists/ClientDocumentList'
	}
	'VDL' {
		if (!$Test) {
			$siteUrl = 'https://tecnimont.sharepoint.com/sites/vdm_' + $siteCode
		}
		else { $siteUrl = 'https://tecnimont.sharepoint.com/sites/POC_vdm' }
		$listName = '/Lists/VendorDocumentsList'
	}
}

#Creazione file di log
$logPath = "$pwd\Logs\$($siteCode)-$SiteType-$($ExecutionDate).log";
Write-Log -Message 'Start' -Path $logPath

$dbPath = Read-Host -Prompt 'Nome file csv'

$database = Import-Csv -Path "$pwd\$dbPath" -Delimiter ';'

#Connessione al sito
Connect-PnPOnline -Url $siteUrl -UseWebLogin -ErrorAction Stop

#Caricamento VDL su Array in locale
$itemList = Get-PnPListItem -List $listName -Query "<View>
		<ViewFields>
			<FieldRef Name='ID'/>
			<FieldRef Name='Title'/>
			<FieldRef Name='IssueIndex'/>
			<FieldRef Name='DocumentsPath'/>
		</ViewFields>
	</View>" -PageSize 5000 | ForEach-Object {
	$item = New-Object PSObject
	$item | Add-Member -MemberType NoteProperty -Name ID -Value $_['ID']
	$item | Add-Member -MemberType NoteProperty -Name TCM_DN -Value $_['Title']
	$item | Add-Member -MemberType NoteProperty -Name Rev -Value $_['IssueIndex']
	$item | Add-Member -MemberType NoteProperty -Name DocPath -Value $_['DocumentsPath']
	$item
}

$rowCounter = 0
ForEach ($row in $database) {
	$rowCounter++
	$item = $itemList | Where-Object -FilterScript { $($_.TCM_DN) -eq $($row.TCM_DN) -and $($_.Rev) -eq $($row.Rev) }
	$relativeURL = $item.DocPath -replace $siteURL, ''

	#Recupera tutti i FILE correlati al documento
	[Array]$folder = Get-PnPFolderItem -FolderSiteRelativeUrl $relativeURL | Where-Object -FilterScript { $($_.TypedObject.ToString()) -eq 'Microsoft.SharePoint.Client.File' -and !$_.Name.Contains('CRS') } | Select-Object Name, ServerRelativeUrl

	#Se la cartella è vuota, salva un Warning nel log
	if ($null -eq $folder) {
		$msg = "[WARNING] - File non trovato - Document: '$($row.TCM_DN)'"
		Write-Log -Message $msg -Path $logPath
	}

	$fileCounter = 0
	ForEach ($file in $folder) {
		$fileCounter++
		#Ottiene l'ID del file tramite il Path relativo
		$fileItem = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem

		#Ottiene il nome della Document Library
		$dir = $($file.ServerRelativeUrl.Split('/')[3])

		$oldAppRes = ((Get-PnPListItem -List $dir -Id $($fileItem.ID) -Fields 'ApprovalResult').FieldValues).Values

		If ($null -ne $oldAppRes) {
			#Se non è null, aggiorna il Client Code sul file
			try {
				Set-PnPListItem -List $dir -Identity $($fileItem.ID) -Values @{'ApprovalResult' = $($row.AppRes) }
				$msg = "[SUCCESS] - $($fileCounter)/$($folder.Count) - File '$($file.Name)' - UPDATED Approval Result: '$($row.AppRes)' - Previous: '$($oldAppRes)'"
			}
			catch {
				$msg = "[ERROR] - $($fileCounter)/$($folder.Count) - File '$($file.Name)' - FAILED Approval Result: '$($oldAppRes)'"
			}
		}
		Write-Log -Message $msg -Path $logPath
	}
}