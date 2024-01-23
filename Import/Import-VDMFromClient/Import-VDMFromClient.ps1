param(
	[parameter(Mandatory = $true)][string]$ProjectCode
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
		Add-Content $newLog 'Timestamp;Type;ListName;Item;Action'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - Doc: ', ';').Replace(' - Folder: ', ';').Replace(' - File: ', ';').Replace(' - ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

$trnClientPath = (Read-Host -Prompt 'Client Transmittal Path').Trim('"')

$vdmUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
$VDL = 'Vendor Documents List'

$CSVPath = (Read-Host -Prompt 'CSV Path').Trim('"')
if (Test-Path -Path $CSVPath) { $csv = Import-Csv $CSVPath -Delimiter ';' }
else {
	Write-Host 'CSV not found.' -ForegroundColor Red
	Exit
}

$vdmConn = Connect-PnPOnline -Url $vdmURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

Write-Log "Caricamento '$($VDL)'..."
$VDLItems = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $vdmConn | ForEach-Object {
	[PSCustomObject] @{
		ID         = $_['ID']
		TCM_DN     = $_['VD_DocumentNumber']
		Rev        = $_['VD_RevisionNumber']
		ClientCode = $_['VD_ClientDocumentNumber']
		DocPath    = $_['VD_DocumentPath']
	}
}
Write-Log 'Caricamento lista completato.'

$rowCounter = 0
Write-Log 'Inizio operazione...'
ForEach ($row in $csv) {
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Import' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

	# Aggiornamento VDL
	$VDLItem = $VDLItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.'TCM CODE' -and $_.Rev -eq $row.'ISSUE INDEX' }
	if ($null -eq $VDLItem) { $msg = "[ERROR] - List: $($VDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - NOT FOUND" }
	elseif ($VDLItem -is [Array]) { $msg = "[WARNING] - List: $($VDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - DUPLICATED" }
	else {
		Write-Log "Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX')"

		# Creazione cartella FromClient lato VDM
		$vdmPathSplit = $VDLItem.DocPath.Split('/')
		$subSiteUrl = $vdmPathSplit[0..5] -join '/'
		$subConn = Connect-PnPOnline -Url $subSiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
		$fromClientPath = ($vdmPathSplit[6..7] -join '/') + '/FromClient'

		# Creazione cartella FromClient
		if (-not (Get-PnPFolder -Url "$($VDLItem.DocPath)/FromClient" -ErrorAction SilentlyContinue -Connection $subConn)) {
			Resolve-PnPFolder -SiteRelativePath $fromClientPath -Connection $subConn | Out-Null
			try { Set-PnPFolderPermission -List $vdmPathSplit[6] -Identity $fromClientPath -Group "VD $($VDLItem.VendorName)" -RemoveRole 'MT Readers' -Connection $subConn | Out-Null } catch {}
			Write-Log "[SUCCESS] - List: $($vdmPathSplit[6]) - Folder: FromClient - CREATED"
		}

		#Start-Process "$($VDLItem.DocPath)/FromClient"

		try {
			$files = Get-ChildItem -Path $trnClientPath -Recurse -ErrorAction Stop | Where-Object -FilterScript { $_.Name.Contains($VDLItem.ClientCode) -or $_.Name.Contains($VDLItem.TCM_DN) }

			if ($null -eq $files) { $msg = "[WARNING] - LocalFolder: $($trnClientPath) - NOT FOUND OR EMPTY" }
			else {
				ForEach ($file in $files) {
					try {
						$newFileName = $file.Name -replace $VDLItem.ClientCode, $VDLItem.TCM_DN
						Add-PnPFile -Path $file.FullName -Folder $fromClientPath -NewFileName $newFileName -Connection $subConn | Out-Null
						$msg = "[SUCCESS] - List: $($vdmPathSplit[6]) - File: FromClient/$($newFileName) - UPLOADED"
					}
					catch { $msg = "[ERROR] - List: $($vdmPathSplit[6]) - File: FromClient/$($newFileName) - FAILED ($_)" }
				}
			}
		}
		catch {
			Write-Host "[ERROR] - LocalFolder: $($clientTrnPath) - NOT FOUND" -ForegroundColor Red
			Exit
		}
		Write-Log $msg
	}
}
Write-Progress -Activity 'Import' -Completed
Write-Log 'Operazione completata.'