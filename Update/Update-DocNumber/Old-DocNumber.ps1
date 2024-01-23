# AGGIORNAMENTO TCM DOCUMENT NUMER E CLIENT CODE VD -> CD
param(
    [parameter(Mandatory=$true)][string]$SiteCode, #Codice del progetto/sito
	[parameter(Mandatory=$true)][String]$TCMDocNumber,
	[parameter(Mandatory=$true)][AllowEmptyString()][String]$IssueIndex
)

#Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Message,
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
	$Message = $Message.Replace(" - List: ",";").Replace(" - ID: ",";").Replace(" - Previous: ",";").Replace(" - ",";").Replace(": ",";")
	Add-Content $logPath "$FormattedDate;$Message"
}

$VDUrl = "https://tecnimont.sharepoint.com/sites/vdm_" + $SiteCode
$CDUrl = "https://tecnimont.sharepoint.com/sites/" + $SiteCode + "DigitalDocumentsC"
if ($IssueIndex -eq '') { $IssueIndex = $null }

$PFSList = "Process Flow Status List"
$VDList = "Vendor Documents List"
$TQDRegistry = "TransmittalQueueDetails_Registry"
$CSReport = "Comment Status Report"
$CDList = "Client Document List"
$CTQDRegistry = "ClientTransmittalQueueDetails_Registry"

#Connessione al sito
Connect-PnPOnline -Url $VDUrl -UseWebLogin -ValidateConnection -ErrorAction Stop
$siteCode = (Get-PnPWeb).Title.Split(" ")[0]

# Scarica la Process Flow Status
Write-Log "Caricamento '$($PFSList)'..."
$PFSItems = Get-PnPListItem -List $PFSList -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID = $_["ID"]
			TCM_DN = $_["VD_DocumentNumber"]
			Rev = $_["VD_RevisionNumber"]
			ClientCode = $_["VD_ClientDocumentNumber"]
			VDL_ID = $_["VD_VDL_ID"]
			Comments_ID = $_["VD_CommentsStatusReportID"]
		}
	}
Write-Log "Caricamento lista completato."

# Filtra elemento sulla Process Flow Status
$PFSItem = $PFSItems | Where-Object -FilterScript {$_.TCM_DN -eq $TCMDocNumber -and $_.Rev -eq $IssueIndex}

if ($null -eq $PFSItem) {
	Write-Host "[ERROR] TCM Document Number non trovato." -ForegroundColor Red
	Exit
}

# Dalla Vendor Documents List ottiene il DocumentPath
$lastTrn = [string]((Get-PnPListItem -List $VDList -Id $PFSItem.VDL_ID -Fields "LastTransmittal").FieldValues).Values
$oldDocPath = [string]((Get-PnPListItem -List $VDList -Id $PFSItem.VDL_ID -Fields "VD_DocumentPath").FieldValues).Values
$pathSplit = $oldDocPath.Split("/")
$subSite = ($pathSplit[0..5]) -join "/"

$newTCM_DN = Read-Host -Prompt "Nuovo TCM Document Number"
$newClientCode = Read-Host -Prompt "Nuovo Client Code"

If ($newTCM_DN -eq ''){ $newTCM_DN = $TCMDocNumber }
If ($newClientCode -eq ''){ $newClientCode = $PFSItem.ClientCode }

$newDocPath = $oldDocPath -replace $TCMDocNumber,$newTCM_DN

# Sulla Process Flow, aggiorna il TCM Document Number
try {
	Set-PnPListItem -List $PFSList -Identity $PFSItem.ID -Values @{
		"VD_DocumentNumber" = $newTCM_DN
		"VD_ClientDocumentNumber" = $newClientCode
	} | Out-Null
	$msg = "[SUCCESS] - List: $($PFSList) - ID: $($PFSItem.ID) - UPDATED - VD_DocumentNumber: '$($newTCM_DN)' - Previous: '$($PFSItem.TCM_DN)'"
}
catch { $msg = "[ERROR] - List: $($PFSList) - ID: $($PFSItem.ID) - FAILED - VD_DocumentNumber: '$($PFSItem.TCM_DN)'" }
Write-Log -Message $msg

# Sulla Comment Status Report, aggiorna il TCM Document Number
if($null -eq $PFSItem.Comments_ID) {
	try {
		Set-PnPListItem -List $CSReport -Identity $PFSItem.Comments_ID -Values @{
			"VD_DocumentNumber" = $newTCM_DN
		} | Out-Null
		$msg = "[SUCCESS] - List: $($CSReport) - ID: $($PFSItem.Comments_ID) - UPDATED - VD_DocumentNumber: '$($newTCM_DN)' - Previous: '$($TCMDocNumber)'"
	}
	catch { $msg = "[ERROR] - List: $($CSReport) - ID: $($PFSItem.Comments_ID) - FAILED - VD_DocumentNumber: '$($TCMDocNumber)'" }
	Write-Log -Message $msg
}

# Sulla VDL, aggiorna il TCM Document Number e il DocumentPath
try {
	Set-PnPListItem -List $VDList -Identity $PFSItem.VDL_ID -Values @{
		"VD_DocumentNumber" = $newTCM_DN
		"VD_DocumentPath" = $newDocPath
		"VD_ClientDocumentNumber" = $newClientCode
	} | Out-Null
	$msg = "[SUCCESS] - List: $($VDList) - ID: $($PFSItem.VDL_ID) - UPDATED - VD_DocumentNumber: '$($newTCM_DN)' - Previous: '$($TCMDocNumber)'"
}
catch { $msg = "[ERROR] - List: $($VDList) - ID: $($PFSItem.VDL_ID) - FAILED - VD_DocumentNumber: '$($oldValue)'" }
Write-Log -Message $msg

if ($lastTrn -ne '') {

	# Scarica la Transmittal Details
	Write-Log "Caricamento '$($TQDRegistry)'..."
	$TQDItems = Get-PnPListItem -List $TQDRegistry -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID = $_["ID"]
			TCM_DN = $_["Title"]
			Rev = $_["IssueIndex"]
			ClientCode = $_["ClientCode"]
		}
	}
	Write-Log "Caricamento lista completato."

	$TQDItem = $TQDItems | Where-Object -FilterScript {$_.TCM_DN -eq $TCMDocNumber -and $_.Rev -eq $IssueIndex}

	# Sulla Transmittal Registry Details, aggiorna il Title (TCM Document Number)
	try {
		Set-PnPListItem -List $TQDRegistry -Identity $TQDItem.ID -Values @{
			"Title" = $newTCM_DN
			"ClientCode" = $newClientCode
		} | Out-Null
		$msg = "[SUCCESS] - List: $($TQDRegistry) - ID: $($TQDItem.ID) - UPDATED - Title: '$($newTCM_DN)' - Previous: '$($TQDItem.TCM_DN)'"
	}
	catch { $msg = "[ERROR] - List: $($TQDRegistry) - ID: $($TQDItem.ID) - FAILED - Title: '$($TQDItem.TCM_DN)'" }
	Write-Log -Message $msg
}

# Connessione al sottosito del Vendor
Connect-PnPOnline -Url $subSite -UseWebLogin -ErrorAction Stop

# Scarica la Document Library del PO
Write-Log "Caricamento '$($pathSplit[6])'..."
$POItems = Get-PnPListItem -List $pathSplit[6] -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID = $_["ID"]
			Name = $_["FileLeafRef"]
			TCM_DN = $_["VD_DocumentNumber"]
			Rev = $_["VD_RevisionNumber"]
			ItemPath = $_["FileRef"]
			ObjectType = $_["FSObjType"] #1 folder, 0 file.
		}
	}
Write-Log "Caricamento lista completato."

# Filtra la cartella del documento
$docFolder = $POItems | Where-Object -FilterScript {($_.TCM_DN -eq $TCMDocNumber -or $_.TCM_DN -eq $newTCM_DN)-and $_.Rev -eq $IssueIndex}

If ($null -eq $docFolder) {
	Write-Log -Message "[ERROR] - Document folder not found" -Error
	Exit
}

$fileList = $POItems | Where-Object -FilterScript {$_.ObjectType -eq 0 -and $_.ItemPath.Contains($docFolder.ItemPath)}

ForEach ($file in $fileList) {
	$newFileName = $file.Name -replace $TCMDocNumber,$newTCM_DN

	# Sulla Document Library del PO, aggiorna il FileLeafRef (Nome file) su ogni file del documento
	try {
		Set-PnPListItem -List $pathSplit[6] -Identity $file.ID -Values @{ "FileLeafRef" = $newFileName } | Out-Null
		$msg = "[SUCCESS] - List: $($pathSplit[6]) - ID: $($file.ID) - UPDATED - Title: '$($newFileName)' - Previous: '$($file.Name)'"
	}
	catch { $msg = "[ERROR] - List: $($pathSplit[6]) - ID: $($file.ID) - FAILED - Title: '$($file.Name)'" }
	Write-Log -Message $msg
}

$newFolderName = $docFolder.Name -replace $TCMDocNumber,$newTCM_DN

# Sulla Document Library del PO, aggiorna il FileLeafRef (Nome cartella) del documento
try {
	Set-PnPListItem -List $pathSplit[6] -Identity $docFolder.ID -Values @{
		"FileLeafRef" = $newFolderName
		"VD_DocumentNumber" = $newTCM_DN
	} | Out-Null
	$msg = "[SUCCESS] - List: $($pathSplit[6]) - ID: $($docFolder.ID) - UPDATED - FolderName: '$($newFolderName)' - Previous: '$($docFolder.Name)'"
}
catch { $msg = "[ERROR] - List: $($pathSplit[6]) - ID: $($docFolder.ID) - FAILED - FolderName: '$($docFolder.Name)'" }
Write-Log -Message $msg

# Se il documento ha un Last Transmittal, aggiorna anche la Client Area
if ($lastTrn -ne '') {
	# Connessione al sito Client
	Connect-PnPOnline -Url $CDUrl -UseWebLogin -ValidateConnection -ErrorAction Stop
	$listsArray = Get-PnPList

	# Scarica la Client Document List
	Write-Log "Caricamento '$($CDList)'..."
	$CDItems = Get-PnPListItem -List $CDList -PageSize 5000 | ForEach-Object {
			[PSCustomObject]@{
				ID = $_["ID"]
				TCM_DN = $_["Title"]
				Rev = $_["IssueIndex"]
				ClientCode = $_["ClientCode"]
				DocPath = $_["DocumentsPath"]
				LastClientTransmittal = $_["LastClientTransmittal"]
			}
		}
	Write-Log "Caricamento lista completato."

	# Filtra documento sulla CDL
	$CDItem = $CDItems | Where-Object -FilterScript {$_.TCM_DN -eq $TCMDocNumber -and $_.Rev -eq $IssueIndex}

	$newDocPath = $CDItem.DocPath -replace $CDItem.ClientCode,$newClientCode
	$pathSplit = $CDItem.DocPath.Split("/")

	$DLRelPath = "/" + ($pathSplit[3..5] -join "/")
	$DL = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }

	# Sulla CDL, aggiorna il TCM Document Number
	try {
		Set-PnPListItem -List $CDList -Identity $CDItem.ID -Values @{
			"Title" = $newTCM_DN
			"ClientCode" = $newClientCode
			"DocumentsPath" = $newDocPath
		} | Out-Null
		$msg = "[SUCCESS] - List: $($CDList) - ID: $($CDItem.ID) - UPDATED - Title: '$($newTCM_DN)' - Previous: '$($TCMDocNumber)'"
	}
	catch { $msg = "[ERROR] - List: $($CDList) - ID: $($CDItem.ID) - FAILED - Title: '$($TCMDocNumber)'" }
	Write-Log -Message $msg

	# Se il documento ha un Last Client Transmittal
	if ($null -ne $CDItem.LastClientTransmittal) {
		# Scarica la Client Transmittal Details Registry
		Write-Host "Caricamento '$($CTQDRegistry)'..." -ForegroundColor Cyan
		$CTQDItems = Get-PnPListItem -List $CTQDRegistry -PageSize 5000 | ForEach-Object {
			[PSCustomObject]@{
				ID = $_["ID"]
				TCM_DN = $_["Title"]
				TrnID = $_["TransmittalID"]
				OriginPath = $_["OriginPath"]
			}
		}
		Write-Log "Caricamento lista completato."

		$CTQDItem = $CTQDItems | Where-Object -FilterScript { $_.TCM_DN -eq $TCMDocNumber -and $_.TrnID -eq $CDItem.LastClientTransmittal }

		# Sulla CDL, aggiorna il TCM Document Number
		try {
			Set-PnPListItem -List $CTQDRegistry -Identity $CTQDItem.ID -Values @{
				"Title" = $newTCM_DN
				"IssueIndex" = $IssueIndex
				"ClientCode" = $newClientCode
				"OriginPath" = $newDocPath
			} | Out-Null
			$msg = "[SUCCESS] - List: $($CTQDRegistry) - ID: $($CTQDItem.ID) - UPDATED - Title: '$($newTCM_DN)' - Previous: '$($TCMDocNumber)'"
		}
		catch { $msg = "[ERROR] - List: $($CTQDRegistry) - ID: $($CTQDItem.ID) - FAILED - Title: '$($TCMDocNumber)'" }
		Write-Log -Message $msg
	}

	# Scarica la Document Library del documento lato Client
	Write-Host "Caricamento '$($DL.Title)'..." -ForegroundColor Cyan
	$CCItems = Get-PnPListItem -List $DL.Title -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID = $_["ID"]
			Name = $_["FileLeafRef"]
			ClientCode = $_["ClientCode"]
			Rev = $_["IssueIndex"]
			ItemPath = $_["FileRef"]
			ObjectType = $_["FSObjType"] #1 folder, 0 file
		}
	}
	Write-Log "Caricamento lista completato."

	$docFolder = $CCItems | Where-Object -FilterScript { $_.Name -eq $CDItem.ClientCode }
	$fileList = $CCItems | Where-Object -FilterScript { $_.ObjectType -eq 0 -and $_.ItemPath.Contains($docFolder.ItemPath) }

	ForEach ($file in $fileList) {
		$newFileName = $file.Name -replace $CDItem.ClientCode,$newClientCode

		# Sulla Document Library lato Client, aggiorna il FileLeafRef (Nome file) su ogni file del documento
		try {
			Set-PnPListItem -List $DL.Title -Identity $file.ID -Values @{
				"FileLeafRef" = $newFileName
				"ClientCode" = $newClientCode
			} | Out-Null
			$msg = "[SUCCESS] - List: $($DL.Title) - ID: $($file.ID) - UPDATED - Name: '$($newFileName)' - Previous: '$($file.Name)'"
		}
		catch { $msg = "[ERROR] - List: $($DL.Title) - ID: $($file.ID) - FAILED - Name: '$($file.Name)'" }
		Write-Log -Message $msg
	}

	# Sulla Document Library lato Client, aggiorna il FileLeafRef (Nome cartella) del documento
	try {
		Set-PnPListItem -List $DL.Title -Identity $docFolder.ID -Values @{ "FileLeafRef" = $newClientCode } | Out-Null
		$msg = "[SUCCESS] - List: $($DL.Title) - ID: $($docFolder.ID) - UPDATED - FolderName: '$($newClientCode)' - Previous: '$($docFolder.Name)'"
	}
	catch { $msg = "[ERROR] - List: $($DL.Title) - ID: $($docFolder.ID) - FAILED - FolderName: '$($docFolder.Name)'" }
	Write-Log -Message $msg
}