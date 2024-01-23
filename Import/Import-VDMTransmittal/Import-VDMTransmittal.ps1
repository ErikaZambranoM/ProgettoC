<#
	Last Transmittal (Folder based on 'Client Transmission' attribute in VDL): Fornire percorso cartella vuota per far scaricare automaticamente i file allo script.
	Creare cartella 'ClientToVDM' e inserire dentro le cartelle con i transmittal from client da importare come descritto di seguito:
		Client Transmittal Path: Fornire percorso cartella in cui vanno preventivamente scaricato i file di risposta dal cliente (dalla root del documento in area C).
			Creare subfolder con nome = $LastClientTransmittal

	ToDo:
		- #! Gestire colonne:
			- Actual Date
			- Approval Result
		- Add logics from Import-FromClientCoverDDtoDDC.ps1 to copy files without download/upload
		- Validate input
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

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

# Funzione che carica la Client Department Code Mapping
function Find-CDCM {
	param (
		[Parameter(Mandatory = $true)]$SiteConn
	)

	Write-Log "Caricamento '$($CDCM)'..."
	$items = Get-PnPListItem -List $CDCM -PageSize 5000 -Connection $SiteConn | ForEach-Object {
		[PSCustomObject] @{
			ID       = $_['ID']
			Title    = $_['Title']
			Value    = $_['Value']
			ListPath = $_['ListPath']
		}
	}
	Write-Log 'Caricamento lista completato.'

	Return $items
}

try {
	$trnVDMPath = (Read-Host -Prompt 'VDM Transmittal Path').Trim('"')
	$trnClientPath = (Read-Host -Prompt 'Client Transmittal Path').Trim('"')

	$vdmUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
	$tcmUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
	$clientUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"

	$VDL = 'Vendor Documents List'
	$CDL = 'Client Document List'
	$CDCM = 'ClientDepartmentCodeMapping'

	$CSVPath = (Read-Host -Prompt 'CSV Path or TCM Document Number').Trim('"')
	if ($CSVPath.ToLower().Contains('.csv')) {
		$csv = Import-Csv -Path $CSVPath -Delimiter ';'
		# Validazione colonne
		$validCols = @('TCM CODE', 'ISSUE INDEX', 'CLIENT CODE', 'LAST TRANSMITTAL', 'LAST TRANSMITTAL DATE', 'COMMENT DUE DATE')
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
		$rev = Read-Host -Prompt 'Issue Index'
		$clientCode = Read-Host -Prompt 'Client Code'
		$lastTrn = Read-Host -Prompt 'Last Transmittal'
		$lastTrnDate = Read-Host -Prompt 'Last Transmittal Date'
		$commentDueDate = Read-Host -Prompt 'Comment Due Date'
		$csv = [PSCustomObject]@{
			'TCM CODE'                = $CSVPath
			'ISSUE INDEX'             = $rev
			'CLIENT CODE'             = $clientCode
			'LAST TRANSMITTAL'        = $lastTrn
			'LAST TRANSMITTAL DATE'   = $lastTrnDate
			'COMMENT DUE DATE'        = $commentDueDate
			'LAST CLIENT TRANSMITTAL' = ''
			Count                     = 1
		}
	}

	# Connessione ai siti
	$vdmConn = Connect-PnPOnline -Url $vdmURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
	$tcmConn = Connect-PnPOnline -Url $tcmUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
	$clientConn = Connect-PnPOnline -Url $clientURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

	# Caricamento Vendor Documents List
	Write-Log "Caricamento '$($VDL)'..."
	$VDLItems = Get-PnPListItem -List $VDL -PageSize 5000 -Connection $vdmConn | ForEach-Object {
		[PSCustomObject] @{
			ID                 = $_['ID']
			TCM_DN             = $_['VD_DocumentNumber']
			Rev                = $_['VD_RevisionNumber']
			ClientCode         = $_['VD_ClientDocumentNumber']
			Index              = $_['VD_Index']
			ReasonForIssue     = $_['VD_ReasonForIssue']
			IsCurrent          = $_['VD_isCurrent']
			DocTitle           = $_['VD_EnglishDocumentTitle']
			VendorName         = $_['VD_VendorName'].LookupValue
			DocClass           = $_['DocumentClass']
			EnableReview       = $_['VD_EnableReviewClient']
			ClientTransmission = $_['VD_ClientTransmission']
			DeptCode4Struct    = $_['VD_DepartmentForStructure_Calc']
			PONumber           = $_['VD_PONumber']
			MRCode             = $_['VD_MRCode']
			DocPath            = $_['VD_DocumentPath']
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Caricamento Client Document List
	Write-Log "Caricamento '$($CDL)'..."
	$CDLItems = Get-PnPListItem -List $CDL -PageSize 5000 -Connection $clientConn | ForEach-Object {
		[PSCustomObject] @{
			ID                        = $_['ID']
			TCM_DN                    = $_['Title']
			Rev                       = $_['IssueIndex']
			ClientCode                = $_['ClientCode']
			Title                     = $_['DocumentTitle']
			LastTransmittal           = $_['LastTransmittal']
			LastTransmittalDate       = $_['LastTransmittalDate']
			CommentDueDate            = $_['CommentDueDate']
			LastClientTransmittal     = $_['LastClientTransmittal']
			LastClientTransmittalDate = $_['LastClientTransmittalDate']
			ApprovalResult            = $_['ApprovalResult']
			DocPath                   = $_['DocumentsPath']
			ReasonForIssue            = $_['ReasonForIssue']
			IDDocumentList            = $_['IDDocumentList']
			Index                     = $_['Index']
			SourceEnvironment         = $_['DD_SourceEnvironment']
			CDC                       = $_['ClientDepartmentCode']
			CDD                       = $_['ClientDepartmentDescription']
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Caricamento Client Department Code Mapping
	$CDCMItems = Find-CDCM -SiteConn $tcmConn

	$rowCounter = 0
	Write-Log 'Inizio operazione...'
	ForEach ($row in $csv) {
		if ($csv.Length -gt 1) { Write-Progress -Activity 'Import' -Status "$($rowCounter+1)/$($csv.Length)" -PercentComplete (($rowCounter++ / $csv.Length) * 100) }

		# Aggiornamento VDL
		$VDLItem = $VDLItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.'TCM CODE' -and $_.Rev -eq $row.'ISSUE INDEX' }
		if ($null -eq $VDLItem) { Write-Log "[ERROR] - List: $($VDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - NOT FOUND" }
		elseif ($VDLItem -is [Array]) { Write-Log "[WARNING] - List: $($VDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - DUPLICATED" }
		else {
			Write-Log "Doc: $($VDLItem.TCM_DN)/$($VDLItem.Rev) - VDM to Client: $($row.'LAST TRANSMITTAL') - Client to VDM: $($row.'LAST CLIENT TRANSMITTAL')"

			# Converte le date nel formato richiesto da SharePoint
			$lastTrnPath = "$($trnVDMPath)\$($row.'LAST TRANSMITTAL')"
			$lastClientTrnPath = "$($trnClientPath)\$($row.'LAST CLIENT TRANSMITTAL')"
			try { $lastTrnDateConv = Get-Date $($row.'LAST TRANSMITTAL DATE') -Format 'MM/dd/yy' } catch { Throw }
			try { $lastClientTrnDateConv = Get-Date $($row.'LAST CLIENT TRANSMITTAL DATE') -Format 'MM/dd/yy' } catch { $lastClientTrnDateConv = $null }
			try { $commentDueDateConv = Get-Date $($row.'COMMENT DUE DATE') -Format 'MM/dd/yy' } catch { $commentDueDateConv = $null }
			try { $actualDateConv = Get-Date $($row.'ACTUAL DATE') -Format 'MM/dd/yy' } catch { $actualDateConv = $null }
			if ( $row.'LAST CLIENT TRANSMITTAL' -eq '') { $commentRequestCalc = $true }
			else { $commentRequestCalc = $false }

			# Aggiornamento record VDL
			try {
				Set-PnPListItem -List $VDL -Identity $VDLItem.ID -Values @{
					LastTransmittal           = $row.'LAST TRANSMITTAL'
					LastTransmittalDate       = $lastTrnDateConv
					CommentDueDate            = $commentDueDateConv
					LastClientTransmittal     = $row.'LAST CLIENT TRANSMITTAL'
					LastClientTransmittalDate = $lastClientTrnDateConv
					ApprovalResult            = $row.'APPROVAL RESULT'
				} -UpdateType SystemUpdate -Connection $vdmConn | Out-Null
				$msg = "[SUCCESS] - List: $($VDL) - Doc: $($VDLItem.TCM_DN)/$($VDLItem.Rev) - UPDATED"
			}
			catch { $msg = "[ERROR] - List: $($VDL) - Doc: $($VDLItem.TCM_DN)/$($VDLItem.Rev) - FAILED" }
			Write-Log $msg

			# Creazione cartella FromClient lato VDM
			$vdmPathSplit = $VDLItem.DocPath.Split('/')
			$subSiteUrl = $vdmPathSplit[0..5] -join '/'
			$subConn = Connect-PnPOnline -Url $subSiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

			# Caricamento file Last Client Transmittal in FromClient (VDM)
			if ($row.'LAST CLIENT TRANSMITTAL' -ne '') {
				$fromClientPath = ($vdmPathSplit[6..7] -join '/') + '/FromClient'
				# Creazione cartella FromClient
				if (-not (Get-PnPFolder -Url "$($VDLItem.DocPath)/FromClient" -ErrorAction SilentlyContinue -Connection $subConn)) {
					Resolve-PnPFolder -SiteRelativePath $fromClientPath -Connection $subConn | Out-Null
					try {
						Set-PnPFolderPermission -List $vdmPathSplit[6] -Identity $fromClientPath -Group "VD $($VDLItem.VendorName)" -RemoveRole 'MT Readers' -Connection $subConn | Out-Null
					}
					catch {}
					Write-Log "[SUCCESS] - List: $($vdmPathSplit[6]) - Folder: FromClient - CREATED"
				}

				# Caricamento file nella FromClient da Locale
				try {
					$files = Get-ChildItem -Path $trnClientPath -Recurse -ErrorAction Stop | Where-Object -FilterScript { $_.Name.Contains($VDLItem.ClientCode) -or $_.Name.Contains($VDLItem.TCM_DN) }

					if ($null -eq $files) { $msg = "[WARNING] - LocalFolder: $($trnClientPath) - Doc: $($row.'LAST CLIENT TRANSMITTAL') - NOT FOUND OR EMPTY" }
					else {
						ForEach ($file in $files) {
							try {
								$newFileName = $file.Name -replace $VDLItem.ClientCode, $vdlitem.TCM_DN
								Add-PnPFile -Path $file.FullName -Folder $fromClientPath -NewFileName $newFileName -Connection $subConn | Out-Null
								$msg = "[SUCCESS] - List: $($vdmPathSplit[6]) - File: FromClient/$($newFileName) - UPLOADED"
							}
							catch { $msg = "[ERROR] - List: $($vdmPathSplit[6]) - File: FromClient/$($newFileName) - FAILED - $($_)" }
						}
					}
				}
				catch {
					Write-Host "[ERROR] - LocalFolder: $($clientTrnPath) - NOT FOUND" -ForegroundColor Red
					Exit
				}
				Write-Log $msg
			}

			# Aggiornamento lato Client
			$CDLitem = $CDLItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.'TCM CODE' -and $_.Rev -eq $row.'ISSUE INDEX' }

			if ($CDLitem -is [Array]) { $msg = "[WARNING] - List: $($CDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - DUPLICATED" }
			elseif ($null -eq $CDLitem) {
				$CDCMfound = $null
				# Ricerca path nella Client Department Code Mapping
				While (!$CDCMfound) {
					[Array]$CDCMfound = $CDCMItems | Where-Object -FilterScript { $_.Title -eq $VDLItem.DeptCode4Struct }
					if ($null -eq $CDCMfound) {
						Write-Log "[WARNING] - List: $($CDCM) - DeptCode4Struct: $($VDLItem.DeptCode4Struct) - NOT FOUND"
						Write-Log 'Aggiungere il record sulla lista e riprovare.'
						Pause
						$CDCMItems = Find-CDCM -SiteConn $tcmConn
					}
					else {
						Write-Log "[SUCCESS] - List: $($CDCM) - DeptCode4Struct: $($VDLItem.DeptCode4Struct) - FOUND"
						$CDCMItem = $CDCMfound[0]
					}
				}
				$docNumberSplit = $VDLItem.TCM_DN.Split('-')
				$clientRelDocPath = "$($CDCMItem.ListPath)/$($docNumberSplit[1][1])/VD/$($VDLItem.ClientCode)/$($VDLItem.Rev)"
				$clientDocPath = "$($clientUrl)/$($clientRelDocPath)"

				# Creazione record CDL
				try {
					$CDLItem = Add-PnPListItem -List $CDL -Values @{
						Title                       = $VDLItem.TCM_DN
						ClientCode                  = $VDLItem.ClientCode
						IssueIndex                  = $VDLItem.Rev
						Index                       = $VDLItem.Index
						DocumentTitle               = $VDLItem.DocTitle
						ReasonForIssue              = $VDLItem.ReasonForIssue
						IsCurrent                   = $VDLItem.IsCurrent
						PONumber                    = $VDLItem.PONumber
						MRCode                      = $VDLItem.MRCode
						IDDocumentList              = $VDLItem.ID
						TransmittalPurpose          = $row.'TRANSMITTAL PURPOSE'
						LastTransmittal             = $row.'LAST TRANSMITTAL'
						LastTransmittalDate         = $lastTrnDateConv
						CommentDueDate              = $commentDueDateConv
						LastClientTransmittal       = $row.'LAST CLIENT TRANSMITTAL'
						LastClientTransmittalDate   = $lastClientTrnDateConv
						ApprovalResult              = $row.'APPROVAL RESULT'
						ActualDate                  = $actualDateConv
						DocumentsPath               = $clientDocPath
						ClientDepartmentCode        = $CDCMItem.ListPath
						ClientDepartmentDescription = $CDCMItem.Value
						DocumentTypology            = 'VD'
						DocumentClass               = $VDLItem.DocClass
						CommentRequest              = $true
						DD_SourceEnvironment        = 'VendorDocuments'
					} -Connection $clientConn
					$msg = "[SUCCESS] - List: $($CDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - CREATED"
				}
				catch { $msg = "[ERROR] - List: $($CDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - CREATION FAILED" }
				Write-Log $msg
				$pathSplit = $clientDocPath.Split('/')
			}
			else {
				# Se esiste, aggiorna il record sulla CDL
				try {
					Set-PnPListItem -List $CDL -Identity $CDLitem.ID -Values @{
						TransmittalPurpose        = $row.'TRANSMITTAL PURPOSE'
						LastTransmittal           = $row.'LAST TRANSMITTAL'
						LastTransmittalDate       = $lastTrnDateConv
						LastClientTransmittal     = $row.'LAST CLIENT TRANSMITTAL'
						LastClientTransmittalDate = $lastClientTrnDateConv
						ApprovalResult            = $row.'APPROVAL RESULT'
						ActualDate                = $actualDateConv
						CommentRequest            = $commentRequestCalc
					} -UpdateType SystemUpdate -Connection $clientConn | Out-Null
					$msg = "[SUCCESS] - List: $($CDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - UPDATED"
				}
				catch { $msg = "[ERROR] - List: $($CDL) - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - UPDATE FAILED" }
				Write-Log $msg
				$pathSplit = $CDLItem.DocPath.Split('/')
			}

			# Caricamento file in Client Area
			$rootPath = $pathSplit[5..$pathSplit.Length] -join '/'
			$originPath = "$($rootPath)/Originals"
			$attPath = "$($rootPath)/Attachments"

			# Creazione cartella Root e Attachment in Client Area
			Resolve-PnPFolder -SiteRelativePath $rootPath -Connection $clientConn | Out-Null
			Resolve-PnPFolder -SiteRelativePath $attPath -Connection $clientConn | Out-Null
			if (!(Test-Path -Path $lastTrnPath)) { mkdir -Path $lastTrnPath | Out-Null }
			$files = $null

			# Step 1: Caricamento file del transmittal in Root
			try {
				While (!$files) {
					# Ricerca file da caricare
					$files = Get-ChildItem -Path $lastTrnPath -ErrorAction Stop | Where-Object -FilterScript { $_.Name.Contains($row.'CLIENT CODE') -or $_.Name.Contains($row.'TCM CODE') }

					if ($null -eq $files) {
						# Se non li trova, li scarica dal Source in locale
						$trnFilesPath = ($vdmPathSplit[6..7] -join '/') + "/$($VDLItem.ClientTransmission)"
						$trnFolderDocs = Get-PnPFolderItem -FolderSiteRelativeUrl $trnFilesPath -ItemType File -Connection $subConn
						ForEach ($file in $trnFolderDocs) {
							try {
								Get-PnPFile -Url $($file.ServerRelativeUrl) -AsFile -Path "$($lastTrnPath)" -Filename $file.Name -Force -Connection $subConn | Out-Null
								$msg = "[SUCCESS] - List: $($vdmPathSplit[6]) - File: $($VDLItem.ClientTransmission)/$($file.Name) - DOWNLOADED"
							}
							catch {}
							Write-Log $msg
						}
					}
					else {
						# Caricamento file in Root Client Area
						ForEach ($file in $files) {
							try {
								$NewFileName = $File.Name -replace $($row.'TCM CODE'), $($row.'CLIENT CODE')

								Add-PnPFile -Path $file.FullName -Folder $rootPath -NewFileName $NewFileName -Values @{
									'IssueIndex'                  = $VDLItem.Rev
									'ReasonForIssue'              = $VDLItem.ReasonForIssue
									'ClientCode'                  = $VDLItem.ClientCode
									'IDDocumentList'              = $VDLItem.ID
									'IDClientDocumentList'        = $CDLItem.ID
									'CommentRequest'              = $true
									'IsCurrent'                   = $true
									'Transmittal_x0020_Number'    = $row.'LAST TRANSMITTAL'
									'TransmittalDate'             = $lastTrnDateConv
									'CommentDueDate'              = $commentDueDateConv
									'Index'                       = $VDLItem.Index
									'DocumentTitle'               = $VDLItem.DocTitle
									'DD_SourceEnvironment'        = 'VendorDocuments'
									'ClientDepartmentCode'        = $CDCMItem.ListPath
									'ClientDepartmentDescription' = $CDCMItem.Value
								} -Connection $clientConn | Out-Null
								$msg = "[SUCCESS] - List: $($pathSplit[5]) - File: Root/$($file.Name) - UPLOADED"
							}
							catch { $msg = "[ERROR] - List: $($pathSplit[5]) - File: Root/$($file.Name) - FAILED $($_)" }
							Write-Log $msg
						}
					}
				}
			}
			catch {
				Write-Host "[ERROR] - LocalFolder: $($lastTrnPath) - NOT FOUND - $($_)" -ForegroundColor Red
				Exit
			}

			# Controllo Last Client Transmittal
			if ($row.'LAST CLIENT TRANSMITTAL' -ne '') {

				# Crea la cartella Originals
				Resolve-PnPFolder -SiteRelativePath $originPath -Connection $clientConn | Out-Null

				# Step 2: Sposta dalla Root all'Originals
				$files = Get-PnPFolderItem -FolderSiteRelativeUrl $rootPath -ItemType File -Connection $clientConn | Select-Object Name
				foreach ($file in $files) {
					$relDocPath = "$($rootPath)/$($file.Name)"
					try {
						Move-PnPFile -SourceUrl $relDocPath -TargetUrl $originPath -Force -Connection $clientConn | Out-Null
						$msg = "[SUCCESS] - List: $($pathSplit[5]) - File: Originals/$($file.Name) - MOVED"
					}
					catch {
						$folderUrl = "$($SiteUrl)/$($rootPath)"
						Write-Host "[WARNING] Spostamento file '$($file.Name)' in Originals per documento '$($row.'TCM CODE')/$($row.'ISSUE INDEX')' non riuscito." -ForegroundColor Yellow
						Write-Host 'Path della ROOT aperto nel browser. Procedere manualmente.' -ForegroundColor Yellow
						Start-Process $folderUrl
						Pause
						$msg = "[SUCCESS] - List: $($pathSplit[5]) - File: Originals/$($file.Name) - MOVED"
					}
					Write-Log -Message $msg
				}

				# Step 3: Caricamento risposta in Root
				try {
					$files = Get-ChildItem -Path $lastClientTrnPath -ErrorAction Stop | Where-Object -FilterScript { $_.Name.Contains($VDLItem.ClientCode) -or $_.Name.Contains($VDLItem.TCM_DN) }

					if ($null -eq $files) { Write-Log "[WARNING] - LocalFolder: $($row.'LAST CLIENT TRANSMITTAL') - Doc: $($row.'TCM CODE')/$($row.'ISSUE INDEX') - EMPTY" }
					else {
						ForEach ($file in $files) {
							try {
								Add-PnPFile -Path $file.FullName -Folder $rootPath -NewFileName $NewFileName -Values @{
									'IssueIndex'                  = $VDLItem.Rev
									'ReasonForIssue'              = $VDLItem.ReasonForIssue
									'ClientCode'                  = $VDLItem.ClientCode
									'IDDocumentList'              = $VDLItem.ID
									'IDClientDocumentList'        = $CDLItem.ID
									'CommentRequest'              = $false
									'IsCurrent'                   = $VDLItem.IsCurrent
									'Transmittal_x0020_Number'    = $row.'LAST TRANSMITTAL'
									'TransmittalDate'             = $lastTrnDateConv
									'CommentDueDate'              = $commentDueDateConv
									'LastClientTransmittal'       = $row.'LAST CLIENT TRANSMITTAL'
									'LastClientTransmittalDate'   = $lastClientTrnDateConv
									'ApprovalResult'              = $row.'APPROVAL RESULT'
									'ActualDate'                  = $actualDateConv
									'Index'                       = $VDLItem.Index
									'DocumentTitle'               = $VDLItem.DocTitle
									'DD_SourceEnvironment'        = 'VendorDocuments'
									'ClientDepartmentCode'        = $CDCMItem.ListPath
									'ClientDepartmentDescription' = $CDCMItem.Value
								} -Connection $clientConn | Out-Null
								$msg = "[SUCCESS] - List: $($pathSplit[5]) - File: Root/$($file.Name) - UPLOADED"
							}
							catch { $msg = "[ERROR] - List: $($pathSplit[5]) - File: Root/$($file.Name) - FAILED" }
							Write-Log $msg
						}
					}
				}
				catch {
					Write-Host "[ERROR] - LocalFolder: $($lastClientTrnPath) - NOT FOUND - $($_)" -ForegroundColor Red
					Exit
				}
			}
		}
	}
	if ($csv.Length -gt 1) { Write-Progress -Activity 'Import' -Completed }
	Write-Log 'Operazione completata.'
}
catch { Throw }