#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

<#
	AGGIORNAMENTO TCM DOCUMENT NUMER E CLIENT CODE VD o DD -> CD

	Provide CSV files with current values and values to be set:
		- TCM_DN
		- NewTCM_DN
		- NewCC
#>

<#
	ToDo:
		- Return errors in output and Log
		- Specify which area has being handled during script execution
		- Map client area to DD area with IDDocumentList (so that even if TCN_DN is differente they can be updated)
#>

param(
	[parameter(Mandatory = $true)][string]$SiteURL
)

#Funzione di log to CSV
function Write-Log
{
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $siteCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath))
	{
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog 'Timestamp;Type;ListName;TCM_DN;Rev;Action;ID;Errors'
	}
	$FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

	if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
	else
	{
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(' - List: ', ';').Replace(' - TCM_DN: ', ';').Replace(' - Rev: ', ';').Replace(' - Files: ', ';').Replace(' - FolderName: ', ';').Replace(' - ID: ', ';')
	Add-Content $logPath "$FormattedDate;$Message"
}

try
{
	# Caricamento CSV/Documento/Tutta la lista
	$CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
	if ($CSVPath.ToLower().Contains('.csv'))
	{
		$csv = Import-Csv $CSVPath -Delimiter ';'
		# Validazione colonne
		$validCols = @('TCM_DN', 'NewTCM_DN', 'NewCC')
		$validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
			if ($_ -in $validCols) { $validCounter++ }
		}
		if ($validCounter -lt $validCols.Count)
		{
			Write-Host "Colonne obbligatorie mancanti: $($validCols -join ', ')" -ForegroundColor Red
			Exit
		}
	}
	elseif ($CSVPath -ne '')
	{
		$newTCM_DN = Read-Host -Prompt 'Nuovo TCM Document Number'
		$newClientCode = Read-Host -Prompt 'Nuovo Client Code'
		$csv = [PSCustomObject] @{
			TCM_DN    = $CSVPath
			NewTCM_DN = $newTCM_DN
			NewCC     = $newClientCode
			Count     = 1
		}
	}

	# Variabile $DDLogic
	If ($SiteURL.ToLower().Contains('digitaldocuments') -or $SiteURL.ToLower().Contains('ddwave2')) { $DDLogic = $true }
	else { $DDLogic = $false }

	# Elenco liste coinvolte
	$DDList = 'DocumentList'
	$VDList = 'Vendor Documents List'
	$PFSList = 'Process Flow Status List'
	$CSReport = 'Comment Status Report'
	$TQDRegistry = 'TransmittalQueueDetails_Registry'
	$CDList = 'Client Document List'
	$CTQDRegistry = 'ClientTransmittalQueueDetails_Registry'

	# Connessione al sito principale (DD o VDM)
	$mainConn = Connect-PnPOnline -Url $SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
	If ($DDLogic) { $siteCode = $SiteURL.Split('/')[-1].ToLower() -replace 'digitaldocuments', '' }
	Else { $siteCode = $SiteURL.Split('/')[-1].ToLower() -replace 'vdm_', '' }

	$mainDLs = Get-PnPList -Connection $mainConn

	# Connessione al sito Client (DDc)
	$CDUrl = "https://tecnimont.sharepoint.com/sites/$($siteCode)DigitalDocumentsC"
	$clientConn = Connect-PnPOnline -Url $CDUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
	$clientDLs = Get-PnPList -Connection $clientConn

	if ($DDLogic)
	{
		$mainList = $DDList
		# Se DD, scarica la DocumentList
		Write-Log "Caricamento '$($mainList)'..."
		$mainItems = Get-PnPListItem -List $mainList -PageSize 5000 -Connection $mainConn | ForEach-Object {
			[PSCustomObject]@{
				ID         = $_['ID']
				TCM_DN     = $_['Title']
				Rev        = $_['IssueIndex']
				ClientCode = $_['ClientCode']
				LastTrn    = $_['LastTransmittal']
				DocPath    = $_['DocumentsPath']
			}
		}
	}
	else
	{
		$mainList = $PFSList
		# Se VDM, scarica la Process Flow Status
		Write-Log "Caricamento '$($mainList)'..."
		$mainItems = Get-PnPListItem -List $mainList -PageSize 5000 -Connection $mainConn | ForEach-Object {
			[PSCustomObject]@{
				ID          = $_['ID']
				TCM_DN      = $_['VD_DocumentNumber']
				Rev         = $_['VD_RevisionNumber']
				ClientCode  = $_['VD_ClientDocumentNumber']
				LastTrn     = $null
				DocPath     = $null
				VDL_ID      = $_['VD_VDL_ID']
				Comments_ID = $_['VD_CommentsStatusReportID']
			}
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Caricamento lista Transmittal Queue Details Registry
	Write-Log "Caricamento '$($TQDRegistry)'..."
	$TQDItems = Get-PnPListItem -List $TQDRegistry -PageSize 5000 -Connection $mainConn | ForEach-Object {
		[PSCustomObject]@{
			ID         = $_['ID']
			TCM_DN     = $_['Title']
			Rev        = $_['IssueIndex']
			ClientCode = $_['ClientCode']
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Scarica la Client Document List
	Write-Log "Caricamento '$($CDList)'..."
	$CDItems = Get-PnPListItem -List $CDList -PageSize 5000 -Connection $clientConn | ForEach-Object {
		[PSCustomObject]@{
			ID             = $_['ID']
			TCM_DN         = $_['Title']
			Rev            = $_['IssueIndex']
			ClientCode     = $_['ClientCode']
			ClientDeptCode = $_['ClientDepartmentCode']
			DocPath        = $_['DocumentsPath']
			LastClientTrn  = $_['LastClientTransmittal']
		}
	}
	Write-Log 'Caricamento lista completato.'

	# Scarica la Client Transmittal Details Registry
	Write-Log "Caricamento '$($CTQDRegistry)'..."
	$CTQDItems = Get-PnPListItem -List $CTQDRegistry -PageSize 5000 -Connection $clientConn | ForEach-Object {
		[PSCustomObject]@{
			ID         = $_['ID']
			TCM_DN     = $_['Title']
			TrnID      = $_['TransmittalID']
			OriginPath = $_['OriginPath']
		}
	}
	Write-Log 'Caricamento lista completato.'

	$rowCounter = 0
	Write-Log 'Inizio operazione...'
	ForEach ($row in $csv)
	{
		if ($csv.Count -gt 1) { Write-Progress -Activity 'Document' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

		# Verifica nuovo TCM_DN
		if ($row.NewTCM_DN -eq '') { $row.NewTCM_DN = $row.TCM_DN }
		else
		{
			$found = $mainItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.NewTCM_DN }
			if ($null -ne $found)
			{
				Write-Host "[WARNING] - List: $($mainList) - TCM_DN: $($row.NewTCM_DN) - ALREADY EXIST" -ForegroundColor DarkYellow
				Continue
			}
		}

		# Filtro revisioni sulla DocumentList (DD) / Process Flow Status (VDM)
		[Array]$revs = $mainItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN }
		if ($null -eq $revs)
		{
			Write-Log "[ERROR] - List: $($mainList) - TCM_DN: $($row.TCM_DN) - NOT FOUND"
			Continue
		}

		# Calcolo e creazione nuova posizione del documento lato TCM
		$newTCM_DNSplit = $row.NewTCM_DN.Split('-')
		$newDocFolder = "$($newTCM_DNSplit[1][0])/$($newTCM_DNSplit[1][1])/$($newTCM_DNSplit[2])/$($row.NewTCM_DN)"

		# Per ogni revisione del documento
		$revCounter = 0
		ForEach ($rev in $revs)
		{
			$revCounter++
			if ($null -eq $rev.Rev) { $rev.Rev = 'TBD' }
			# Verifica nuovo Client Code
			if ($row.NewCC -eq '') { $row.NewCC = $rev.ClientCode }
			else
			{
				$found = $mainItems | Where-Object -FilterScript { $_.ClientCode -eq $row.NewCC }
				if ($null -ne $found)
				{
					Write-Host "[WARNING] - List: $($mainList) - ClientCode: $($row.NewCC) - ALREADY EXIST" -ForegroundColor DarkYellow
					Continue
				}
			}

			Write-Host "Document: $($row.TCM_DN) - Rev: $($rev.Rev) ($($revCounter)/$($revs.Count))" -ForegroundColor Blue

			# Logica DD
			if ($DDLogic)
			{
				# Calcolo percorso corretto
				Resolve-PnPFolder -SiteRelativePath $newDocFolder -Connection $mainConn | Out-Null
				$pathSplit = $rev.DocPath.Split('/')
				$siteRelPath = ($pathSplit[5..$pathSplit.Length] -join '/')
				$DLRelPath = '/' + ($pathSplit[3..5] -join '/')
				$DLRelPath_Calc = "/$($pathSplit[3])/$($pathSplit[4])/$($newTCM_DNSplit[1][0])"
				$DL = $mainDLs | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }
				$DL_Calc = $mainDLs | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath_Calc }
				$DocPath_Calc = ($pathSplit[0..4] -join '/') + "/$($newDocFolder)/$($rev.Rev)"

				# Aggiorna record sulla DocumentList
				try
				{
					Set-PnPListItem -List $mainList -Identity $rev.ID -Values @{
						'Title'                  = $row.NewTCM_DN
						'ClientCode'             = $row.NewCC
						'DocumentsPath'          = $DocPath_Calc
						'DepartmentCode'         = $($newTCM_DNSplit[1][0])
						'DocumentClassification' = $($newTCM_DNSplit[1])
						'DocumentTypology'       = $($newTCM_DNSplit[2])
					} -Connection $mainConn | Out-Null
					Write-Log "[SUCCESS] - List: $($mainList) - TCM_DN: $($row.NewTCM_DN) - Rev: $($rev.Rev) - UPDATED - ID: $($rev.ID)"
				}
				catch { Write-Log "[ERROR] - List: $($mainList) - TCM_DN: $($row.TCM_DN) Rev: $($rev.Rev) - FAILED - ID: $($rev.ID)" }

				# Ottiene file e cartelle del DocumentsPath
				$fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $siteRelPath -ItemType File -Recursive -Connection $mainConn | Select-Object Name, ServerRelativeUrl
				$folderList = Get-PnPFolderItem -FolderSiteRelativeUrl $siteRelPath -ItemType Folder -Connection $mainConn | Select-Object Name, ServerRelativeUrl

				if ($null -eq $fileList) { Write-Log "[WARNING] - List: $($DL.Title) - $($rev.TCM_DN) - $($rev.Rev) - EMPTY OR NOT FOUND" }
				else
				{
					# Aggiorna ogni singolo file del DocumentsPath
					foreach ($file in $fileList)
					{
						$fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem -Connection $mainConn
						$newFileName = $file.Name -replace $rev.TCM_DN, $row.NewTCM_DN

						If ($null -ne $fileAtt.FieldValues.DocumentNumber)
						{
							try
							{
								Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
									'FileLeafRef'    = $newFileName
									'DocumentNumber' = $row.NewTCM_DN
									'ClientCode'     = $row.NewCC
								} -Connection $mainConn | Out-Null
								Write-Log "[SUCCESS] - List: $($DL.Title) - FileName: $($newFileName) - Rev: $($rev.Rev) - UPDATED"
							}
							catch { Write-Log "[ERROR] - List: $($DL.Title) - FileName: $($file.Name) - Rev: $($rev.Rev) - FAILED" }
						}
					}

					# Aggiorna ogni nome delle sottocartelle
					ForEach ($folder in $folderList)
					{
						$folderPathSplit = $folder.ServerRelativeUrl.Split('/')
						$folderRelPath = $folderPathSplit[3..$folderPathSplit.Length] -join '/'
						$newFolderName = $folder.Name -replace $rev.TCM_DN, $row.NewTCM_DN
						try
						{
							Rename-PnPFolder -Folder $folderRelPath -TargetFolderName $newFolderName -Connection $mainConn | Out-Null
							Write-Log "[SUCCESS] - List: $($DL.Title) - FolderName: $($newFolderName) - Rev: $($rev.Rev) - UPDATED"
						}
						catch { Write-Log "[ERROR] - List: $($DL.Title) - FolderName: $($folder.Name) - Rev: $($rev.Rev) - FAILED" }
					}
				}

				# Aggiornamento posizione documento lato TCM
				if ($rev.TCM_DN -ne $row.NewTCM_DN)
				{
					try
					{
						Move-PnPFolder -Folder $siteRelPath -TargetFolder $newDocFolder -Connection $mainConn | Out-Null
						Write-Log "[SUCCESS] - List: $($DL_Calc.Title) - TCM_DN: $($row.NewTCM_DN) - Folder: $($rev.Rev) - MOVED"
					}
					catch { Write-Log "[ERROR] - List: $($DL.Title) - TCM_DN: $($rev.TCM_DN) - Folder: $($rev.Rev) - FAILED" }
				}
			}
			# Logica VDM
			else
			{
				# Calcolo variabili
				$rev.LastTrn = [string]((Get-PnPListItem -List $VDList -Id $rev.VDL_ID -Fields 'LastTransmittal' -Connection $mainConn).FieldValues).Values
				$rev.DocPath = [string]((Get-PnPListItem -List $VDList -Id $rev.VDL_ID -Fields 'VD_DocumentPath' -Connection $mainConn).FieldValues).Values
				$pathSplit = $rev.DocPath.Split('/')
				$subSite = ($pathSplit[0..5]) -join '/'
				$siteRelPath = ($pathSplit[6..$pathSplit.Length] -join '/')
				$DocPath_Calc = $rev.DocPath -replace $rev.TCM_DN, $row.NewTCM_DN

				# Aggiornamento Process Flow Status List
				try
				{
					Set-PnPListItem -List $mainList -Identity $rev.ID -Values @{
						'VD_DocumentNumber'       = $row.NewTCM_DN
						'VD_ClientDocumentNumber' = $row.NewCC
					} -Connection $mainConn | Out-Null
					Write-Log "[SUCCESS] - List: $($mainList) - TCM_DN: $($row.NewTCM_DN) - Rev: $($rev.Rev) - UPDATED - ID: $($rev.ID)"
				}
				catch { Write-Log "[ERROR] - List: $($mainList) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - FAILED - ID: $($rev.ID)" }

				# Aggiornamento Comments Status Report
				if ($null -ne $rev.Comments_ID)
				{
					try
					{
						Set-PnPListItem -List $CSReport -Identity $rev.Comments_ID -Values @{
							'VD_DocumentNumber' = $row.NewTCM_DN
						} -Connection $mainConn | Out-Null
						Write-Log "[SUCCESS] - List: $($CSReport) - TCM_DN: $($row.NewTCM_DN) - Rev: $($rev.Rev) - UPDATED - ID: $($rev.Comments_ID)"
					}
					catch { Write-Log "[ERROR] - List: $($CSReport) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - FAILED - ID: $($rev.Comments_ID)" }
				}

				# Aggiornamento Vendor Documents List
				try
				{
					Set-PnPListItem -List $VDList -Identity $rev.VDL_ID -Values @{
						'VD_DocumentNumber'       = $row.NewTCM_DN
						'VD_ClientDocumentNumber' = $row.NewCC
						'VD_DocumentPath'         = $DocPath_Calc
					} -Connection $mainConn | Out-Null
					Write-Log "[SUCCESS] - List: $($VDList) - TCM_DN: $($row.NewTCM_DN) - Rev: $($rev.Rev) - UPDATED - ID: $($rev.VDL_ID)"
				}
				catch { Write-Log "[ERROR] - List: $($VDList) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - FAILED - ID: $($rev.VDL_ID)" }

				# Connessione al sottosito del Vendor
				$subConn = Connect-PnPOnline -Url $subSite -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

				# Scarica la Document Library del PO
				Write-Log "Caricamento '$($pathSplit[6])'..."
				$POItems = Get-PnPListItem -List $pathSplit[6] -PageSize 5000 -Connection $subConn | ForEach-Object {
					[PSCustomObject]@{
						ID         = $_['ID']
						Name       = $_['FileLeafRef']
						TCM_DN     = $_['VD_DocumentNumber']
						Rev        = $_['VD_RevisionNumber']
						PFS_ID     = $_['VD_AGCCProcessFlowItemID']
						ItemPath   = $_['FileRef']
						ObjectType = $_['FSObjType'] #1 folder, 0 file.
					}
				}

				# Filtra la cartella del documento
				$docFolder = $POItems | Where-Object -FilterScript { ($_.TCM_DN -eq $rev.TCM_DN -or $_.TCM_DN -eq $row.NewTCM_DN) -and $_.PFS_ID -eq $rev.ID }

				If ($null -eq $docFolder) { Write-Log "[ERROR] - List: $($pathSplit[6]) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - NOT FOUND" }
				else
				{
					# Ottiene i file del DocumentsPath
					$fileList = $POItems | Where-Object -FilterScript { $_.ObjectType -eq 0 -and $_.ItemPath.Contains($docFolder.ItemPath) }

					# Aggiornamento file sottosito PO
					if ($null -eq $fileList) { Write-Host "[WARNING] - List: $($pathSplit[6]) - Folder: $($rev.TCM_DN)/$($rev.Rev) - EMPTY" -ForegroundColor Yellow }
					else
					{
						foreach ($file in $fileList)
						{
							$newFileName = $file.Name -replace $rev.TCM_DN, $row.NewTCM_DN
							try
							{
								#Rename-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl -TargetFileName $newFileName | Out-Null
								Set-PnPListItem -List $pathSplit[6] -Identity $file.ID -Values @{
									'FileLeafRef' = $newFileName
								} -UpdateType SystemUpdate -Connection $subConn | Out-Null
								Write-Log "[SUCCESS] - List: $($pathSplit[6]) - FileName: $($file.Name) - Rev: $($rev.Rev) - UPDATED - ID: $($file.ID)"
							}
							catch { Write-Log "[ERROR] - List: $($pathSplit[6]) - FileName: $($file.Name) - Rev: $($rev.Rev) - FAILED - ID: $($file.ID)" }
						}
					}

					# Aggiornamento cartella sottosito PO
					if ($rev.TCM_DN -ne $row.NewTCM_DN)
					{
						# Calcolo nuovo nome cartella
						$newFolderName = $docFolder.Name -replace $rev.TCM_DN, $row.NewTCM_DN
						try
						{
							Set-PnPListItem -List $pathSplit[6] -Identity $docFolder.ID -Values @{
								'FileLeafRef'       = $newFolderName
								'VD_DocumentNumber' = $row.NewTCM_DN
							} -UpdateType SystemUpdate -Connection $subConn | Out-Null
							Write-Log "[SUCCESS] - List: $($pathSplit[6]) - FolderName: $($newFolderName) - Rev: $($rev.Rev) - UPDATED - ID: $($docFolder.ID)"
						}
						catch { Write-Log "[ERROR] - List: $($pathSplit[6]) - FolderName: $($docFolder.Name) - Rev: $($rev.Rev) - FAILED - ID: $($docFolder.ID)" }
					}
				}
			}

			# Controllo presenza Last Transmittal
			if ($null -ne $rev.LastTrn -and $rev.LastTrn -ne '')
			{
				# Filtro documento su Transmittal Queue Details Registry
				$TQDItem = $TQDItems | Where-Object -FilterScript { $_.TCM_DN -eq $rev.TCM_DN -and $_.Rev -eq $rev.Rev }

				# Aggiornamento Transmittal Queue Details Registry
				if ($null -eq $TQDItem) { Write-Log "[ERROR] - List: $($TQDRegistry) - TCM_DN: $($rev.TCM_DN) - $($rev.Rev) - NOT FOUND" }
				else
				{
					try
					{
						Set-PnPListItem -List $TQDRegistry -Identity $TQDItem.ID -Values @{
							'Title'      = $row.NewTCM_DN
							'ClientCode' = $row.NewCC
						} -Connection $mainConn | Out-Null
						Write-Log "[SUCCESS] - List: $($TQDRegistry) - TCM_DN: $($row.NewTCM_DN) - Rev: $($rev.Rev) - UPDATED - ID: $($TQDItem.ID)"
					}
					catch { Write-Log "[ERROR] - List: $($TQDRegistry) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - FAILED - ID: $($TQDItem.ID)" }
				}

				# Filtra sulla Client Document List
				$CDItem = $CDItems | Where-Object -FilterScript { $_.TCM_DN -eq $rev.TCM_DN -and $_.Rev -eq $rev.Rev }

				if ($null -eq $CDItem) { Write-Log "[ERROR] - List: $($CDList) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - NOT FOUND" }
				else
				{
					$pathSplit = $CDItem.DocPath.Split('/')
					$siteRelPath = ($pathSplit[5..$pathSplit.Length] -join '/')
					$newClientDocFolder = "$($pathSplit[5..7] -join '/')/$($row.NewCC)"
					$DocPath_Calc = $CDItem.DocPath -replace $siteRelPath, "$($newClientDocFolder)/$($CDItem.Rev)"
					$DLRelPath = '/' + ($pathSplit[3..5] -join '/')
					$DL = $clientDLs | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }

					# Crea la nuova cartella del documento col nuovo Client Code
					Resolve-PnPFolder -SiteRelativePath $newClientDocFolder -Connection $clientConn | Out-Null

					# Aggiornamento record Client Document List
					try
					{
						Set-PnPListItem -List $CDList -Identity $CDItem.ID -Values @{
							'Title'         = $row.NewTCM_DN
							'ClientCode'    = $row.NewCC
							'DocumentsPath' = $DocPath_Calc
						} -Connection $clientConn | Out-Null
						Write-Log "[SUCCESS] - List: $($CDList) - TCM_DN: $($row.NewTCM_DN) - Rev: $($CDItem.Rev) - UPDATED - ID: $($CDItem.ID)"
					}
					catch { Write-Log "[ERROR] - List: $($CDList) - TCM_DN: $($rev.TCM_DN) - Rev: $($CDItem.Rev) - FAILED - ID: $($CDItem.ID)" }

					# Ottiene i file del DocumentsPath
					$fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $siteRelPath -ItemType File -Recursive -Connection $clientConn | Select-Object Name, ServerRelativeUrl

					# Aggiornamento file Client
					if ($null -eq $fileList) { Write-Log "[WARNING] - List: $($DL.Title) - $($rev.TCM_DN) - $($rev.Rev) - EMPTY OR NOT FOUND" }
					else
					{
						foreach ($file in $fileList)
						{
							$fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem -Connection $clientConn
							$newFileName = $file.Name -replace $CDItem.ClientCode, $row.NewCC
							If ($null -ne $fileAtt.FieldValues.ClientCode)
							{
								try
								{
									Set-PnPListItem -List $DL.Title -Identity $fileAtt.Id -Values @{
										'FileLeafRef' = $newFileName
										'ClientCode'  = $row.NewCC
									} -UpdateType SystemUpdate -Connection $clientConn | Out-Null
									Write-Log "[SUCCESS] - List: $($DL.Title) - FileName: $($newFileName) - Rev: $($rev.Rev) - UPDATED - ID: $($fileAtt.Id)"
								}
								catch { Write-Log "[ERROR] - List: $($DL.Title) - FileName: $($file.Name) - Rev: $($rev.Rev) - FAILED - ID: $($fileAtt.Id)" }
							}
						}
					}

					# Aggiornamento posizione documento lato Client
					if ($CDItem.ClientCode -ne $row.NewCC)
					{
						try
						{
							Move-PnPFolder -Folder $siteRelPath -TargetFolder $newClientDocFolder -Connection $clientConn | Out-Null
							Write-Log "[SUCCESS] - List: $($DL.Title) - TCM_DN: $($row.NewTCM_DN) - Folder: $($rev.Rev) - MOVED"
						}
						catch { Write-Log "[ERROR] - List: $($DL.Title) - TCM_DN: $($rev.TCM_DN) - Folder: $($rev.Rev) - FAILED" }
					}

					# Aggiornamento Client Transmittal Queue Details Registry
					if ($null -ne $CDItem.LastClientTrn -and $CDItem.LastClientTrn -ne '')
					{
						$CTQDItem = $CTQDItems | Where-Object -FilterScript { $_.TCM_DN -eq $rev.TCM_DN -and $_.TrnID -eq $CDItem.LastClientTrn }

						if ($null -eq $CTQDItem) { Write-Log "[ERROR] - List: $($CTQDRegistry) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - NOT FOUND" }
						else
						{
							try
							{
								Set-PnPListItem -List $CTQDRegistry -Identity $CTQDItem.ID -Values @{
									'Title'      = $row.NewTCM_DN
									'IssueIndex' = $rev.Rev
									'ClientCode' = $row.NewCC
									'OriginPath' = $DocPath_Calc
								} -Connection $clientConn | Out-Null
								Write-Log "[SUCCESS] - List: $($CTQDRegistry) - TCM_DN: $($row.NewTCM_DN) - Rev: $($rev.Rev) - UPDATED - ID: $($CTQDItem.ID)"
							}
							catch { Write-Log "[ERROR] - List: $($CTQDRegistry) - TCM_DN: $($rev.TCM_DN) - Rev: $($rev.Rev) - FAILED  - ID: $($CTQDItem.ID)" }
						}
					}
				}
			}
		}
		Write-Host ''
	}
	Write-Progress -Activity 'Document' -Completed
	Write-Log 'Operazione completata.'
}
catch { Throw }