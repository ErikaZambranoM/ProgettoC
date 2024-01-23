<#
	Il CSV necessita solo la colonna 'Site' con gli URL dei siti da aggiornare.
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

# Funzione di log to CSV
Function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog "Timestamp;Type;Site;App Name;Version;Action"
	}
	$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

	if ($Message.Contains("[SUCCESS]")) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains("[ERROR]")) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains("[WARNING]")) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(" - Site: ", ";").Replace(" - App: ", ";").Replace(" - Package: ", ";").Replace(" - Version: ", ";").Replace(" - ", ";")
	Add-Content $logPath "$FormattedDate;$Message"
}

Try {
	# Required assembly to get content of sppkg file
	Add-Type -Assembly "System.IO.Compression.FileSystem"

	# Remove quotes from paths
	$Packages = (Read-Host -Prompt "Package Folder/Path").Trim('"')

	# Check if Packages is a folder, filter all .sppkg inside it
	if (Test-Path $Packages -PathType Container) { $PackagesToUpdate = Get-ChildItem -Path $Packages -Filter "*.sppkg" -File }
	elseif (Test-Path $Packages -PathType Leaf) {
		[Array]$PackagesToUpdate = [PSCustomObject]@{
			FullName = $Packages
			Count = 1
		}
	}
	else { [Array]$PackagesToUpdate = @() }

	If ($PackagesToUpdate.Count -eq 0) {
		Write-Host "Unable to find package(s)." -ForegroundColor Red
		Exit
	}

	#$METUrl = "https://tecnimont.sharepoint.com/sites/METDigitalDocuments"
	#$mainSite = Connect-PnPOnline -Url $METUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection

	# Caricamento CSV o sito singolo
	$CSVPath = (Read-Host -Prompt "CSV Path o Site Url").Trim('"')
	if ($CSVPath.ToLower().Contains(".csv")) { $csv = Import-Csv -Path $CSVPath -Delimiter ";" }
	elseif ($CSVPath -ne "") {
		$csv = [PSCustomObject]@{
			Site  = $CSVPath
			Count = 1
		}
	}
	else { Exit }

	$newUpload = Read-Host -Prompt "New installation? (True/False)"

	$rowCounter = 0
	Write-Log "Inizio operazione..."
	Foreach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity "Rilascio" -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }
		Write-Host "Sito: $($row.Site.Split("/")[-1])" -ForegroundColor Blue
		Connect-PnPOnline -Url $row.Site -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

		# Scarica lista delle App installate sul sito
		$InstalledApps = Get-PnPApp -Scope Site

		# Loop through all sppkg packages
		ForEach ($Package in $PackagesToUpdate) {
			# Treat the package as a zip file
			$SPPKG_AsZip = [IO.Compression.ZipFile]::OpenRead("$($Package.FullName)")

			# Filter the content of the sppkg package to find the AppManifest.xml file
			$ManifestFileObject = $SPPKG_AsZip.Entries | Where-Object { $_.Name -eq "AppManifest.xml" }

			# Open the AppManifest.xml file in ram  and read the content as XML
			$FileStream = $ManifestFileObject.Open()
			$FileReader = [System.IO.StreamReader]$FileStream
			[XML]$ManifestFileText = $FileReader.ReadToEnd()

			# Create an XmlNamespaceManager and add the namespace from your document
			$Namespaces = New-Object System.Xml.XmlNamespaceManager($ManifestFileText.NameTable)
			$Namespaces.AddNamespace("AppNameSpace", "http://schemas.microsoft.com/sharepoint/2012/app/manifest")

			# Select the App element and get the Name attribute to obtain the App DisplayName as showed  on SharePoint Online
			$AppDisplayName = $ManifestFileText.SelectSingleNode("//AppNameSpace:App", $Namespaces).Name
			$AppVersion = $ManifestFileText.SelectSingleNode("//AppNameSpace:App", $Namespaces).Version

			$FileReader.Close()
			$FileStream.Close()
			$SPPKG_AsZip.Dispose()

			# Filtra l'app da aggiornare tramite il nome ($AppDisplayName)
			$found = $InstalledApps | Where-Object -FilterScript { $_.Title -eq $AppDisplayName }
			if ($null -eq $found -and [System.Convert]::ToBoolean($newUpload) -eq $true) {
				try {
					Write-Log "App $($AppDisplayName) non trovata.`nUpload in corso..."
					$found = Add-PnPApp -Path $($Package.FullName) -Scope Site -Publish
					Write-Log = "[SUCCESS] - Site: $($row.Site) - Package: $($AppDisplayName) - Version: $($AppVersion) - UPLOADED on AppCatalog"
					Write-Log "Installazione app in corso..."
					Install-PnPApp -Scope Site -Identity $found.Id | Out-Null
					Write-Log "[SUCCESS] - Site: $($row.Site) - App: $($AppDisplayName) - INSTALLED"
				}
				catch {
					Write-Log "[ERROR] - Site: $($row.Site) - App: $($AppDisplayName) - FAILED"
					Throw
				}
				finally { Write-Host '' }
			}
			elseif ($null -eq $found) { Write-Log "[WARNING] - Site: $($row.Site) - App: $($AppDisplayName) - MISSING" }
			else {
				try {
					Write-Log "Controllo versione package '$($AppDisplayName)' in AppCatalog..."
					If ($found.AppCatalogVersion.ToString() -ne $AppVersion) {
						Write-Log "Trovata versione $($found.AppCatalogVersion.ToString()).`nUpload in corso della nuova versione $AppVersion..."
						Add-PnPApp -Path $($Package.FullName) -Scope Site -Overwrite -Publish | Out-Null
						$msg = "[SUCCESS] - Site: $($row.Site) - Package: $($AppDisplayName) - Version: $($AppVersion) - UPLOADED on AppCatalog"
					}
					Else {
						$msg = "[WARNING] - Site: $($row.Site) - Package: $($AppDisplayName) - ALREADY same version on AppCatalog ($($AppVersion))"
					}
					Write-Log $msg

					Write-Log "Controllo versione App '$($AppDisplayName)' installata..."
					If ($found.InstalledVersion) {
						If ($found.InstalledVersion.ToString() -ne $AppVersion) {
							Write-Log "Trovata versione $($found.AppCatalogVersion.ToString()).`nUpdate in corso alla nuova versione $AppVersion..."
							Update-PnPApp -Scope Site -Identity $($found.Id) | Out-Null
							$msg = "[SUCCESS] - Site: $($row.Site) - App: $($AppDisplayName) - UPDATED to $($AppVersion)"
						}
						Else {
							$msg = "[WARNING] - Site: $($row.Site) - App: $($AppDisplayName) - ALREADY UPDATED to $($AppVersion)"
						}

					}
					Else {
						$msg = "[WARNING] - Site: $($row.Site) - App: $($AppDisplayName) - NOT INSTALLED"
					}
					Write-Log $msg
				}
				catch {
					Write-Log "[ERROR] - Site: $($row.Site) - App: $($AppDisplayName) - FAILED"
					Throw
				}
				finally { Write-Host '' }
			}
		}
		Start-Sleep -Seconds 1
	}
	if ($csv.Count -gt 1) { Write-Progress -Activity "Rilascio" -Completed }
	Write-Log "Operazione completata."
}
Catch {
	Write-Progress -Activity "Rilascio" -Completed
	Throw
}