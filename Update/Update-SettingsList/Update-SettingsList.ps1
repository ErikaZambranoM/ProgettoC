#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

<#
Questo Script va ad aggiungere una chiave SettingTitle alla lista Settings se non già esistente.

Per liste 'Settings' o 'MD Settings' creare CSV con:
SiteURL;Title;Value;Key;Remarks
(Remarks column is needed but can be empty)

Per lista 'Configuration List' creare CSV con:
SiteURL;Title;Value;Remarks
(Remarks column is needed but can be empty)

Per lista FTAEvolutionSettings:
SiteURL;PowerAutomateLink;NotifyTeamsPowerAutomateLink

TODO:
- Aggiungere verifica se il settaggio è presente più di una volta nel CSV
#>

[CmdletBinding(SupportsShouldProcess)]
Param(
	[Parameter(Mandatory = $true)]
	[ValidateSet('Settings', 'MD Settings', 'Configuration List', 'FTAEvolutionSettings')]
	[String]
	$SettingsListName,

	[Parameter(Mandatory = $true)][ValidateSet('ADD', 'UPDATE', 'BOTH', 'DELETE')]
	[String]
	$Mode,

	[Switch]
	$System
)

Function Connect-SPOSite
{
	<#
    .SYNOPSIS
        Connects to a SharePoint Online Site or Sub Site.

    .DESCRIPTION
        This function connects to a SharePoint Online Site or Sub Site and returns the connection object.
        If a connection to the specified Site already exists, the function returns the existing connection object.

    .PARAMETER SiteUrl
        Mandatory parameter. Specifies the URL of the SharePoint Online site or subsite.

    .EXAMPLE
        PS C:\> Connect-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/contoso"
        This example connects to the "https://contoso.sharepoint.com/sites/contoso" site.

    .OUTPUTS
        The function returns an object with the following properties:
            - SiteUrl: The URL of the SharePoint Online site or subsite.
            - Connection: The connection object to the SharePoint Online site or subsite as returned by the Connect-PnPOnline cmdlet.
#>

	Param(
		# SharePoint Online Site URL
		[Parameter(Mandatory = $true)]
		[ValidateScript({
				# Match a SharePoint Main Site or Sub Site URL
				If ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$')
				{
					$True
				}
				Else
				{
					Throw "`n'$($_)' is not a valid SharePoint Online site or subsite URL."
				}
			})]
		[String]
		$SiteUrl
	)

	Try
	{

		# Initialize Global:SPOConnections array if not already initialized
		If (-not $Script:SPOConnections)
		{
			$Script:SPOConnections = @()
		}
		Else
		{
			# Check if SPOConnection to specified Site already exists
			$SPOConnection = ($Script:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $SiteUrl }).Connection
		}

		# Create SPOConnection to specified Site if not already established
		If (-not $SPOConnection)
		{
			# Create SPOConnection to SiteURL
			Write-Host "Creating connection to '$($SiteUrl)'..." -ForegroundColor Cyan
			$SPOConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

			# Add SPOConnection to the list of connections
			$Script:SPOConnections += [PSCustomObject]@{
				SiteUrl    = $SiteUrl
				Connection = $SPOConnection
			}
		}
		else
		{
			Write-Host "Using existing connection to '$($SiteUrl)'..." -ForegroundColor Cyan
		}

		Return $SPOConnection
	}
	Catch
	{
		Throw
	}
}

Try
{
	$WhatIfPreference = $true
	$system ? ( $updateType = 'SystemUpdate' ) : ( $updateType = 'Update' ) | Out-Null
	$Mode = $Mode.ToUpper()

	# Caricamento ID / CSV / Documento
	$CSVPath = (Read-Host -Prompt 'CSV Path o SiteUrl').Trim('"')
	if ($CSVPath.ToLower().Contains('.csv'))
	{
		$csv = Import-Csv -Path $CSVPath -Delimiter ';'
		# Validazione colonne
		switch ($SettingsListName)
		{
			'FTAEvolutionSettings'
			{
				if ($Mode -ne 'UPDATE')
				{
					throw "FTAEvolutionSettings support only 'UPDATE' mode."
				}
				$RequiredColumns = @('SiteUrl', 'PowerAutomateLink', 'NotifyTeamsPowerAutomateLink')
				break
			}

			'Configuration List'
			{
				$RequiredColumns = @('SiteUrl', 'Title', 'Value', 'Remarks')
				break
			}

			# 'Settings' or 'MD Settings'
			Default
			{
				$RequiredColumns = @('SiteUrl', 'Title', 'Key', 'Value', 'Remarks')
			}
		}
		$validCounter = 0
		($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
			if ($_ -in $RequiredColumns) { $validCounter++ }
		}
		if ($validCounter -lt $RequiredColumns.Count)
		{
			Write-Host "Colonne obbligatorie mancanti: $($RequiredColumns -join ', ')" -ForegroundColor Red
			Exit
		}
	}
	else
	{
		if ($SettingsListName -ne 'FTAEvolutionSettings')
		{
			$title = Read-Host -Prompt 'Title'
			$key = Read-Host -Prompt 'Key'
			$value = Read-Host -Prompt 'Value'
			$remarks = Read-Host -Prompt 'Remarks'
			$csv = @(
				[PSCustomObject]@{
					SiteUrl = $CSVPath
					Title   = $title
					Key     = $key
					Value   = $value
					Remarks = $remarks
					Count   = 1
				}
			)
		}
		else
		{
			$PowerAutomateLink = Read-Host -Prompt 'PowerAutomateLink'
			$NotifyTeamsPowerAutomateLink = Read-Host -Prompt 'NotifyTeamsPowerAutomateLink'
			$csv = @(
				[PSCustomObject]@{
					SiteUrl                      = $CSVPath
					PowerAutomateLink            = $PowerAutomateLink
					NotifyTeamsPowerAutomateLink = $NotifyTeamsPowerAutomateLink
				}
			)
		}
	}

	Write-Host ''
	Start-Transcript -Path "$($PSScriptRoot)\Logs\$($WhatIfPreference ? 'WhatIf-' : $null)$((Get-Item -Path $MyInvocation.MyCommand.Path).BaseName)_$(Get-Date -Format 'dd-MM-yyyy_HH-mm-ss').log" -Force -IncludeInvocationHeader -WhatIf:$false
	Write-Host ''

	# Set CSV Log properties
	$CSVLogPath = "$($PSScriptRoot)\Logs\$($WhatIfPreference ? 'WhatIf-' : $null)$((Get-Item -Path $MyInvocation.MyCommand.Path).BaseName)_$(Get-Date -Format 'dd-MM-yyyy_HH-mm-ss').csv"
	[Array]$CSVOutputProperties = ($CSV[0].PSObject.Members | Where-Object -FilterScript { $_.MemberType -eq 'NoteProperty' }).Name


	$NewCsv = @()
	$rowCounter = 0
	$ItemIndex = $null
	$SettingsItems = @()
	$FTAEvolutionSettings_Items = $null

	if ($SettingsListName -eq 'FTAEvolutionSettings')
	{
		# Duplicate rows for each item in FTAEvolutionSettings_Items
		foreach ($row in $csv)
		{
			$SiteConnection = Connect-SPOSite -SiteUrl $row.SiteURL
			$FTAEvolutionSettings_Items = Get-PnPListItem -List $SettingsListName -Connection $SiteConnection -PageSize 5000 | ForEach-Object {
				[PSCustomObject] @{
					ID                           = $_['ID']
					FTAName                      = $_['Title']
					PowerAutomateLink            = $_['PowerAutomateLink']
					NotifyTeamsPowerAutomateLink = $_['NotifyTeamsPowerAutomateLink']
				}
			}

			if ($null -ne $FTAEvolutionSettings_Items)
			{
				for ($i = 0; $i -lt $($FTAEvolutionSettings_Items.Count); $i++)
				{
					$NewCsv += $row
				}

				$SettingsItems += $FTAEvolutionSettings_Items
			}
			else
			{
				Write-Host "`nNo FTA found on $($row.SiteUrl)/Lists/FTAEvolutionSettings`n" -ForegroundColor Yellow
			}
		}
		$csv = $NewCsv | Sort-Object -Property SiteURL
	}


	Write-Host "Inizio operazione...`n" -ForegroundColor Cyan
	foreach ($row in $csv)
	{
		$rowCounter++
		if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter)/$($csv.Count)" -PercentComplete ([Math]::Round($rowCounter / $csv.Count) * 100) }
		Write-Host "$rowCounter / $($csv.Count)" -ForegroundColor Blue
		Write-Host "Site: $($row.SiteUrl.Split('/')[-1])" -ForegroundColor Blue
		$SiteConnection = Connect-SPOSite -SiteUrl $row.SiteURL

		Switch ($SettingsListName)
		{
			'Settings'
			{
				if (-not $SettingsItems)
				{
					$SettingsItems = Get-PnPListItem -List $SettingsListName -Connection $SiteConnection -PageSize 5000 | ForEach-Object {
						[PSCustomObject] @{
							ID             = $_['ID']
							SettingTitle   = $_['Title']
							SettingValue   = $_['Value']
							Value          = $_['Value']
							SettingKey     = $_['Key']
							Key            = $_['Key']
							SettingRemarks = $_['Remarks']
							Remarks        = $_['Remarks']
						}
					}
				}

				# Filtra dalla lista settings l'elemento con il SettingsTitle indicato nel CSV
				$found = $SettingsItems | Where-Object { $_.SettingTitle -eq $row.Title }

				$SettingToImport = @{
					Value = $row.Value
					Key   = $row.Key
				}
				if ($row.Remarks)
				{
					$SettingToImport.Remarks = $row.Remarks
				}
				Break
			}

			'Configuration List'
			{
				if (-not $SettingsItems)
				{
					$SettingsItems = Get-PnPListItem -List $SettingsListName -Connection $SiteConnection -PageSize 5000 | ForEach-Object {
						[PSCustomObject] @{
							ID             = $_['ID']
							SettingTitle   = $_['Title']
							SettingValue   = $_['VD_ConfigValue']
							VD_ConfigValue = $_['VD_ConfigValue']
							SettingRemarks = $_['Remarks']
							Remarks        = $_['Remarks']
						}
					}
				}

				# Filtra dalla lista settings l'elemento con il SettingsTitle indicato nel CSV
				$found = $SettingsItems | Where-Object { $_.SettingTitle -eq $row.Title }

				$SettingToImport = @{
					VD_ConfigValue = $row.Value
				}
				if ($row.Remarks)
				{
					$SettingToImport.Remarks = $row.Remarks
				}
				Break
			}

			'MD Settings'
			{
				if (-not $SettingsItems)
				{
					$SettingsItems = Get-PnPListItem -List $SettingsListName -Connection $SiteConnection -PageSize 5000 | ForEach-Object {
						[PSCustomObject] @{
							ID           = $_['ID']
							SettingKey   = $_['DDMDKey']
							SettingValue = $_['DDMDValue']
							DDMDValue    = $_['DDMDValue']
							SettingNote  = $_['DDMDNote']
							DDMDNote     = $_['DDMDNote']
						}
					}
				}

				# Filtra dalla lista settings l'elemento con il SettingsTitle indicato nel CSV
				$found = $SettingsItems | Where-Object { $_.SettingKey -eq $row.Key }

				$SettingToImport = @{
					DDMDValue = $row.Value
				}
				if ($row.Remarks)
				{
					$SettingToImport.DDMDNote = $row.Remarks
				}
				Break
			}

			'FTAEvolutionSettings'
			{
				if ($null -eq $ItemIndex)
				{
					$ItemIndex = 0
					$found = $SettingsItems[$ItemIndex]
				}
				else
				{
					$ItemIndex++
					$found = $SettingsItems[$ItemIndex]
				}

				$SettingToImport = @{
					PowerAutomateLink            = $row.PowerAutomateLink
					NotifyTeamsPowerAutomateLink = $row.NotifyTeamsPowerAutomateLink
				}
				Break
			}

			Default
			{ Throw ("List '{0}' not supported!" -f $SettingsListName) }
		}

		try
		{
			# Create the CSV output object
			$CSVOutput = [PSCustomObject]@{}

			# Add the source CSV row to the CSV output
			$CSVOutputProperties | ForEach-Object {
				$CSVOutput | Add-Member -NotePropertyName $_ -NotePropertyValue $row.$($_)
			}

			# Add found item properties to the CSV output
			$SettingToImport.GetEnumerator() | ForEach-Object {
				$CSVOutput | Add-Member -NotePropertyName "OLD $($_.Name)" -NotePropertyValue $found.$($_.Name)
			}

			# Add the operation type to the CSV output
			$CSVOutput | Add-Member -NotePropertyName 'Operation' -NotePropertyValue $Mode

			# More than one item to update found, skip
			if ([Array]$found.Count -gt 1)
			{
				$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Skipped'
				$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue "Multiple items ($($found.Count)) found with the same key."
				Write-Host "[WARNING] - List: $($SettingsListName) - Key: $($row.Title ? $row.Title : $row.Key) - DUPLICATED" -ForegroundColor Yellow
			}

			# No item found
			if ($null -eq $found)
			{
				# Add the item
				if ($Mode -eq 'ADD' -or $Mode -eq 'BOTH')
				{
					if ($PSCmdlet.ShouldProcess($SettingsListName, "Add key '$($SettingToImport.DDMDKey ?? $SettingToImport.Title)' with value '$($SettingToImport.DDMDValue ?? $SettingToImport.VD_ConfigValue ?? $SettingToImport.Value)'"))
					{
						Add-PnPListItem -List $SettingsListName -Values $SettingToImport -Connection $SiteConnection | Out-Null
						$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Success'
						$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Added'
						Write-Host "[SUCCESS] - List: $($SettingsListName) - Key: $($row.Title ? $row.Title : $row.Key) - ADDED" -ForegroundColor Green
					}
					else
					{
						$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Success (WhatIf)'
						$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Added (WhatIf)'
					}
				}
				# Skip the item
				elseif ($Mode -eq 'UPDATE' -or $Mode -eq 'DELETE')
				{
					$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Skipped'
					$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Not found'
					Write-Host "[WARNING] - List: $($SettingsListName) - Key: $($row.Title ? $row.Title : $row.Key) - NOT FOUND" -ForegroundColor Yellow
				}
			}
			# One item found
			else
			{
				# Update the item
				if ($Mode -eq 'UPDATE' -or $mode -eq 'BOTH')
				{
					if ($PSCmdlet.ShouldProcess($SettingsListName, "Update key '$($SettingToImport.DDMDKey ?? $SettingToImport.Title ?? $found.FTAName)' with value '$($SettingToImport.DDMDValue ?? $SettingToImport.VD_ConfigValue ?? $SettingToImport.Value ?? $($SettingToImport | Format-List | Out-String).TrimEnd() + "`n")'"))
					{
						Set-PnPListItem -List $SettingsListName -Identity $found.ID -Values $SettingToImport -UpdateType $updateType -Connection $SiteConnection | Out-Null
						$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Success'
						$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Updated'
						Write-Host "[SUCCESS] - List: $($SettingsListName) - Key: $($row.Title ?? $row.Key ?? $($SettingToImport | Format-List | Out-String)) - UPDATED" -ForegroundColor Green
					}
					else
					{
						$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Success (WhatIf)'
						$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Updated (WhatIf)'
					}
				}
				# Skip the item (already exists)
				elseif ($Mode -eq 'ADD')
				{
					$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Skipped'
					$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Already exists'
					Write-Host "[WARNING] - List: $($SettingsListName) - Key: $($row.Title ? $row.Title : $row.Key) - ALREADY EXISTS" -ForegroundColor Yellow
				}
				# Delete the item
				elseif ($Mode -eq 'DELETE')
				{
					if ($PSCmdlet.ShouldProcess($SettingsListName, "Remove key '$($SettingToImport.DDMDKey ?? $SettingToImport.Title)'"))
					{
						Remove-PnPListItem -List $SettingsListName -Identity $found.ID -Force -Recycle -Connection $SiteConnection | Out-Null
						$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Success'
						$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Deleted'
						Write-Host "[SUCCESS] - List: $($SettingsListName) - Key: $($row.Title ? $row.Title : $row.Key) - DELETED" -ForegroundColor Green
					}
					else
					{
						$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Success (WhatIf)'
						$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue 'Deleted (WhatIf)'
					}
				}
			}
			Write-Host ''
		}
		catch
		{
			$CSVOutput | Add-Member -NotePropertyName 'Result' -NotePropertyValue 'Failed'
			$CSVOutput | Add-Member -NotePropertyName 'Result Details' -NotePropertyValue "$($_)"
			Write-Host "[ERROR] - List: $($SettingsListName) - Key: $($row.Title ? $row.Title : $row.Key) - $($_)" -ForegroundColor Red
		}
		finally
		{
			# Export the output to the log file
			$CSVOutput | Export-Csv -Path $CSVLogPath -Append -NoTypeInformation -Delimiter ';' -Encoding UTF8BOM -WhatIf:$false
		}
	}
}
catch { Throw }
finally
{
	if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Completed }
	Stop-Transcript -WhatIf:$false
	Write-Host ''
}