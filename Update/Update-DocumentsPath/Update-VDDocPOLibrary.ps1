<#
	This script will move a Document folder from one PO Library to another within the same subsite.
	As input, it accepts a single 'TCM Document Number' and the destination PO Number or a CSV file with 'TCM_DN' and 'DestinationPONumber' columns.
	Only Documents with no Last Transmittal will be moved.
#>

# Parameters
Param (
	[Parameter(Mandatory = $true)]
	[ValidateScript( # Validate that the input is a valid SharePoint MainSite URL
		{
			If ($_ -notmatch '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+/?$') {
				Throw ("{0}Error:'{1}' is not a valid SharePoint MainSite URL." -f "`n" , $_)
			}
			Else {
				$true
			}
		}
	)]
	[String]
	$VDMSiteURL
)

Function Connect-SPOSite {
	Param(
		# SharePoint Online Site URL
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[String]
		$Url
	)

	Try {
		# Create SPOConnection to specified Site if not already established
		$SPOConnection = ($Global:SPOConnections | Where-Object -FilterScript { $_.SiteUrl -eq $Url }).Connection
		If (-not $SPOConnection) {
			# Create SPOConnection to SiteURL provided in the PATicket
			$SPOConnection = Connect-PnPOnline -Url $Url -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

			# Add SPOConnection to the list of connections
			$Global:SPOConnections += [PSCustomObject]@{
				SiteUrl    = $Url
				Connection = $SPOConnection
			}
		}

		Return $SPOConnection
	}
	Catch {
		Throw
	}
}

Try {
	# Prompt for target single Document or CSV file, then validate input
	Do {
		$TargetDocuments = Read-Host -Prompt 'TCM Document Number or CSV Path'
		If ($TargetDocuments -eq '') {
			Write-Host ('Error: Input cannot be empty.{0}' -f "`n") -ForegroundColor Red
		}
	}
	While
	(
		# Validate that the input is not empty
		$TargetDocuments -eq ''
	)
	# If input is not a CSV file, ask for Destination PO Number
	$TargetDocuments = $TargetDocuments -replace '"', ''
	If (!($TargetDocuments.EndsWith('.csv'))) {
		Do {
			$DestinationPONumber = Read-Host -Prompt 'Destination PO Number'
			If ($DestinationPONumber -notmatch '^\d+$') {
				Write-Host ("Error:'{0}' is not a valid VDM PO Number.{1}" -f $DestinationPONumber, "`n") -ForegroundColor Red
			}
			Else {
				$TargetDocuments = [PSCustomObject] @{
					TCM_DN              = $TargetDocuments
					DestinationPONumber = $DestinationPONumber
				}
			}
		}
		While
		(
			# Validate that the input is a number
			$DestinationPONumber -notmatch '^\d+$'
		)
	}
	Else
 { # Import CSV file and check that it contains the required columns and at least 1 row
		$TargetDocuments = Import-Csv -Path $TargetDocuments -Delimiter ';' | Select-Object -Unique -Property TCM_DN, DestinationPONumber
		If ($TargetDocuments) {
			$CSVColumns = ($TargetDocuments | Get-Member -MemberType NoteProperty).Name
			$MissingColumns = @()
			ForEach ( $Column in @('TCM_DN', 'DestinationPONumber')) {
				If ($Column -notin $CSVColumns) {
					$MissingColumns += $Column
				}
			}
			If ($MissingColumns.Count -gt 0) {
				Throw ('{0}Error: the following columns are missing in the CSV file:{0}{1}' -f "`n", ($MissingColumns -join ', '))
			}
		}
		Else {
			Throw ('{0}Error: Input CSV file is empty.' -f "`n")
		}
	}

	# Connect to MainSite
	$MainSiteConnection = Connect-PnPOnline -Url $VDMSiteURL -ValidateConnection -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue

	# Compose CSV log file path
	$CSVLogFilePath = ('{0}\Logs\{1}_{2}.csv' -f $PSScriptRoot, $MyInvocation.MyCommand.Name, (Get-Date -Format 'dd-MM-yyyy_HH-mm-ss'))

	# Get 'Vendor Documents List' fields, then get the internal name of the 'Vendor Site URL' column
	[Array]$ListFields = Get-PnPField -List 'Vendor Documents List' -Connection $MainSiteConnection | Select-Object -Property InternalName, Title, TypeAsString
	$VendorSiteUrl_ColumnInteralName = $(($ListFields | Where-Object -FilterScript {
				$_.InternalName -eq 'VD_VendorName' -or
				$_.InternalName -eq 'VendorName_x003a_Site_x0020_Url'
			}).InternalName | Sort-Object -Descending -Top 1)

	# Get all items from 'Vendor Documents List'
	Write-Host "Getting all items from 'Vendor Documents List'..." -ForegroundColor Cyan
	[Array]$VDL_Items = Get-PnPListItem -List 'Vendor Documents List' -Connection $MainSiteConnection -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID              = $_['ID']
			TCM_DN          = $_['VD_DocumentNumber']
			Rev             = $_['VD_RevisionNumber']
			Index           = $_['VD_Index']
			PONumber        = $_['VD_PONumber']
			LastTransmittal = $_['LastTransmittal']
			VendorName      = $_['VD_VendorName'].LookupValue
			VendorSiteUrl   = $_[$VendorSiteUrl_ColumnInteralName].LookupValue
			Path            = $_['VD_DocumentPath']
		}
	}

	# Get all items from 'Process Flow Status List'
	Write-Host "Getting all items from 'Process Flow Status List'..." -ForegroundColor Cyan
	[Array]$PFSL_Items = Get-PnPListItem -List 'Process Flow Status List' -Connection $MainSiteConnection -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID            = $_['ID']
			TCM_DN        = $_['VD_DocumentNumber']
			Index         = $_['VD_Index']
			Rev           = $_['VD_RevisionNumber']
			VDL_ID        = $_['VD_VDL_ID']
			CSR_ID        = $_['VD_CommentsStatusReportID']
			PONumber      = $_['VD_PONumber']
			Status        = $_['VD_DocumentStatus']
			SubSitePOURL  = $($_['VD_PONumberUrl'].Url)
			VendorSiteUrl = $($_['VD_PONumberUrl'].Url -replace '/[^/]*$', '')
		}
	}

	# Get all items from 'Comment Status Report'
	Write-Host "Getting all items from 'Comment Status Report'..." -ForegroundColor Cyan
	[Array]$CSR_Items = Get-PnPListItem -List 'Comment Status Report' -Connection $MainSiteConnection -PageSize 5000 | ForEach-Object {
		[PSCustomObject]@{
			ID       = $_['ID']
			TCM_DN   = $_['VD_DocumentNumber']
			Index    = $_['VD_Index']
			Rev      = $_['VD_RevisionNumber']
			VDL_ID   = $_['VD_VDL_ID']
			PONumber = $_['VD_PONumber']
		}
	}

	# Loop through each input Document
	$CurrentDocumentCounter = 0
	ForEach ($Document in $TargetDocuments) {
		Write-Host ''
		# Create progress bar
		$CurrentDocumentCounter++
		$Progress = [PSCustomObject]@{
			Activity = "Processing Document '{0}'..." -f $Document.TCM_DN
			Status   = 'Processing {0}/{1}' -f $CurrentDocumentCounter, $TargetDocuments.Count
			Current  = $CurrentDocumentCounter
			Total    = $TargetDocuments.Count
			ID       = 0
		}
		Write-Host ('Document {0}/{1}: {2}' -f $CurrentDocumentCounter, $TargetDocuments.Count, $Document.TCM_DN, "`n") -ForegroundColor Green
		Write-Progress -Activity $Progress.Activity -Status $Progress.Status -CurrentOperation $Progress.Current -PercentComplete ($Progress.Current / $Progress.Total * 100)

		# Create Log entry for current document
		$LogEntry = [PSCustomObject]@{
			TCM_DN              = $Document.TCM_DN
			Rev                 = $null
			DestinationPONumber = $Document.DestinationPONumber
			PreviousPONumber    = $null
			Status              = $null
			Message             = $null
		}

		# Filter 'Vendor Documents List' items based on currently processed Document
		[Array]$Targeted_VDL_Item = $VDL_Items | Where-Object -FilterScript {
			$_.TCM_DN -eq $Document.TCM_DN
		}

		# Skip document if Document has been transmitted
		If (($Targeted_VDL_Item.LastTransmittal | Where-Object -FilterScript { $_ }).Count -gt 0) {
			# Warn the user on the console, add log entry to CSV file and continue to next Document
			$Msg = ("One or more revisions for this Document have a Transmittal so can't be moved." )
			Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
			$LogEntry.Status = 'Error'
			$LogEntry.Message = $Msg
			$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
			Continue
		}

		# Skip if Document is not found on VDL
		If ($Targeted_VDL_Item.Count -eq 0) {
			# Warn the user on the console, add log entry to CSV file and continue to next Document
			$Msg = ("No 'Vendor Documents List' item was found for TCM Document Number '{0}'." -f $Document.TCM_DN)
			Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
			$LogEntry.Status = 'Error'
			$LogEntry.Message = $Msg
			$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
			Continue
		}

		# Loop through each Document revision
		ForEach ($Revision in $Targeted_VDL_Item) {

			$MsgArray = @()
			$LogEntry.Rev = $Revision.Rev
			$LogEntry.PreviousPONumber = $Revision.PONumber

			# Filter 'Process Flow Status List' items based on currently processed Document revision
			$Targeted_PFSL_Item = $PFSL_Items | Where-Object -FilterScript {
				$_.TCM_DN -eq $Revision.TCM_DN -and
				$_.Rev -eq $Revision.Rev
			}

			# Filter 'Comment Status Report' items based on currently processed Document revision
			$Targeted_CSR_Item = $CSR_Items | Where-Object -FilterScript {
				$_.TCM_DN -eq $Revision.TCM_DN -and
				$_.Rev -eq $Revision.Rev
			}

			# Skip document if not found on 'Process Flow Status List'
			If ($Targeted_PFSL_Item.Count -eq 0) {
				# Warn the user on the console, add log entry to CSV file and continue to next Document revision
				$Msg = ("Document '{0}' (Rev {1}) not found on 'Process Flow Status List'." -f $Revision.TCM_DN, $Revision.Rev)
				Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
				$LogEntry.Status = 'Error'
				$LogEntry.Message = $Msg
				$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
				Continue
			}

			# Skip document if more then 1 Document is found on 'Process Flow Status List'
			If ($Targeted_PFSL_Item.Count -gt 1) {
				# Warn the user on the console, add log entry to CSV file and continue to next Document revision
				$Msg = ("More than 1 item found on 'Process Flow Status List' for TCM Document Number '{0}': {1}" -f $Revision.TCM_DN, $($Targeted_PFSL_Item.ID -join ', '))
				Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
				$LogEntry.Status = 'Error'
				$LogEntry.Message = $Msg
				$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
				Continue
			}

			# Skip document if more then 1 Document is found on 'Comment Status Report'
			If ($Targeted_CSR_Item.Count -gt 1) {
				# Warn the user on the console, add log entry to CSV file and continue to next Document revision
				$Msg = ("More than 1 item found on 'Comment Status Report' for TCM Document Number '{0}': {1}" -f $Revision.TCM_DN, $($Targeted_CSR_Item.ID -join ', '))
				Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
				$LogEntry.Status = 'Error'
				$LogEntry.Message = $Msg
				$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
				Continue
			}

			# Connect to Vendor Site
			$SubSiteSPOConnection = Connect-SPOSite -Url $Revision.VendorSiteUrl

			# Check both Document Libraries existance
			Get-PnPList -Identity $Revision.PONumber -Connection $SubSiteSPOConnection -ThrowExceptionIfListNotFound | Out-Null
			Get-PnPList -Identity $Document.DestinationPONumber -Connection $SubSiteSPOConnection -ThrowExceptionIfListNotFound | Out-Null

			# Check both Document Libraries existance
			$SourceFolder = Get-PnPFolder -Url $Revision.Path -Connection $SubSiteSPOConnection
			$NewDocumentUrlPath = $Revision.Path -replace $($Revision.PONumber), $Document.DestinationPONumber
			$NewPOUrl = $($NewDocumentUrlPath.Trim('/') -replace '/[^/]*$', '')

			If ($NewDocumentUrlPath -eq $Revision.Path) {
				# Warn the user on the console, add log entry to CSV file and continue to next Document revision
				$Msg = ('Destination PO Number is the same as the current PO Number')
				Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
				$LogEntry.Status = 'Error'
				$LogEntry.Message = $Msg
				$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
				Continue
			}

			$DestinationFolder = Get-PnPFolder -Url $NewDocumentUrlPath -Connection $SubSiteSPOConnection -ErrorAction SilentlyContinue
			If (-not $SourceFolder) {
				# Warn the user on the console, add log entry to CSV file and continue to next Document revision
				$Msg = ("Source folder '{0}' not found in '{1}'." -f $Revision.Path, $Revision.PONumber)
				Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
				$LogEntry.Status = 'Error'
				$LogEntry.Message = $Msg
				$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
				Continue
			}
			If ($DestinationFolder) {
				# Warn the user on the console, add log entry to CSV file and continue to next Document revision
				$Msg = ("Destination folder folder '{0}' already exists in '{1}'." -f $DestinationFolder.Name, $Document.DestinationPONumber)
				Write-Host ('Error: {0}' -f $Msg ) -ForegroundColor Red
				$LogEntry.Status = 'Error'
				$LogEntry.Message = $Msg
				$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force
				Continue
			}

			# Get all items from 'Revision Folder Dashboard'
			Write-Host "Getting all items from 'Revision Folder Dashboard'..." -ForegroundColor Cyan
			[Array]$CSR_Items = Get-PnPListItem -List 'Revision Folder Dashboard' -Connection $SubSiteSPOConnection -PageSize 5000 | ForEach-Object {
				[PSCustomObject]@{
					ID       = $_['ID']
					TCM_DN   = $_['VD_DocumentNumber']
					Rev      = $_['VD_RevisionNumber']
					PONumber = $_['VD_PONumber']
				}
			}

			# Get 'Revision Folder Dashboard' item
			$RFD_Item = Get-PnPListItem -List 'Revision Folder Dashboard' -Connection $SubSiteSPOConnection -PageSize 5000 | Where-Object -FilterScript {
				$_.TCM_DN -eq $Revision.TCM_DN -and
				$_.Rev -eq $Revision.Rev
			}

			#Move Document folder between PO
			Write-Host ("Moving folder '{0}' from '{1}' to '{2}'..." -f $Revision.Path, $Revision.PONumber, $Document.DestinationPONumber) -ForegroundColor Cyan
			Move-PnPFolder -Folder $SourceFolder -TargetFolder $NewPOUrl -Connection $SubSiteSPOConnection | Out-Null
			$MsgArray += 'Folder moved'

			# Update 'Vendor Documents List' item
			Write-Host ("Updating 'Vendor Documents List' item '{0}'..." -f $Revision.ID) -ForegroundColor Cyan
			Set-PnPListItem -List 'Vendor Documents List' -Identity $Revision.ID -Values @{
				VD_PONumber     = $Document.DestinationPONumber
				VD_DocumentPath = $NewDocumentUrlPath
			} -Connection $MainSiteConnection | Out-Null
			$MsgArray += 'VDL item updated'

			# Update 'Process Flow Status List' item
			Write-Host ("Updating 'Process Flow Status List' item '{0}'..." -f $Targeted_PFSL_Item.ID) -ForegroundColor Cyan
			Set-PnPListItem -List 'Process Flow Status List' -Identity $Targeted_PFSL_Item.ID -Values @{
				VD_PONumber    = $Document.DestinationPONumber
				VD_PONumberUrl = $NewPOUrl
			} -Connection $MainSiteConnection | Out-Null
			$MsgArray += 'PFSL item updated'

			# Update 'Comment Status Report' item (if exists)
			If ($Targeted_CSR_Item) {
				Write-Host ("Updating 'Comment Status Report' item '{0}'..." -f $Targeted_CSR_Item.ID) -ForegroundColor Cyan
				Set-PnPListItem -List 'Comment Status Report' -Identity $Targeted_CSR_Item.ID -Values @{
					VD_PONumber = $Document.DestinationPONumber
				} -Connection $MainSiteConnection | Out-Null
				$MsgArray += 'CSR item updated'
			}
			Else {
				Write-Host ("No 'Comment Status Report' item to udpate.") -ForegroundColor Yellow
				$MsgArray += 'No CSR item found'
			}

			# Update 'Revision Folder Dashboard' item
			If ($RFD_Item) {
				Write-Host ("Updating 'Revision Folder Dashboard' item '{0}'..." -f $RFD_Item.ID) -ForegroundColor Cyan
				Set-PnPListItem -List 'Revision Folder Dashboard' -Identity $RFD_Item.ID -Values @{
					VD_PONumber = $Document.DestinationPONumber
				} -Connection $SubSiteSPOConnection | Out-Null
				$MsgArray += 'RFD item updated'
			}
			Else {
				Write-Host ("No 'Revision Folder Dashboard' item to update.") -ForegroundColor Yellow
				$MsgArray += 'No RFD item found'
			}

			Write-Host ("Document move completed successfully for '{0}'.{1}" -f $Revision.TCM_DN, "`n") -ForegroundColor Green

			$LogEntry.Message = $MsgArray -join ', '
			$LogEntry.Status = 'Success'
			$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation -Force

		}
	}
	# Terminate progress bar
	Write-Progress -Activity $Progress.Activity -Status $Progress.Status -Completed
}
Catch {
	Write-Progress -Activity $Progress.Activity -Status $Progress.Status -Completed
	$LogEntry.Message = $_
	$LogEntry.Status = 'Error'
	$LogEntry | Export-Csv -Path $CSVLogFilePath -Delimiter ';' -Append -NoTypeInformation
	Throw
}