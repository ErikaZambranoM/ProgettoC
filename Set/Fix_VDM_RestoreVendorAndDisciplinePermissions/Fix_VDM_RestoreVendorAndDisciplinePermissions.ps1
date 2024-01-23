#Script to fix permissions as raised issue in INC0718246
#Partially tested, not all messages are up to date, but operations seem fine.
#PnP.Powershell 1.12.0
#UPDATE 2022.12.28
#SCOPO:
# QUESTO SCRIPT, CREATO PER I PORTALI VDM AGGIORNA I PERMESSI DELLE FOLDER PER UN PO SPECIFICO
# I PERMESSI CHE VENGONO VERIFICATI SONO  QUELLI DELLE DISCIPLINE E DEL VENDOR PER LE REVISION FOLDER
# E PER LE SOTTOCARTELLE
# ES USO ./Fix_VDM_RestoreVendorAndDisciplinePermissions.ps1 -VendorPortalUrl "https://tecnimont.sharepoint.com/sites/vdm_43P4" -PONumber "7500106086"

param(
	[Parameter(Mandatory)]
	$VendorPortalUrl,
	[Parameter(Mandatory)]
	$PONumber
)

#START CHECKS
if ($PONumber -eq $null -or $VendorPortalUrl -eq $null -or $VendorPortalUrl.Length -le 0)
{
	Write-Host 'Missing starts parameters PONumber - VendorPortalUrl' -f yellow
	return
}

Write-Host "Fix permission tool starts for $VendorPortalUrl and $PONumber" -f green
#Permissions
$PermissionContribute = 'MT Contributors'
$PermissionVendorContribute = 'MT Contributors - Vendor'
$PermissionReader = 'MT Readers'
#$PermissionFullCtrl = "Full Control"


#Site and lists urls
$VendorPortal = $VendorPortalUrl #"https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
$VendorDocListUrl = 'Lists/VendorDocumentsList'
$VendorsListURL = 'Lists/Vendors'
$DisciplineListTitle = 'Disciplines'

$RootConnection = Connect-PnPOnline -Url $VendorPortal -UseWebLogin -ReturnConnection -ValidateConnection -WarningAction SilentlyContinue

$SubsiteConnection = $null
$currentSubsite = $null
#Get PONumber, DocNumber,'VD_VendorName', "VD_DisciplineOwnerTCM", "VD_DisciplinesTCM", "VD_Index"
# SOSTITUITO CON IL CAML
#$VendorDocRecord = (Get-PnPListItem -Connection $RootConnection -List $VendorDocListUrl -Fields 'VD_PONumber', 'VD_DocumentNumber', 'VD_VendorName', "VD_DisciplineOwnerTCM", "VD_DisciplinesTCM", "VD_Index").FieldValues |
#Where-Object { $_.VD_PONumber -eq $($PONumber) }

$poFilter = '<View><ViewFields><FieldRef Name="VD_PONumber"/><FieldRef Name="VD_DocumentNumber"/><FieldRef Name="VD_VendorName"/><FieldRef Name="VD_DisciplineOwnerTCM"/><FieldRef Name="VD_DisciplinesTCM"/><FieldRef Name="VD_Index"/></ViewFields><Query><Where><Eq><FieldRef Name="VD_PONumber"/><Value Type="Text">' + $PONumber + '</Value></Eq></Where></Query></View>'
# RESTRICTED FILTER
#$poFilter = '<View><ViewFields><FieldRef Name="VD_PONumber"/><FieldRef Name="VD_DocumentNumber"/><FieldRef Name="VD_VendorName"/><FieldRef Name="VD_DisciplineOwnerTCM"/><FieldRef Name="VD_DisciplinesTCM"/><FieldRef Name="VD_Index"/></ViewFields><Query><Where><And><And><Geq><FieldRef Name="Created" /><Value IncludeTimeValue="FALSE" Type="DateTime">2022-11-30</Value></Geq><Leq><FieldRef Name="Created" /><Value IncludeTimeValue="FALSE" Type="DateTime">2022-11-30</Value></Leq></And><Eq><FieldRef Name="VD_PONumber"/><Value Type="Text">' + $PONumber + '</Value></Eq></And></Where></Query></View>'
#$poFilter = '<View><ViewFields><FieldRef Name="VD_PONumber"/><FieldRef Name="VD_DocumentNumber"/><FieldRef Name="VD_VendorName"/><FieldRef Name="VD_DisciplineOwnerTCM"/><FieldRef Name="VD_DisciplinesTCM"/><FieldRef Name="VD_Index"/></ViewFields><Query><Where><And><And><Geq><FieldRef Name="Created" /><Value IncludeTimeValue="FALSE" Type="DateTime">2022-11-29</Value></Geq><Leq><FieldRef Name="Created" /><Value IncludeTimeValue="FALSE" Type="DateTime">2022-12-01</Value></Leq></And><Eq><FieldRef Name="VD_PONumber"/><Value Type="Text">' + $PONumber + '</Value></Eq></And></Where></Query></View>'
$VendorDocRecord = (Get-PnPListItem -Connection $RootConnection -List $VendorDocListUrl -PageSize 200 -Query $poFilter  ).FieldValues

#MANCAVA
$DisciplineListItems = Get-PnPListItem -Connection $RootConnection -List $DisciplineListTitle -PageSize 200

$Vendor = $VendorDocRecord[0].VD_VendorName.LookupValue
#Get VD Subsite Url, Title, Vendor Code, Group Name
$VendorRecord = (Get-PnPListItem -Connection $RootConnection -List $VendorsListUrl -Fields 'VD_SiteUrl', 'Title', 'VD_VendorCode', 'VD_GroupName').FieldValues | Where-Object { $_.Title -eq $Vendor }
$VendorGroup = $null

#Returns the group from lookupValue
function FindGroupByDisciplineLookup()
{
	param($DSLookupValue, $DiscListItems, $RootConnection)
	#Write-Host "Lookup $DSLookupValue"
	#Write-Host "DiscListItems $($DiscListItems.Count)"
	#$GroupName = "DS $($DSLookupValue)" ASSUNZIONE ERRATA. Meglio cercare dalla lista discipline
	$CurrDiscipline = $DiscListItems | Where-Object { $_.FieldValues.Title -eq $DSLookupValue }
	#Write-Host "Current discipline $($CurrDiscipline.FieldValues.VD_GroupName)"
	$GroupName = $CurrDiscipline.FieldValues.VD_GroupName
	$Group = Get-PnPGroup -Identity $GroupName -Connection $RootConnection
	#Write-Host "Current Discipline Group $($Group.Title)"
	return $Group
}

#NUOVE FUNZIONI PER GESTIONE ROLE ASSIGNMENT



function BreakInheritance
{
	param($item, $connection)
	#AP: Non dovrebbe essere così?
	#if($item.HasUniqueRoleAssignments -eq $true) {return}
	if ($item.HasUniqueRoleAssignments -eq $false) { return }
	$item.BreakRoleInheritance($True, $True)
	Invoke-PnPQuery -Connection $connection
}

function ResetInheritance
{
	param($item, $connection)
	#AP: Non dovrebbe essere così?
	#if($item.HasUniqueRoleAssignments -eq $false) {return}
	if ($item.HasUniqueRoleAssignments -eq $true) { return }
	$item.ResetRoleInheritance()
	$item.update()
	Invoke-PnPQuery -Connection $connection
}


function RemoveAllRoleAssignment
{
	param($groupName, $itemRoleAssignments, $connection)
	if ($itemRoleAssignments.Count -gt 0)
	{
		foreach ($roleAssignmentBinding in $itemRoleAssignments.RoleDefinitionBindings)
		{
			Write-Host "Removing $($roleAssignmentBinding.Name) from $groupName" -f Yellow
			$itemRoleAssignments.RoleDefinitionBindings.Remove($roleAssignmentBinding)
			$itemRoleAssignments.Update()
			Invoke-PnPQuery -Connection $connection
		}
	}
}

function RestoreRequiredRoleAssignment
{
	param($listTitle, $itemID, $groupName, $requiredPermission, $itemRoleAssignments, $removeOtherPermissions, $connection)

	#Write-Host "Role assignments for group $groupName"
	#$itemRoleAssignments | Foreach-Object { Write-Host "All role assignments $($_.RoleDefinitionBindings.Name)" }
	$hasRequiredAssignments = $itemRoleAssignments | Where-Object { $_.RoleDefinitionBindings.Name -eq $requiredPermission }
	$hasInvalidAssignments = $itemRoleAssignments | Where-Object { $_.RoleDefinitionBindings.Name -ne $requiredPermission }
	#$hasRequiredAssignments | Foreach-Object { Write-Host "Required role assignments $($_.RoleDefinitionBindings.Name)" }
	if ($hasRequiredAssignments.Count -gt 0 -and $removeOtherPermissions -eq $false)
	{
		return
	}
	if ($hasRequiredAssignments.Count -eq 0)
	{
  Write-Host "Adding $requiredPermission to $groupName" -f Yellow
  Set-PnPListItemPermission -List $listTitle -Identity $itemID -Group $groupName -AddRole $requiredPermission -SystemUpdate -Connection $connection
	}
	if ($hasInvalidAssignments.Count -gt 0 -and $removeOtherPermissions -eq $true)
 {
		Write-Host 'There are some permissions to remove'
		$RolesToRemove = $itemRoleAssignments.RoleDefinitionBindings | Where-Object { $_.Name -ne $requiredPermission }
		foreach ($Role in $RolesToRemove)
		{
			#REMOVE ROLE ASSIGNMENT
			Write-Host "Removing $($Role.Name) from $groupName" -f Yellow
			$itemRoleAssignments.RoleDefinitionBindings.Remove($Role)
			$itemRoleAssignments.Update()
			Invoke-PnPQuery -Connection $connection
		}
	}
}


#Iterate items
$processedRecords = 0
$processedRecordsProgress = 0
Write-Host "Records to manage $($VendorDocRecord.Count)"
foreach ($record in $VendorDocRecord)
{
	$processedRecordsProgress = ($processedRecords / $VendorDocRecord.Count) * 100
	$OuterLoopProgressParameters = @{
		Activity         = "Updating $($record.VD_DocumentNumber)-$($record.VD_Index.toString('000'))"
		Status           = 'Progress->'
		PercentComplete  = $processedRecordsProgress
		CurrentOperation = 'Processing folder'
	}
	Write-Progress @OuterLoopProgressParameters


	$DisciplineOwnerGroup = FindGroupByDisciplineLookup -DSLookupValue $record.VD_DisciplineOwnerTCM.LookupValue -DiscListItems $DisciplineListItems -RootConnection $RootConnection
	$Disciplines = @()
	$DisciplinesNames = @()

	$record.VD_DisciplinesTCM | ForEach-Object {
		$DisciplineGroup = FindGroupByDisciplineLookup -DSLookupValue $_.LookupValue -DiscListItems $DisciplineListItems -RootConnection $RootConnection
		$Disciplines += ($DisciplineGroup)
		$DisciplinesNames += ($DisciplineGroup.Title)
	}

	#Get and store connection to subsite
	if ($null -eq $SubsiteConnection -or $currentSubsite -ne $vendorRecord.VD_SiteUrl)
	{
		$SubsiteConnection = Connect-PnPOnline -Url $vendorRecord.VD_SiteUrl -UseWebLogin -ReturnConnection -ValidateConnection -WarningAction SilentlyContinue
		$currentSubsite = $vendorRecord.VD_SiteUrl
		#AGGIUNTO
		$VendorGroup = Get-PnPGroup -Identity $vendorRecord.VD_GroupName -Connection $RootConnection
	}

	#$POList = Get-PnpList -Connection $SubsiteConnection | Where-Object { $_.Title -eq $record.VD_PONumber }
	#$POList = Get-PnpList -Connection $SubsiteConnection -Identity $record.VD_PONumber

	#IMPORTANT: Get-PnpListItem is recursive
	#$itemsToProcess = Get-PnPListItem -List $POList.Id -PageSize 500 -Connection $SubsiteConnection
	$itemsToProcess = Get-PnPListItem -List $record.VD_PONumber -PageSize 500 -Connection $SubsiteConnection

	#I distinguish parent folder from children like this
	#Parent has VD_DocumentStatus set
	#$rootFoldersList = $itemsToProcess | Where-Object { $null -ne $_.FieldValues.VD_DocumentStatus }
	#Child has title set, but has not VD_DocumentStatus
	#$firstLevelSubFolders = $itemsToProcess | Where-Object { $null -eq $_.FieldValues.VD_DocumentStatus -and $null -ne $_.FieldValues.Title }
	#files have both VD_DocumentStatus and Title empty

	#Sostituite cosi: - oggetti di tipo folder con path di lunghezza specifica. la struttura è posizionale.
	$rootFoldersList = $itemsToProcess | Where-Object { $_.FieldValues.FSObjType -eq 1 -and ($_.FieldValues.FileDirRef.split('/').Count -eq 5 ) }
	$firstLevelSubFolders = $itemsToProcess | Where-Object { $_.FieldValues.FSObjType -eq 1 -and ($_.FieldValues.FileDirRef.split('/').Count -eq 6 ) }

	#Get RootFolder for current item
	#$folderListItem = $rootFoldersList | Where-Object { $_.FieldValues.VD_DocumentNumber -eq $record.VD_DocumentNumber}
	$folderListItem = $rootFoldersList | Where-Object { $_.FieldValues.FileLeafRef -eq "$($record.VD_DocumentNumber)-$($record.VD_Index.toString('000'))" }

	#$folderItem = (Get-PnPListItem -List $POList.Id -UniqueId $folderName.FieldValues.UniqueId-Connection $SubsiteConnection).FieldValues
	#$folderItem = (Get-PnPListItem -List $record.VD_PONumber -UniqueId $folderName.FieldValues.UniqueId-Connection $SubsiteConnection).FieldValues
	$folderListItemValues = $folderListItem.FieldValues

	#Get status and assignments
	#$DocumentStatus = $folderName.FieldValues.VD_DocumentStatus
	$DocumentStatus = $folderListItemValues.VD_DocumentStatus
	#Find subfolders, by matching FileRef on parent with FileDirRef on child
	$ownedSubFolders = $firstLevelSubFolders | Where-Object { $_.FieldValues.FileDirRef -eq $folderListItemValues.FileRef }

	#$folderItem1 = Get-PnPListItem -List $POList.Id -Id $folderItem.ID -Connection $SubsiteConnection
	Get-PnPProperty -ClientObject $folderListItem -Property HasUniqueRoleAssignments, RoleAssignments -Connection $SubsiteConnection
	if ($folderListItem.HasUniqueRoleAssignments -eq $True)
	{
		foreach ($roleAssignment in  $folderListItem.RoleAssignments)
		{
			Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings, Member -Connection $SubsiteConnection
		}
		#check if exists owner discipline
		#$existingOwnerDiscipline = $folderListItem.RoleAssignments | Where-Object { $_.Member.Title -eq $DisciplineOwnerGroup }
		$existingOwnerDiscipline = $folderListItem.RoleAssignments | Where-Object { $_.Member.Title -eq $DisciplineOwnerGroup.Title }
		#check if exist other disciplines
		$existingDisciplines = $folderListItem.RoleAssignments | Where-Object { $DisciplinesNames -contains $_.Member.Title }
		#Write-Host "Roles count $($folderListItem.RoleAssignments.Count), for disciplines $($existingDisciplines.Count)"
		#$existingDisciplines | Foreach-Object { Write-Host "ExistingDisciplines role assignments $($_.RoleDefinitionBindings.Name) for $($_.Member.Title)" }

		#AGGIUNTO -
		$existingVendor = $folderListItem.RoleAssignments | Where-Object { $_.Member.Title -eq $VendorGroup.Title }

		#Set permissions: add only if missing
		#RIVEDERE TUTTO IL BLOCCO.
		Write-Host "Process $($record.VD_DocumentNumber)-$($record.VD_Index.toString('000')) . Status $DocumentStatus"
		switch ($DocumentStatus)
		{
			'Waiting'
			{
				Write-Host 'Skip' -f Yellow
				break
			}
			{ ($_ -eq 'Placeholder') -or ($_ -eq 'Rejected') }
			{

				#Write-Host "Internal loop Placeholder/Rejected"
				#RIPARTIRE DA QUI
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $DisciplineOwnerGroup.Title	-requiredPermission $PermissionContribute -itemRoleAssignments $existingOwnerDiscipline -removeOtherPermissions $true -connection $SubsiteConnection
				$Disciplines | ForEach-Object {
					#Write-Host "Status PlaceHolder - List $PONumber - $($folderName.FieldValues.Title) - Add Role $PermissionContribute to Group $($_.Title)"
					$thisDiscipline = $_.Title
					$existingCurrentDisciplines = $existingDisciplines | Where-Object { $thisDiscipline -eq $_.Member.Title }
					RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $thisDiscipline -requiredPermission $PermissionContribute -itemRoleAssignments $existingCurrentDisciplines -removeOtherPermissions $true -connection $SubsiteConnection
				}
				#MANCA IL VENDOR!
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $VendorGroup.Title -requiredPermission $PermissionVendorContribute -itemRoleAssignments $existingVendor -removeOtherPermissions $true -connection $SubsiteConnection
				break
			}
			'Received'
			{
				#Write-Host "Internal loop Received"
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $DisciplineOwnerGroup.Title	-requiredPermission $PermissionContribute -itemRoleAssignments $existingOwnerDiscipline -removeOtherPermissions $true -connection $SubsiteConnection
				$Disciplines | ForEach-Object {
					$thisDiscipline = $_.Title
					$existingCurrentDisciplines = $existingDisciplines | Where-Object { $thisDiscipline -eq $_.Member.Title }
					RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $thisDiscipline -requiredPermission $PermissionContribute -itemRoleAssignments $existingCurrentDisciplines -removeOtherPermissions $true -connection $SubsiteConnection
				}
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $VendorGroup.Title -requiredPermission $PermissionReader -itemRoleAssignments $existingVendor -removeOtherPermissions $true -connection $SubsiteConnection
				break

			}
			'Commenting'
			{
				#Write-Host "Internal loop Commenting"
				#RESTORE PERMSSIONS AT ROOT LEVEL
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $DisciplineOwnerGroup.Title	-requiredPermission $PermissionContribute -itemRoleAssignments $existingOwnerDiscipline -removeOtherPermissions $true -connection $SubsiteConnection
				$Disciplines | ForEach-Object {
					#Write-Host "Status PlaceHolder - List $PONumber - $($folderName.FieldValues.Title) - Add Role $PermissionContribute to Group $($_.Title)"
					$thisDiscipline = $_.Title
					$existingCurrentDisciplines = $existingDisciplines | Where-Object { $thisDiscipline -eq $_.Member.Title }
					RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $thisDiscipline -requiredPermission $PermissionContribute -itemRoleAssignments $existingCurrentDisciplines -removeOtherPermissions $true -connection $SubsiteConnection
				}
				#REMOVE VENDOR
				RemoveAllRoleAssignment -groupName $VendorGroup.Title -itemRoleAssignments $existingVendor -connection $SubsiteConnection
				#GESTIONE SUB FOLDER
				if ($null -ne $ownedSubFolders)
				{
					$ownedSubFolders | ForEach-Object {
						$subFolderItem = ($_)
						Get-PnPProperty -ClientObject $subFolderItem -Property HasUniqueRoleAssignments, RoleAssignments -Connection $SubsiteConnection

						switch ($subFolderItem.FieldValues.FileLeafRef.ToUpper())
						{
							{ ($_ -eq 'ATTACHMENTS') -or ($_ -eq 'OFV') }
							{
								#EREDITA
								ResetInheritance -item $subFolderItem -connection $SubsiteConnection
								break
							}
							'FROMCLIENT'
							{
								#Disciplina in contribute. Vendor assente
								BreakInheritance -item $subFolderItem -connection $SubsiteConnection
								foreach ($subRoleAssignment in  $subFolderItem.RoleAssignments)
								{
									Get-PnPProperty -ClientObject $subRoleAssignment -Property RoleDefinitionBindings, Member -Connection $SubsiteConnection
								}
								$existingOwnerDisciplineInSub = $subFolderItem.RoleAssignments | Where-Object { $_.Member.Title -eq $DisciplineOwnerGroup.Title }
								$existingDisciplinesInSub = $subFolderItem.RoleAssignments | Where-Object { $DisciplinesNames -contains $_.Member.Title }
								$existingVendorInSub = $subFolderItem.RoleAssignments | Where-Object { $_.Member.Title -eq $VendorGroup.Title }

								RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $subFolderItem.ID -groupName $DisciplineOwnerGroup.Title -requiredPermission $PermissionContribute -itemRoleAssignments $existingOwnerDisciplineInSub -removeOtherPermissions $true -connection $SubsiteConnection
								$Disciplines | ForEach-Object {
									#Write-Host "Status PlaceHolder - List $PONumber - $($folderName.FieldValues.Title) - Add Role $PermissionContribute to Group $($_.Title)"
									$thisDiscipline = $_.Title
									$existingCurrentDisciplines = $existingDisciplinesInSub | Where-Object { $thisDiscipline -eq $_.Member.Title }
									RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $subFolderItem.ID -groupName $thisDiscipline -requiredPermission $PermissionContribute -itemRoleAssignments $existingCurrentDisciplines -removeOtherPermissions $true -connection $SubsiteConnection
								}
								#REMOVE VENDOR
								RemoveAllRoleAssignment -groupName $VendorGroup.Title -itemRoleAssignments $existingVendorInSub -connection $SubsiteConnection

								break
							}
						}
					}
				}
				break

			}
			'Closed'
			{
				#Write-Host "Closed"
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $DisciplineOwnerGroup.Title	-requiredPermission $PermissionReader -itemRoleAssignments $existingOwnerDiscipline -removeOtherPermissions $true -connection $SubsiteConnection
				$Disciplines | ForEach-Object {
					$thisDiscipline = $_.Title
					$existingCurrentDisciplines = $existingDisciplines | Where-Object { $thisDiscipline -eq $_.Member.Title }
					#Write-Host "Closed - Existing assignements found $($existingCurrentDisciplines.Count)"
					RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $thisDiscipline -requiredPermission $PermissionReader -itemRoleAssignments $existingCurrentDisciplines -removeOtherPermissions $true -connection $SubsiteConnection
				}
				RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $folderListItem.ID -groupName $VendorGroup.Title -requiredPermission $PermissionReader -itemRoleAssignments $existingVendor -removeOtherPermissions $true -connection $SubsiteConnection


		  #GESTIONE SUB FOLDER
				if ($null -ne $ownedSubFolders)
				{
					#Write-Host "ownedSubFolders Count $($ownedSubFolders.Count)"
					$ownedSubFolders | ForEach-Object {
						$subFolderItem = $_
						Get-PnPProperty -ClientObject $subFolderItem -Property HasUniqueRoleAssignments, RoleAssignments -Connection $SubsiteConnection
						#Write-Host "subFolderItem $($subFolderItem.FieldValues.FileLeafRef) $($subFolderItem.ID), $($subFolderItem.HasUniqueRoleAssignments)"

						switch ($subFolderItem.FieldValues.FileLeafRef.ToUpper())
						{
							{ ($_ -eq 'ATTACHMENTS') -or ($_ -eq 'OFV') -or ($_ -eq 'CMTD') }
							{
								#EREDITA
								ResetInheritance -item $subFolderItem -connection $SubsiteConnection
								break
							}
							'FROMCLIENT'
							{
								#Disciplina in contribute. Vendor assente
								BreakInheritance -item $subFolderItem -connection $SubsiteConnection
								foreach ($subRoleAssignment in  $subFolderItem.RoleAssignments)
								{
									Get-PnPProperty -ClientObject $subRoleAssignment -Property RoleDefinitionBindings, Member -Connection $SubsiteConnection
								}
								$existingOwnerDisciplineInSub = $subFolderItem.RoleAssignments | Where-Object { $_.Member.Title -eq $DisciplineOwnerGroup.Title }
								$existingDisciplinesInSub = $subFolderItem.RoleAssignments | Where-Object { $DisciplinesNames -contains $_.Member.Title }
								$existingVendorInSub = $subFolderItem.RoleAssignments | Where-Object { $_.Member.Title -eq $VendorGroup.Title }

								RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $subFolderItem.ID -groupName $DisciplineOwnerGroup.Title -requiredPermission $PermissionReader -itemRoleAssignments $existingOwnerDisciplineInSub -removeOtherPermissions $true -connection $SubsiteConnection
								$Disciplines | ForEach-Object {
									#Write-Host "Status PlaceHolder - List $PONumber - $($folderName.FieldValues.Title) - Add Role $PermissionReader to Group $($_.Title)"
									$thisDiscipline = $_.Title
									$existingCurrentDisciplines = $existingDisciplinesInSub | Where-Object { $thisDiscipline -eq $_.Member.Title }
									RestoreRequiredRoleAssignment -listTitle $record.VD_PONumber -itemID $subFolderItem.ID -groupName $thisDiscipline -requiredPermission $PermissionReader -itemRoleAssignments $existingCurrentDisciplines -removeOtherPermissions $true -connection $SubsiteConnection
								}
								#REMOVE VENDOR
								RemoveAllRoleAssignment -groupName $VendorGroup.Title -itemRoleAssignments $existingVendorInSub -connection $SubsiteConnection

								break
							}
						}
					}

				}

				break
			}
		}


	}



	$processedRecords = $processedRecords + 1
}