########################
### VARIABLES REGION ###
########################

# This script update $EmptyFieldToCorrect on ClientDocumentList and TransmittalQueueDetails_Registry as set on DocumentList
#ToDo:
# Check behaviour when duplicate us found on DL or CDL
# CSV Log

$SiteUrl = 'https://tecnimont.sharepoint.com/sites/43U4DigitalDocuments'
$EmptyFieldToCorrect = 'DocumentTitle'

# Array of TRN codes TCM to CLIENT
$IncompleteTrns = @(
	'4300-TCM-P4-T-01931'
	'4300-TCM-P4-T-01930',
	'4300-TCM-X4-T-01225',
	'4300-TCM-X4-T-01226',
	'4300-TCM-X4-T-01229',
	'4300-TCM-X4-T-01230',
	'4300-TCM-U4-T-02745',
	'4300-TCM-U4-T-02746',
	'4300-TCM-U4-T-02747',
	'4300-TCM-U4-T-02748',
	'4300-TCM-U4-T-02749',
	'4300-TCM-U4-T-02755'
)

##################
### END REGION ###
##################

Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Get DocumentList
$DocumentListItems = Get-PnPListItem -List 'DocumentList' -Query ("<View>
		<ViewFields>
			<FieldRef Name='ID'/>
			<FieldRef Name='Title'/>
			<FieldRef Name='IssueIndex'/>
			<FieldRef Name='LastTransmittal'/>
			<FieldRef Name='{0}'/>
		</ViewFields>
	</View>" -f $EmptyFieldToCorrect) -PageSize 5000 | ForEach-Object {
	$item = New-Object PSObject
	$item | Add-Member -MemberType NoteProperty -Name ID -Value $_['ID']
	$item | Add-Member -MemberType NoteProperty -Name TCMCode -Value $_['Title']
	$item | Add-Member -MemberType NoteProperty -Name IssueIndex -Value $_['IssueIndex']
	$item | Add-Member -MemberType NoteProperty -Name LastTransmittal -Value $_['LastTransmittal']
	$item | Add-Member -MemberType NoteProperty -Name "$EmptyFieldToCorrect" -Value $_["$EmptyFieldToCorrect"]
	$item
}
$DocsToCorrect = $DocumentListItems | Where-Object -FilterScript { $_.LastTransmittal -in $IncompleteTrns }
#$DocsToCorrect | Out-GridView

# Get TransmittalQueueDetails_Registry
$TrnRegistryDocItems = Get-PnPListItem -List 'TransmittalQueueDetails_Registry' -Query ("<View>
		<ViewFields>
			<FieldRef Name='ID'/>
			<FieldRef Name='Title'/>
			<FieldRef Name='IssueIndex'/>
			<FieldRef Name='TransmittalID'/>
			<FieldRef Name='{0}'/>
		</ViewFields>
	</View>" -f $EmptyFieldToCorrect) -PageSize 5000 | ForEach-Object {
	$item = New-Object PSObject
	$item | Add-Member -MemberType NoteProperty -Name ID -Value $_['ID']
	$item | Add-Member -MemberType NoteProperty -Name TCMCode -Value $_['Title']
	$item | Add-Member -MemberType NoteProperty -Name IssueIndex -Value $_['IssueIndex']
	$item | Add-Member -MemberType NoteProperty -Name TransmittalID -Value $_['TransmittalID']
	$item | Add-Member -MemberType NoteProperty -Name "$EmptyFieldToCorrect" -Value $_["$EmptyFieldToCorrect"]
	$item
}
$TrnItemsToCorrect = $TrnRegistryDocItems | Where-Object -FilterScript { $_.TransmittalID -in $IncompleteTrns }
#$TrnItemsToCorrect | Out-GridView

Write-Host 'Updating TransmittalQueueDetails_Registry'
ForEach ($Item in $TrnItemsToCorrect)
{

	$ValueToCorrect = ($DocsToCorrect | Where-Object -FilterScript { $_.TCMCode -eq $Item.TCMCode -and $_.IssueIndex -eq $Item.IssueIndex })."$EmptyFieldToCorrect"

	# Update list TransmittalQueueDetails_Registry
	Try
	{
		Set-PnPListItem -List 'TransmittalQueueDetails_Registry' -Identity $Item.ID -Values @{"$EmptyFieldToCorrect" = "$ValueToCorrect" }
		Write-Host ('SUCCESSFULLY set "{0}" as "{1}" for item {2} on TransmittalQueueDetails_Registry (Document {3} - {4})' -f $ValueToCorrect, $EmptyFieldToCorrect, $Item.ID, $Item.TCMCode, $Item.IssueIndex) -ForegroundColor Green
	}
 Catch
	{
		Write-Host ('FAILED to set "{0}" as "{1}" for item {2} on TransmittalQueueDetails_Registry (Document {3} - {4})' -f $ValueToCorrect, $EmptyFieldToCorrect, $Item.ID, $Item.TCMCode, $Item.IssueIndex) -ForegroundColor Red
		Exit
	}
}

$SiteUrl_C = $SiteUrl + 'C'
Connect-PnPOnline -Url $SiteUrl_C -UseWebLogin

# Get Client Document List
$CDLDocItems = Get-PnPListItem -List 'Client Document List' -Query ("<View>
		<ViewFields>
			<FieldRef Name='ID'/>
			<FieldRef Name='Title'/>
			<FieldRef Name='IssueIndex'/>
			<FieldRef Name='LastTransmittal'/>
			<FieldRef Name='{0}'/>
		</ViewFields>
	</View>" -f $EmptyFieldToCorrect) -PageSize 5000 | ForEach-Object {
	$item = New-Object PSObject
	$item | Add-Member -MemberType NoteProperty -Name ID -Value $_['ID']
	$item | Add-Member -MemberType NoteProperty -Name TCMCode -Value $_['Title']
	$item | Add-Member -MemberType NoteProperty -Name IssueIndex -Value $_['IssueIndex']
	$item | Add-Member -MemberType NoteProperty -Name LastTransmittal -Value $_['LastTransmittal']
	$item | Add-Member -MemberType NoteProperty -Name "$EmptyFieldToCorrect" -Value $_["$EmptyFieldToCorrect"]
	$item
}
$CDLItemsToCorrect = $CDLDocItems | Where-Object -FilterScript { $_.LastTransmittal -in $IncompleteTrns }
#$CDLItemsToCorrect | Out-GridView

Write-Host 'Updating Client Document List'
ForEach ($Item in $CDLItemsToCorrect)
{

	$ValueToCorrect = ($DocsToCorrect | Where-Object -FilterScript { $_.TCMCode -eq $Item.TCMCode -and $_.IssueIndex -eq $Item.IssueIndex })."$EmptyFieldToCorrect"

	# Update list "Client Document List"
	Try
	{
		Set-PnPListItem -List 'Client Document List' -Identity $Item.ID -Values @{"$EmptyFieldToCorrect" = "$ValueToCorrect" }
		Write-Host ('SUCCESSFULLY set "{0}" as "{1}" for item {2} on Client Document List (Document {3} - {4})' -f $ValueToCorrect, $EmptyFieldToCorrect, $Item.ID, $Item.TCMCode, $Item.IssueIndex) -ForegroundColor Green
	}
 Catch
	{
		Write-Host ('FAILED to set "{0}" as "{1}" for item {2} on Client Document List (Document {3} - {4})' -f $ValueToCorrect, $EmptyFieldToCorrect, $Item.ID, $Item.TCMCode, $Item.IssueIndex) -ForegroundColor Red
		Exit
	}
}