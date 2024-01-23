#SCRIPT - CLEAN EXPIRED TASKS ON PLANNER
#PNP VERSION: 1.12
#AUTHOR: G.VIGNATI
#LAST UPDATE: 30.12.2022
#SCRIPT VERSION: 1.0
# EXAMPLE .\Fix_CleanExpiredPlannerTasks.ps1 -siteurl https://tecnimont.sharepoint.com/sites/poc_VDM

param(
	[Parameter(Mandatory = $true)][String]$siteurl
)

#START CHECKS
if ($siteurl -eq $null -or $siteurl.Length -le 0) {
	Write-Host 'Missing starts parameters siteurl' -f yellow
	return
}

# IMPOSTO IL FILTRO DI CANCELLAZIONE A 10 GIORNI DALLA DATA CORRENTE
$filterDate = (Get-Date).AddDays(-10)

Connect-PnPOnline $siteurl -Interactive

$projectDisciplines = Get-PnPListItem -List 'Disciplines'
$processedDisciplines = 0
$removedTasks = 0
$updatedComments = 0
$processedTasks = 0

foreach ($impactedDiscipline in $projectDisciplines) {

	#PER OGNI DISCIPLINA DELLA LISTA RECUPERO IL PLANNER ID ed I TASKS

	$plannerId = $impactedDiscipline.FieldValues.VD_PlannerID
	$disciplineName = $impactedDiscipline.FieldValues.Title
	Write-Host "$disciplineName under process"

	if ($plannerId -eq $null) {
		Write-Host "$disciplineName Skipped"
		continue
	}
	$tasks = Get-PnPPlannerTask -PlanId $plannerId

	$processedDisciplineProgress = ($processedDisciplines / $projectDisciplines.Count) * 100
	$OuterLoopProgressParameters = @{
		Activity         = "Processing $disciplineName. Tasks $($tasks.Count)"
		Status           = 'Progress'
		PercentComplete  = $processedDisciplineProgress
		CurrentOperation = 'Processing discipline'
	}
	Write-Progress @OuterLoopProgressParameters



	#SELEZIONO SOLO TASK COMPLETATI CON DATA DI CHIUSURA SUPERATA DA X GIORNI
	$tasksToDelete = $tasks | Where-Object { $filterDate -gt $_.CompletedDateTime -and $_.PercentComplete -eq 100 }
	$tasksToDelete = $tasksToDelete | Sort-Object -Property CompletedDateTime

	$processedTasksForDiscipline = 0
	$tasksToDelete | ForEach-Object {

		$fullTcmCode = $_.Title.Split(' ')[0]
		$tcmCode = $fullTcmCode.Substring(0, $fullTcmCode.Length - 4)
		$index = [int]$fullTcmCode.Substring($fullTcmCode.Length - 3)
		$taskPlanId = $_.PlanId
		$taskClosureDate = $_.CompletedDateTime

		$processedRecordsProgress = ($processedTasksForDiscipline / $tasksToDelete.Count) * 100
		$InnerLoopProgressParameters = @{
			ID               = 1
			Activity         = "Updating $fullTcmCode"
			Status           = 'Progress'
			PercentComplete  = $processedRecordsProgress
			CurrentOperation = 'Processing tasks'
		}
		Write-Progress @InnerLoopProgressParameters

		# CERCO SULLA LISTA PROCESS FLOW IL DOCUMENTO ASSOCIATO AL TASK, SE E' IN STATO CLOSED VALUTO LA CANCELLAZIONE
		$documentFilter = '<View><ViewFields><FieldRef Name="VD_PONumber"/><FieldRef Name="VD_DocumentNumber"/><FieldRef Name="VD_Index"/><FieldRef Name="VD_DocumentStatus"/><FieldRef Name="VD_VDL_ID"/></ViewFields><Query><Where><And><Eq><FieldRef Name="VD_DocumentNumber"/><Value Type="Text">' + $tcmCode + '</Value></Eq><Eq><FieldRef Name="VD_Index"/><Value Type="Number">' + $index + '</Value></Eq></And></Where></Query></View>'
		$processFlowItem = Get-PnPListItem -List 'Process Flow Status List' -Query $documentFilter
		if ($processFlowItem.Count -eq 1 -and $processFlowItem.FieldValues.VD_DocumentStatus -eq 'Closed') {
		 #Posso rimuovere il task se lo trovo chiuso in process flow
		 #Per prima cosa controllo i commenti
		 $vdlID = $processFlowItem.FieldValues.VD_VDL_ID
		 $commentsFilter = '<View><Query><Where><Eq><FieldRef Name="VD_VDL_ID"/><Value Type="Number">' + $vdlID + '</Value></Eq></Where></Query></View>'
		 $commentsItem = Get-PnPListItem -List 'Comment Status Report' -Query $commentsFilter
		 #$impactedDiscipline = $projectDisciplines | Where-Object{$_.FieldValues.VD_PlannerID -eq $taskPlanId }
		 if ($impactedDiscipline.Count -eq 1) {
				$disciplineField = $commentsItem.FieldValues.Keys | Where-Object { $_ -eq "VD_$($impactedDiscipline.FieldValues.Title.Replace(' ',''))" }
				if ($disciplineField.Count -eq 1 -and $commentsItem.FieldValues[$disciplineField] -like 'To Do') {
					#Se il valore della colonna disciplina è TO DO è necessario correggere il valore nel report
					Write-Host "$fullTcmCode - Update Comment status for discipline $($impactedDiscipline.FieldValues.Title)" -f Yellow
					$hideMessage = Set-PnPListItem -List 'Comment Status Report' -Identity $commentsItem.Id -Values @{$disciplineField = $taskClosureDate } -UpdateType SystemUpdate
					$updatedComments++
				}

		 }
		 #Rimozione effettiva del task
		 Write-Host "$fullTcmCode - Remove task. Closed on $taskClosureDate" -f Yellow
		 Remove-PnPPlannerTask -Task $_.Id
		 $removedTasks++
		}
		$processedTasks++
		$processedTasksForDiscipline++
	}
	$processedDisciplines++
}
Write-Host "Processed Disciplines $processedDisciplines Removed tasks $removedTasks, Updated Comments $updatedComments" -f Green