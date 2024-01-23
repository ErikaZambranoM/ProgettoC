# Questo script può essere lanciato sia su Disciplines (DD lato TCM) e DDDisciplines (VDM)

param (
	[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$SiteUrl,
	[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$CSVPath
)

# Funzione di log to CSV
function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $siteCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog "Timestamp;Type;ListName;DeptCode;Action"
	}
	$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

	if ($Message.Contains("[SUCCESS]")) { Write-Host $Message -ForegroundColor Green }
	elseif ($Message.Contains("[ERROR]")) { Write-Host $Message -ForegroundColor Red }
	elseif ($Message.Contains("[WARNING]")) { Write-Host $Message -ForegroundColor Yellow }
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(" - List: ", ";").Replace(" - DeptCode: ", ";")
	Add-Content $logPath "$FormattedDate;$Message"
}

try {
	if ($SiteUrl.ToLower().Contains("vdm_")) { $listName = "DDDisciplines" }
	else { $listName = "Disciplines" }

	$csv = import-csv $CSVPath -Delimiter ";"
	$validCols = @("Title", "DepartmentCode", "Recipients", "CCRecipients")
	$validCounter = 0
	($csv | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
		if ($_ -in $validCols) { $validCounter++ }
	}
	if ($validCounter -lt $validCols.Count) {
		Write-Host "Colonne obbligatorie mancanti: $($validCols -join ', ')" -ForegroundColor Red
		Exit
	}

	Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop
	$siteCode = (Get-PnPWeb).Title.Split(" ")[0]
	$siteUsers = Get-PnPUser

	# Legge tutta la document library
	Write-Log "Caricamento '$($listName)'..."
	$listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
		[PSCustomObject] @{
			ID       = $_["ID"]
			DeptCode = $_["DepartmentCode"]
		}
	}
	Write-Log "Caricamento lista completato."

	$rowCounter = 0
	Write-Log "Inizio aggiornamento..."
	foreach ($row in $csv) {
		if ($csv.Count -gt 1) { Write-Progress -Activity "Aggiornamento" -Status "$($rowCounter+1)/$($csv.Count) - $($row.DepartmentCode)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

		$item = $listItems | Where-Object -FilterScript { $_.DeptCode -eq $row.DepartmentCode }

		if ($null -ne $row.Reviewers) {
			$reviewers = $row.Reviewers.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries)

			$arrReviewers = @()
			for ($j = 0; $j -lt $reviewers.Length; $j++ ) {
				$email = $reviewers[$j]
				$user = $siteUsers | Where-Object { $_.Email -eq $email }
				#If($null -eq $user) {
				#  #$user = New-PnPUser -LoginName $reviewers[$j]
				#  if(-not $arrInvalidUser.Contains($email)){
				#	  $UserItem = New-Object PSObject
				#	  $UserItem | Add-Member -MemberType NoteProperty -name "Mail" -value $email
				#	  $arrInvalidUser += $UserItem
				#  }
				#}
				If ($null -ne $user) {
					if (-not $arrReviewers.Contains($email)) {
						$arrReviewers += $user.ID
					}
				}
			}
		}

		if ($null -ne $row.Informed) {
			$informed = $row.Informed.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries)

			$arrInformed = @()
			for ($k = 0; $k -lt $informed.Length; $k++ ) {
				$email = $informed[$k]
				$user = $siteUsers | Where-Object { $_.Email -eq $email }
				#if($null -eq $user) {
				#  #$user = New-PnPUser -LoginName $informed[$j]
				#  if(-not $arrInvalidUser.Contains($email)){
				#	  $UserItem = New-Object PSObject
				#	  $UserItem | Add-Member -MemberType NoteProperty -name "Mail" -value $email
				#	  $arrInvalidUser += $UserItem
				#  }
				#}
				If ($null -ne $user) {
					if (-not $arrReviewers.Contains($email)) {
						$arrInformed += $user.ID
					}
				}
			}
		}

		if ($null -ne $row.Consolidators) {
			$consolidators = $row.Consolidators.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries)

			$arrConsolidators = @()
			for ($y = 0; $y -lt $consolidators.Length; $y++ ) {
				$email = $consolidators[$y]
				$user = $siteUsers | Where-Object { $_.Email -eq $email }
				#If($null -eq $user) {
				#  #$user = New-PnPUser -LoginName $consolidators[$j]
				#  if(-not $arrInvalidUser.Contains($email)){
				#	  $UserItem = New-Object PSObject
				#	  $UserItem | Add-Member -MemberType NoteProperty -name "Mail" -value $email
				#	  $arrInvalidUser += $UserItem
				#  }
				#}
				If ($null -ne $user) {
					if (-not $arrReviewers.Contains($email)) {
						$arrConsolidators += $user.ID
					}
				}
			}
		}

		if ($null -ne $row.Recipients) {
			$recipients = $($row.Recipients).Replace("`n", "").Replace(" ", "")
		}

		if ($null -ne $row.CCRecipients) {
			$ccRecipients = $($row.CCRecipients).Replace("`n", "").Replace(" ", "")
		}

		$values = @{
			"Title"          = $row.Title;
			"DepartmentCode" = $row.DepartmentCode;
			"Recipients"     = $recipients;
			"CCRecipients"   = $ccRecipients;
			"Reviewers"      = $reviewers;
			"Consolidators"  = $consolidators;
			"Informed"       = $informed;
		}

		try {
			if ($null -eq $item) {
				Add-PnPListItem -List $listName -Values $values | Out-Null
				Write-Log "[SUCCESS] - List: $($listName) - DeptCode: $($row.DepartmentCode) - CREATED"
			}
			else {
				Set-PnPListItem -List $listName -Identity $item.ID -Values $values -UpdateType SystemUpdate | Out-Null
				Write-Log "[SUCCESS] - List: $($listName) - DeptCode: $($row.DepartmentCode) - UPDATED"
			}
		}
		catch { Write-Log "[ERROR] - List: $($listName) - DeptCode: $($row.DepartmentCode) - FAILED" }
	}
	Write-Log "Operazione completata."
}
catch { Throw }
finally { if ($csv.Count -gt 1) { Write-Progress -Activity "Aggiornamento" -Completed } }