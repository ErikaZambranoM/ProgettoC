param(
	[Parameter(Mandatory = $true)][string]$ProjectCode
)

function Write-Log {
	param (
		[Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
		[String]$Code = $ProjectCode
	)

	$ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
	$logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

	if (!(Test-Path -Path $logPath)) {
		$newLog = New-Item $logPath -Force -ItemType File
		Add-Content $newLog "Timestamp;Type;ListName;TCM_DN;Rev;Action;Value"
	}
	$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

	if ($Message.Contains("[SUCCESS]")) {
		Write-Host $Message -ForegroundColor Green
	}
	elseif ($Message.Contains("[ERROR]")) {
		Write-Host $Message -ForegroundColor Red
	}
	elseif ($Message.Contains("[WARNING]")) {
		Write-Host $Message -ForegroundColor Yellow
	}
	else {
		Write-Host $Message -ForegroundColor Cyan
		return
	}
	$Message = $Message.Replace(" - List: ", ";").Replace(" - TCM_DN: ", ";").Replace(" - Rev: ", ";").Replace(" - DeptCode: ", ";").Replace(" - Desc: ", ";").Replace(" - ", ";").Replace(": ", ";")
	Add-Content $logPath "$FormattedDate;$Message"
}

$SiteURL = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
$SiteURLClient = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"

$DDList = "DocumentList"
$CDList = "Client Document List"
$CDCMList = "ClientDepartmentCodeMapping"

$TCMConn = Connect-PnPOnline -Url $SiteURL -UseWebLogin -ReturnConnection -WarningAction SilentlyContinue
$ClientConn = Connect-PnPOnline -Url $SiteURLClient -UseWebLogin -ReturnConnection -WarningAction SilentlyContinue

# Caricamento DD lato TCM
Write-Host "Caricamento '$($DDList)'..." -ForegroundColor Cyan
$DDItems = Get-PnPListItem -List $DDList -PageSize 5000 -Connection $TCMConn | ForEach-Object {
	[PSCustomObject]@{
		ID              = $_["ID"]
		TCM_DN          = $_["Title"]
		Rev             = $_["IssueIndex"]
		DeptCode        = $_["DepartmentCode"]
		DeptCode4Struct = $_["DepartmentForStructure_Calc"]
	}
}
Write-Log "Caricamento lista completato."

# Caricamento Client Department Code
Write-Host "Caricamento '$($CDCMList)'..." -ForegroundColor Cyan
$CDCMItems = Get-PnPListItem -List $CDCMList -PageSize 5000 -Connection $TCMConn | ForEach-Object {
	[PSCustomObject]@{
		ID       = $_["ID"]
		DeptCode = $_["Title"]
		Desc     = $_["Value"]
	}
}
Write-Log "Caricamento lista completato."

# Caricamento DD lato Client
Write-Host "Caricamento '$($CDList)'..." -ForegroundColor Cyan
$CDLItems = Get-PnPListItem -List $CDList -PageSize 5000 -Connection $ClientConn | ForEach-Object {
	[PSCustomObject]@{
		ID             = $_["ID"]
		TCM_DN         = $_["Title"]
		Rev            = $_["IssueIndex"]
		DL_ID          = $_["IDDocumentList"]
		ClientDeptCode = $_["ClientDepartmentCode"]
	}
}
Write-Log "Caricamento lista completato."

$itemCounter = 0
Write-Log "Inizio correzione..."
ForEach ($item in $CDLItems) {
	Write-Progress -Activity "Elaborazione" -Status "$($itemCounter)/$($CDLItems.Count) - $($item.TCM_DN) - $($item.Rev)" -PercentComplete (($itemCounter++ / $CDLItems.Count) * 100)

	$DLItem = $DDItems | Where-Object -FilterScript { $_.ID -eq $item.DL_ID }
	$DeptMap = $CDCMItems | Where-Object -FilterScript { $_.DeptCode -eq $DLItem.DeptCode4Struct }

	if ($null -eq $DLItem) { $msg = "[WARNING] - List: $($CDList) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - NOT FOUND" }
	elseif ($null -eq $DeptMap) { $msg = "[WARNING] - List: $($CDCMList) - DeptCode: $($DLItem.DeptCode4Struct) - Desc: NOT FOUND - MAPPING NOT FOUND" }
	elseif ($DeptMap.DeptCode -eq $item.ClientDeptCode) { continue }
	else {
		try {
			Set-PnPListItem -List $CDList -Identity $item.ID -Values @{
				"ClientDepartmentCode"        = $DeptMap.DeptCode
				"ClientDepartmentDescription" = $DeptMap.Desc
			} -UpdateType SystemUpdate -Connection $ClientConn | Out-Null
			$msg = "[SUCCESS] - List: $($CDList) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - UPDATED - DeptCode: $($DeptMap.DeptCode)"
		}
		catch { $msg = "[ERROR] - List: $($CDList) - TCM_DN: $($item.TCM_DN) - Rev: $($item.Rev) - FAILED" }
	}
	Write-Log -Message $msg
}