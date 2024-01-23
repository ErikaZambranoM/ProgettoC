##############################################################
## UPDATE FOLLOWING FIELDS IN VDL LIST:         	        ##
##  - PO DATE                                       	    ##
##  - CLIENT ACCEPTANCE STATUS                  	        ##
##  - MILESTONE SET (ONLY IF EnableMilestoneSetCalculation) ##
##                                              	        ##
## AUTHOR: MIRCO FOCARETE                                   ##
## LAST MOD. DATE: 14, November 2023                        ##
## LAST MOD. BY: FEDERICO BARONE                	        ##
##############################################################

# NOTE: Provided CSV needs to have a column named 'SiteURL' with the site url of each site to process

#Config Variables
$PMCListName = 'SAP Purchase Order List'

function GetMilestoneSet {
    Param ($VDM_MS_ListItems, $VD_DisciplineOwnerTCM, $DocumentTypology)

    $ListItems = @(($VDM_MS_ListItems | Where-Object -FilterScript { $_.VD_DisciplineOwnerTCM -eq $VD_DisciplineOwnerTCM -and $_.DocumentTypology -eq $DocumentTypology }).FieldValues | Select-Object -First 1)

    $MilestoneSet = $null

    foreach ($ListItem in $ListItems) {
        $MilestoneSet = $ListItem['VD_MilestoneSet']
    }

    return $MilestoneSet
}

function GetPORecord {
    Param ($PMC_ListItems, $PONumber)

    $ListItems = @(($PMC_ListItems | Where-Object -FilterScript { $_['Title'] -eq $PONumber -and $null -ne $_['PO_Date'] } | Sort-Object -Property PO_Version -Descending).FieldValues | Select-Object -First 1)

    $PODate = $null
    foreach ($ListItem in $ListItems) {
        $PODate = $ListItem['PO_Date']
    }

    return $PODate
}

function GetApprovalResult {
    Param ($CDL_ListItems, $VDLID)
    $ListItems = @(($CDL_ListItems | Where-Object -FilterScript { $_['IDDocumentList'] -eq $VDLID -and $_['DD_SourceEnvironment'] -eq 'VendorDocuments' }).FieldValues | Select-Object -First 1 )

    $ApprovalResult = $null
    foreach ($ListItem in $ListItems) {
        $ApprovalResult = $ListItem['ApprovalResult']
    }

    return $ApprovalResult
}

try {
    # Caricamento CSV o sito singolo
    $CSVPath = (Read-Host -Prompt 'CSV Path o Site Url').Trim('"').Trim("'").Trim('/')
    if ($CSVPath.ToLower().Contains('.csv')) { $SiteCsv = Import-Csv -Path $CSVPath -Delimiter ';' }
    elseif ($CSVPath -ne '') {
        $SiteCsv = @([PSCustomObject]@{
                SiteURL = $CSVPath
            })
    }
    else { Throw 'No CSV or Site Url provided' }

    # Create Logs folder if not exists and start transcript
    if (-not (Test-Path -Path "$PSScriptRoot\Logs" -PathType Container)) {
        New-Item -Path "$PSScriptRoot\Logs" -ItemType Directory | Out-Null
    }
    $ScriptRunDateTime = Get-Date -Format 'dd-MM-yyyy_HH-mm-ss'
    $ScriptName = (Get-Item -Path $MyInvocation.MyCommand.Path).BaseName
    Start-Transcript -Path "$PSScriptRoot\Logs\$($ScriptName)_$($ScriptRunDateTime).log" -IncludeInvocationHeader

    # Loop on each site
    $SitesCounter = 0
    foreach ($VDMSiteURL in $SiteCsv.SiteURL) {
        $SitesCounter++
        $PONumber = $null
        $UpdateArray = @()
        Write-Host "Processing site: $($VDMSiteURL.Split('/')[-1]) ($SitesCounter / $($SiteCsv.Count))"

        # Return error if site is not a VDM site
        if (!($VDMSiteURL.ToLower().Contains('vdm'))) {
            throw 'Site URL is not a VDM Site'
        }

        $Connection = Connect-PnPOnline -Url $VDMSiteURL -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop

        # Get required configuration values
        $Config_ListItems = @((Get-PnPListItem -List 'Configuration List' -PageSize 50 -Connection $Connection).FieldValues | Where-Object { $_['Title'] -eq 'PMCSiteUrl' -or $_['Title'] -eq 'EnableMilestoneSetCalculation' } )
        foreach ($ListItem in $Config_ListItems) {
            switch ($ListItem['Title']) {
                'PMCSiteUrl' { $PMCSiteURL = $ListItem['VD_ConfigValue'] }
                'EnableMilestoneSetCalculation' { $EnableMilestoneSetCalculation = $ListItem['VD_ConfigValue'] }
            }
        }
        $Settings_ListItems = @((Get-PnPListItem -List 'Settings' -PageSize 50 -Connection $Connection).FieldValues | Where-Object { $_['Title'] -eq 'ClientSiteUrl' } | Select-Object -First 1  )
        foreach ($ListItem in $Settings_ListItems) {
            $ClientSiteURL = $ListItem['Value']
        }

        # Retrieve all items from involved lists
        $VDL_ListItems = @(Get-PnPListItem -List 'Vendor Documents List' -Fields ID, VD_DocumentNumber, VD_Index, VD_PONumber, VD_DisciplineOwnerTCM, VD_DocumentType -PageSize 5000 -Connection $Connection | Sort-Object -Property VD_PONumber).FieldValues
        $TotItems = $VDL_ListItems.Count
        $VDM_MS_ListItems = @(Get-PnPListItem -List 'VDM Milestone Set' -PageSize 5000 -Connection $Connection)
        $ConnectionSub = Connect-PnPOnline -Url $PMCSiteURL -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop
        $ConnectionDD = Connect-PnPOnline -Url $ClientSiteURL -UseWebLogin -ValidateConnection -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop
        $PMC_ListItems = @(Get-PnPListItem -List $PMCListName -PageSize 5000 -Connection $ConnectionSub)
        $CDL_ListItems = @(Get-PnPListItem -List 'Client Document List' -PageSize 5000 -Connection $ConnectionDD)

        # Loop on each item from Vendor Documents List
        $I = 0
        foreach ($VDL_ListItem in $VDL_ListItems) {
            $Id = $VDL_ListItem['ID']
            $VD_DocumentNumber = $VDL_ListItem['VD_DocumentNumber'] + '-' + $VDL_ListItem['VD_Index'].ToString('000')
            $MilestoneSetReturn = $null
            $I += 1
            $ItemCount = '(Record ' + $I + ' of ' + $TotItems + ')'

            if ($PONumber -ne $VDL_ListItem['VD_PONumber']) {
                $PONumber = $VDL_ListItem['VD_PONumber']
                $DataReturn = GetPORecord $PMC_ListItems $PONumber
            }

            $ApprovalResultReturn = GetApprovalResult $CDL_ListItems $Id

            if ($EnableMilestoneSetCalculation -eq 'Yes') {
                $MilestoneSetReturn = GetMilestoneSet $VDM_MS_ListItems $VDL_ListItem['VD_DisciplineOwnerTCM'].LookupValue $VDL_ListItem['VD_DocumentType']
            }

            if ($null -ne $DataReturn -or $null -ne $ApprovalResultReturn -or $null -ne $MilestoneSetReturn ) {
                $data = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Write-Host '[' $data '] Data found for Revision' $VD_DocumentNumber '- ID:' $Id $ItemCount -ForegroundColor Green
                $UpdObject = [pscustomobject]@{
                    Id                = $Id
                    VD_DocumentNumber = $VD_DocumentNumber
                    PO_Date           = $DataReturn
                    ApprovalResult    = $ApprovalResultReturn
                    VD_MilestoneSet   = $MilestoneSetReturn
                }
                $UpdateArray += $UpdObject
            }
            else {
                $data = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Write-Host '[' $data '] No Data found for Revision' $VD_DocumentNumber '- ID:' $Id $ItemCount -ForegroundColor Red
            }
        }

        if ($UpdateArray.Count -gt 0) {
            Connect-PnPOnline -Url $VDMSiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

            $batch = New-PnPBatch
            foreach ($Item in $UpdateArray) {

                $Id = $Item.Id
                $VD_DocumentNumber = $Item.VD_DocumentNumber
                $ApprovalResult = $Item.ApprovalResult
                $PO_Date = $Item.PO_Date
                $VD_MilestoneSet = $Item.VD_MilestoneSet

                if ($EnableMilestoneSetCalculation -eq 'Yes') {
                    $UpdItem = Set-PnPListItem -List 'Vendor Documents List' -Identity $Id -Values @{'VD_PODate' = $PO_Date; 'ApprovalResult' = $ApprovalResult; 'VD_MilestoneSet' = $VD_MilestoneSet } -Batch $batch -UpdateType SystemUpdate
                }
                else {
                    $UpdItem = Set-PnPListItem -List 'Vendor Documents List' -Identity $Id -Values @{'VD_PODate' = $PO_Date; 'ApprovalResult' = $ApprovalResult } -Batch $batch -UpdateType SystemUpdate
                }

                $data = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Write-Host '[' $data '] Updated Revision' $VD_DocumentNumber '- ID:' $Id -ForegroundColor DarkYellow
            }

            Invoke-PnPBatch -Batch $batch
            Disconnect-PnPOnline
        }
    }
    Stop-Transcript
}
catch {
    throw
    Stop-Transcript
}