<#
    Questo script serve a modifica il percorso di un documento.
    Richiede un csv con l'attuale percorso del documento (DocPath) e il percorso del documento corretto (DocPathCalc)

    TODO:
    - At ($null -eq $DocumentLibrary), log and continue instead of Exit
    - In case of duplicates on List, return all IDs
    - Sub-ProgressBar
#>

param(
    [parameter(Mandatory = $true)][string]$SiteCode,
    [Switch]$Client
)

#Funzione di log to CSV
function Write-Log
{
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $SiteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath))
    {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID;TCM_DN;Rev;Action'
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
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Doc: ', ';').Replace(' - File: ', ';').Replace(' - ID: ', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

# Funzione che carica la Client Department Code Mapping
function Get-CDCM
{
    param (
        [Parameter(Mandatory = $true)]$SiteConn
    )

    Write-Log "Caricamento '$($CDCM)'..."
    $items = Get-PnPListItem -List $CDCM -PageSize 5000 -Connection $SiteConn | ForEach-Object {
        [PSCustomObject] @{
            ID       = $_['ID']
            Title    = $_['Title']
            Value    = $_['Value']
            ListPath = $_['ListPath']
        }
    }
    Write-Log 'Caricamento lista completato.'

    Return $items
}

# Funzione che verifica la correttezza del Client Discipline
function Find-ClientDiscipline
{
    param (
        [Parameter(Mandatory = $true)][String]$ClientCode,
        [Parameter(Mandatory = $true)][Array]$List
    )

    $ccSplit = $ClientCode.Split('-')
    if ($ClientCode.toUpper().StartsWith('DS'))
    {
        if ($ccSplit.Length -ge 4)
        {
            $tempCode = $ccSplit[2] + '-' + $ccSplit[3]
            [Array]$found = $List | Where-Object -FilterScript { $_.Title -eq $tempCode }
        }
    }
    else
    {
        if ($ccSplit.Length -ge 2)
        {
            $tempCode = $ccSplit[1]
            if ($tempCode -eq 'CR') { [Array]$found = $List | Where-Object -FilterScript { $_.ListPath -eq 'CSR' } }
            else { [Array]$found = $List | Where-Object -FilterScript { $_.ListPath -eq $tempCode } }
        }
    }
    if ($found.Count -eq 0) { [Array]$found = $List | Where-Object -FilterScript { $_.ListPath -eq 'NA' } }
    return $found[0]
}

$tcmUrl = "https://tecnimont.sharepoint.com/sites/$($SiteCode)DigitalDocuments"
$clientUrl = "https://tecnimont.sharepoint.com/sites/$($SiteCode)DigitalDocumentsC"
$CDCM = 'ClientDepartmentCodeMapping'

$Client ? ( $listName = 'Client Document List' ) : ( $listName = 'DocumentList' ) | Out-Null

# Caricamento CSV/Documento/Tutta la lista
$CSVPath = Read-Host -Prompt 'CSV Path o TCM Document Number'
if ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
elseif ($CSVPath -ne '')
{
    $csv = [PSCustomObject]@{
        TCM_DN = $CSVPath
        Rev    = $rev
        Count  = 1
    }
}
else
{
    Write-Host 'MODE: ALL LIST' -ForegroundColor Red
    Pause
}

$tcmConn = Connect-PnPOnline -Url $tcmUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
$mainConn = $tcmConn
if ($Client)
{
    $clientConn = Connect-PnPOnline -Url $ClientUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
    $mainConn = $clientConn
}
$listsArray = Get-PnPList -Connection $mainConn

# Legge tutta la document library
Write-Log "Caricamento '$($listName)'..."
$listItems = Get-PnPListItem -List $listName -PageSize 5000 -Connection $mainConn | ForEach-Object {
    if ($Client)
    {
        [PSCustomObject]@{
            ID         = $_['ID']
            TCM_DN     = $_['Title']
            ClientCode = $_['ClientCode']
            Rev        = $_['IssueIndex']
            CDC        = $_['ClientDepartmentCode']
            CDD        = $_['ClientDepartmentDescription']
            ID_DL      = $_['IDDocumentList']
            SrcEnv     = $_['DD_SourceEnvironment']
            DocPath    = $_['DocumentsPath']
        }
    }
    else
    {
        [PSCustomObject]@{
            ID              = $_['ID']
            TCM_DN          = $_['Title']
            ClientCode      = $_['ClientCode']
            Rev             = $_['IssueIndex']
            Status          = $_['DocumentStatus']
            DeptCode        = $_['DepartmentCode']
            DocClass        = $_['DocumentClassification']
            DocType         = $_['DocumentTypology']
            ClientDiscpline = $_['ClientDiscipline_Calculated']
            DocPath         = $_['DocumentsPath']
        }
    }
}
Write-Log 'Caricamento lista completato.'

# Caricamento Client Department Code Mapping
$CDCMItems = Get-CDCM -SiteConn $tcmConn

# Filtro per tutta la lista
if ($CSVPath -eq '') { $csv = $listItems }
#if ($CSVPath -eq "") { $csv = $listItems | Where-Object -FilterScript { $_.CDC -eq "NA" } }

$rowCounter = 0
$notExistCount = 0
$doneDocs = @()
Write-Log 'Inizio controllo...'
ForEach ($row in $csv)
{
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Verifica' -Status "$($rowCounter+1)/$($csv.Count) - $($row.TCM_DN)/$($row.Rev)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }

    if ($row.TCM_DN -in $doneDocs) { continue }
    [Array]$issues = $listItems | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN }

    $doneDocs += $row.TCM_DN

    if ($null -eq $issues) { $msg = "[ERROR] - List: $($ListName) - Doc: $($row.TCM_DN) - NOT FOUND" }
    else
    {
        ForEach ($item in $issues)
        {
            Write-Host "Documento: $($row.TCM_DN)/$($item.Rev)" -ForegroundColor Blue
            $docSplit = $item.TCM_DN.Split('-')
            $ccSplit = $item.ClientCode.Split('-')
            $clientDiscipline_Calc = Find-ClientDiscipline -ClientCode $item.ClientCode -List $CDCMItems
            $pathSplit = $item.DocPath.Split('/')
            $issueRelPath = $pathSplit[5..$pathSplit.Length] -join '/'
            $DLRelPath = '/' + ($pathSplit[3..5] -join '/')
            $DL = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath }

            if ($Client)
            {
                $folderRelPath_Calc = "$($clientDiscipline_Calc.ListPath)/$($docSplit[1][1])/$($docSplit[2])/$($item.ClientCode)"
                $issueRelPath_Calc = "$($folderRelPath_Calc)/$($item.Rev)"
                $docPath_Calc = ($pathSplit[0..4] -join '/') + "/$($issueRelPath_Calc)"
                $DLRelPath_Calc = "/$($pathSplit[3])/$($pathSplit[4])/$($clientDiscipline_Calc.ListPath)"
                $DL_Calc = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath_Calc }

                if ($item.CDC -eq $clientDiscipline_Calc.ListPath)
                {
                    <#
                    try {
                        Set-PnPListItem -List $listName -Identity $item.ID -Values @{
                            ClientDepartmentDescription = $clientDiscipline_Calc.Value
                        } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
                        Write-Log "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - MATCH Client Discipline"
                    }
                    catch { Write-Log "[ERROR] - List: DocumentList - Doc: $($row.TCM_DN)/$($item.Rev) - FAILED $($_)" }
                    #>
                    continue
                }
                else
                {
                    Write-Log "[WARNING] - List: $($listName) - Doc: $($item.TCM_DN)/$($item.Rev) - MISMATCH Client Discipline"
                    Write-Log 'Correzione liste in corso...'

                    if ($item.SrcEnv -eq 'DigitalDocuments')
                    {
                        try
                        {
                            Set-PnPListItem -List 'DocumentList' -Identity $item.ID_DL -Values @{
                                ClientDiscipline_Calculated = $clientDiscipline_Calc.Value
                            } -UpdateType SystemUpdate -Connection $tcmConn | Out-Null
                            Write-Log "[SUCCESS] - List: DocumentList - Doc: $($row.TCM_DN)/$($item.Rev) - UPDATED"
                        }
                        catch { Write-Log "[ERROR] - List: DocumentList - Doc: $($row.TCM_DN)/$($item.Rev) - FAILED" }
                    }

                    try
                    {
                        Set-PnPListItem -List $listName -Identity $item.ID -Values @{
                            ClientDepartmentCode        = $clientDiscipline_Calc.ListPath
                            ClientDepartmentDescription = $clientDiscipline_Calc.Value
                            DocumentsPath               = $docPath_Calc
                        } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
                        Write-Log "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - UPDATED"
                    }
                    catch { Write-Log "[ERROR] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - FAILED" }
                }
            }
            else
            {
                $folderRelPath_Calc = "$($docSplit[1][0])/$($docSplit[1][1])/$($docSplit[2])/$($item.TCM_DN)"
                $issueRelPath_Calc = "$($folderRelPath_Calc)/$($item.Rev)"
                $docPath_Calc = ($pathSplit[0..4] -join '/') + "/$($issueRelPath_Calc)"
                $DLRelPath_Calc = "/$($pathSplit[3])/$($pathSplit[4])/$($docSplit[1][0])"
                $DL_Calc = $listsArray | Where-Object -FilterScript { $_.RootFolder.ServerRelativeUrl -eq $DLRelPath_Calc }

                # Controllo Department Code
                if ($item.DeptCode -eq $docSplit[1][0]) { $msg = "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - MATCH Department Code" }
                else { $msg = "[ERROR] - List: $($listName) - $($row.TCM_DN) - $($item.Rev) - MISMATCH Department Code" }
                Write-Log $msg

                # Controllo Document Classification
                if ($item.DocClass -eq $docSplit[1]) { $msg = "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - MATCH Document Classification" }
                else { $msg = "[ERROR] - List: $($listName) - $($row.TCM_DN) - $($item.Rev) - MISMATCH Document Classification" }
                Write-Log $msg

                # Controllo Document Typology
                if ($item.DocType -eq $docSplit[2]) { $msg = "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - MATCH Document Typology" }
                else { $msg = "[ERROR] - List: $($listName) - $($row.TCM_DN) - $($item.Rev) - MISMATCH Document Typology" }
                Write-Log $msg

                # Controllo coerenza Documents Path e correzione
                if ($item.DocPath -eq $docPath_Calc)
                {
                    if (-not (Get-PnPFolder -Url $item.DocPath -ErrorAction SilentlyContinue))
                    {
                        $notExistCount++
                        Resolve-PnPFolder -SiteRelativePath $folderRelPath_Calc -Connection $mainConn | Out-Null
                    }
                    Write-Log "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - MATCH DocumentsPath"
                    continue
                }
                else
                {
                    try
                    {
                        Set-PnPListItem -List $listName -Identity $item.ID -Values @{
                            DepartmentCode         = $docSplit[1][0]
                            DocumentClassification = $docSplit[1]
                            DocumentTypology       = $docSplit[2]
                            DocumentsPath          = $docPath_Calc
                        } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
                        Write-Log "[SUCCESS] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - UPDATED"
                    }
                    catch { Write-Log "[ERROR] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - FAILED" }
                }
            }

            Write-Log "[WARNING] - List: $($listName) - Doc: $($row.TCM_DN)/$($item.Rev) - MISMATCH DocumentsPath"
            Write-Log 'Correzione cartella in corso...'

            Resolve-PnPFolder -SiteRelativePath $folderRelPath_Calc -Connection $mainConn | Out-Null

            try
            {
                Move-PnPFolder -Folder $issueRelPath -TargetFolder $folderRelPath_Calc -Connection $mainConn | Out-Null
                $msg = "[SUCCESS] - List: $($DL_Calc.Title) - Doc: $($row.TCM_DN)/$($item.Rev) - UPDATED DocumentsPath"
            }
            catch
            {
                Write-Host "[WARNING] - List: $($DL.Title) - Doc: $($row.TCM_DN)/$($item.Rev) - SKIPPED" -ForegroundColor Yellow
                Start-Process (($pathSplit[0..4] -join '/') + "/$($folderRelPath_Calc)")
                Pause
                Move-PnPFolder -Folder $issueRelPath -TargetFolder $folderRelPath_Calc -Connection $mainConn | Out-Null
                $msg = "[SUCCESS] - List: $($DL_Calc.Title) - Doc: $($row.TCM_DN)/$($item.Rev) - UPDATED DocumentsPath"
            }
            Write-Log $msg

            if ($Client)
            {
                # Aggiorno la Revisione sui file
                do
                {
                    $connError = $false
                    try
                    {
                        $fileList = Get-PnPFolderItem -FolderSiteRelativeUrl $issueRelPath_Calc -ItemType File -Recursive -Connection $mainConn | Select-Object Name, ServerRelativeUrl

                        if ($null -eq $fileList) { Write-Log "[WARNING] - List: $($DL.Title) - Doc: $($row.TCM_DN)/$($item.Rev) - EMPTY OR NOT FOUND" }
                        else
                        {
                            foreach ($file in $fileList)
                            {
                                $fileAtt = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem -Connection $mainConn
                                try
                                {
                                    Set-PnPListItem -List $DL_Calc.Title -Identity $fileAtt.Id -Values @{
                                        ClientDepartmentCode        = $clientDiscipline_Calc.ListPath
                                        ClientDepartmentDescription = $clientDiscipline_Calc.Value
                                        DD_SourceEnvironment        = $item.SrcEnv
                                    } -UpdateType SystemUpdate -Connection $mainConn | Out-Null
                                    Write-Log "[SUCCESS] - List: $($DL_Calc.Title) - File: $($file.Name) - UPDATED"
                                }
                                catch { Write-Log "[ERROR] - List: $($DL_Calc.Title) - File: $($file.Name) - FAILED $($_)" }
                            }
                        }
                    }
                    catch
                    {
                        Write-Host "[ERROR] - $($_)"
                        $connError = $true
                    }
                }
                while ($connError)
            }
        }
    }
}
if ($csv.Count -gt 1) { Write-Progress -Activity 'Verifica' -Completed }
Write-Log 'Controllo terminato.'
Write-Log "Path mancanti creati: $($notExistCount) su $($rowCounter)."