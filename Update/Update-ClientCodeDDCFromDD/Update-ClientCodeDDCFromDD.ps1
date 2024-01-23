#Questo script va a popolare il ClientCode in DDC, prendendolo da DD e ricrea il documentPath
#Inoltre crea la struttura in area client della cartella importando il file presente in DD

# Connessione al sito
Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please insert the site Code')]
    [string]$ProjectCode,
    [Parameter(Mandatory = $true, HelpMessage = 'Path Cartella Temp')]
    [String]$PathTemp
)
function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $codiceSito
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;Level;Action'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Level: ', ';').Replace(' - ', ';').Replace(': ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

$sitoDD = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
$sitoDDC = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
$DL = 'DocumentList'
$CDL = 'Client Document List'

$CSVPath = (Read-Host -Prompt 'CSV Path o TCM Document Number').Trim('"')
If ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
Else {
    $rev = Read-Host -Prompt 'Issue Index'
    $csv = [PSCustomObject] @{
        TCM_DN = $CSVPath
        Rev    = $rev
        Count  = 1
    }
}

# Connessioni ai siti
$clientConn = Connect-PnPOnline -Url $sitoDDc -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue -ReturnConnection
$DDConn = Connect-PnPOnline -Url $sitoDD -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction SilentlyContinue -ReturnConnection

Write-Log "Caricamento '$($CDL)'..."
$ListaDDc = Get-PnPListItem -List $CDL -PageSize 5000 -Connection $clientConn | ForEach-Object {
    [PSCustomObject]@{
        ID             = $_['ID']
        TCM_DN         = $_['Title']
        Rev            = $_['IssueIndex']
        ClientCode     = $_['ClientCode']
        DocumentsPath  = $_['DocumentsPath']
        IDDocumentList = $_['IDDocumentList']
        CDC            = $_['ClientDepartmentCode']
        CDD            = $_['ClientDepartmentDescription']
    }
}
Write-Log 'Caricamento completato.'

Write-Log "Caricamento '$($DL)'..."
$ListaDD = Get-PnPListItem -List $DL -PageSize 5000 -Connection $DDConn | ForEach-Object {
    [PSCustomObject]@{
        ID                          = $_['ID']
        TCM_DN                      = $_['Title']
        ClientCode                  = $_['ClientCode']
        Rev                         = $_['IssueIndex']
        Index                       = $_['Index']
        ReasonForIssue              = $_['ReasonForIssue']
        IsCurrent                   = $_['isCurrent']
        DocTitle                    = $_['DocumentTitle']
        DocClass                    = $_['DocumentClass']
        DocumentsPath               = $_['DocumentsPath']
        LastTransmittal             = $_['LastTransmittal']
        LastTransmittalDate         = $_['LastTransmittalDate']
        CommentDueDate              = $_['CommentDueDate']
        ClientDepartmentCode        = $_['ClientDepartmentCode']
        ClientDepartmentDescription = $_['ClientDepartmentDescription']
    }
}
Write-Log 'Caricamento completato.'

#cerca nella lista DocumentList i documenti a cui manca in area Client il ClientCode, per poi settare il clientCode in DDc
foreach ($row in $csv) {

    # Filtro documento sulla CDL
    $itemCDL = $ListaDDc | Where-Object -FilterScript { $_.TCM_DN -eq $row.TCM_DN -and $_.Rev -eq $row.Rev }

    if ($null -eq $itemCDL) { Write-Log "[ERROR] - List: $($CDL) - Doc: $($row.TCM_DN)/$($row.Rev) - NOT FOUND" }
    else {
        Write-Log "Doc: $($itemCDL.TCM_DN)/$($itemCDL.Rev)"

        $itemDD = $ListaDD | Where-Object -FilterScript { $_.ID -eq $itemCDL.IDDocumentList }

        # Controllo se è necessario aggiornare il document path
        if ($itemCDL.DocumentsPath.Contains($itemDD.ClientCode)) { $docpath = $itemCDL.DocumentsPath }
        else { $docpath = $itemCDL.DocumentsPath.Replace("/$($itemDD.Rev)", "/$($itemDD.clientCode)/$($itemDD.Rev)") }

        $lastTrnDateConv = Get-Date $($itemDD.LastTransmittalDate) -Format 'MM/dd/yy'
        try { $commentDueDateConv = Get-Date $($itemDD.CommentDueDate) -Format 'MM/dd/yy' } catch { $commentDueDateConv = $null }

        #setta in CDL ClientCode e DocumentPath
        try {
            Set-PnPListItem -List $CDL -Identity $itemCDL.ID -Values @{
                ClientCode    = $itemDD.ClientCode
                DocumentsPath = $docpath
            } -UpdateType SystemUpdate -Connection $clientConn | Out-Null
            Write-Log "[SUCCESS] - List: $($CDL) - Doc: $($itemCDL.Title) - DocPath: $($docpath) - UPDATED"
        }
        catch { Write-Log "[ERROR] - List: $($CDL) - Doc: $($itemCDL.Title) - DocPath: $($docpath)  - UPDATE FAILED" }

        # crea la folder in Area Client e scarica i documenti in locale e li importa in DDC
        $Lengthsito = $sitoDD.Length
        $ClientDocPath = $itemCDL.DocumentsPath.remove(0, $Lengthsito + 2)
        $relativeClientDocPath = $ClientDocPath.Replace('/0', "/$($itemDD.ClientCode)/0")
        $attPath = "$($relativeClientDocPath)/Attachments"
        $relativePathDD = $itemDD.DocumentsPath.remove(0, $Lengthsito)

        # Creazione cartella Root e Attachment in Client Area
        Resolve-PnPFolder -SiteRelativePath $relativeClientDocPath -Connection $clientConn | Out-Null
        Resolve-PnPFolder -SiteRelativePath $attPath -Connection $clientConn | Out-Null
        if (!(Test-Path -Path $PathTemp)) { mkdir -Path $PathTemp | Out-Null }
        #Connect-PnPOnline -Url $sitoDD -UseWebLogin -ValidateConnection -ErrorAction Stop -WarningAction Continue
        $DDFolderDocs = Get-PnPFolderItem -FolderSiteRelativeUrl $relativePathDD -Recursive -ItemType File -Connection $DDConn

        ForEach ($file in $DDFolderDocs) {
            try {
                Get-PnPFile -Url $($file.ServerRelativeUrl) -AsFile -Path "$($PathTemp)" -Filename $file.Name -Force -Connection $DDConn | Out-Null
                Write-Log "[SUCCESS] - List: $($DL) - File: $($file.Name) - DOWNLOADED"
                $NewFileName = $File.Name -replace $($itemDD.TCM_DN), $($itemDD.ClientCode)
                $doc = "$($PathTemp)/$($File.Name)"
                if ($file.ServerRelativeUrl.Contains('Attachments')) {
                    Add-PnPFile -Path $doc -Folder "$relativeClientDocPath/Attachments" -NewFileName $NewFileName -Values @{
                        'IssueIndex'                  = $itemDD.Rev
                        'ReasonForIssue'              = $itemDD.ReasonForIssue
                        'ClientCode'                  = $itemDD.ClientCode
                        'IDDocumentList'              = $itemDD.ID
                        'Transmittal_x0020_Number'    = $itemDD.LastTransmittal
                        'TransmittalDate'             = $lastTrnDateConv
                        'CommentDueDate'              = $commentDueDateConv
                        'Index'                       = $itemDD.Index
                        'DocumentTitle'               = $itemDD.DocTitle
                        'DD_SourceEnvironment'        = 'DigitalDocuments'
                        'ClientDepartmentCode'        = $itemCDL.CDC
                        'ClientDepartmentDescription' = $itemCDL.CDD
                    } -Connection $clientConn | Out-Null
                }
                else {
                    Add-PnPFile -Path $doc -Folder $relativeClientDocPath -NewFileName $NewFileName -Values @{
                        'IssueIndex'                  = $itemDD.Rev
                        'ReasonForIssue'              = $itemDD.ReasonForIssue
                        'ClientCode'                  = $itemDD.ClientCode
                        'IDDocumentList'              = $itemDD.ID
                        'IDClientDocumentList'        = $itemCDL.ID
                        'CommentRequest'              = $true
                        'IsCurrent'                   = $true
                        'Transmittal_x0020_Number'    = $itemDD.LastTransmittal
                        'TransmittalDate'             = $lastTrnDateConv
                        'CommentDueDate'              = $commentDueDateConv
                        'Index'                       = $itemDD.Index
                        'DocumentTitle'               = $itemDD.DocTitle
                        'DD_SourceEnvironment'        = 'DigitalDocuments'
                        'ClientDepartmentCode'        = $itemCDL.CDC
                        'ClientDepartmentDescription' = $itemCDL.CDD
                    } -Connection $clientConn | Out-Null
                }
                Write-Log "[SUCCESS] - List: $($relativeClientDocPath[0]) - File: $($file.Name) - UPLOADED"
            }
            catch { throw }
        }
    }
}

Write-Log 'Operazione Completata'