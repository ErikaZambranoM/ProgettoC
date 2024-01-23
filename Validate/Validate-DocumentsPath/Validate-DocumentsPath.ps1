<#
    TODO:
    - Add logics for Reserved status
#>

param(
    [parameter(Mandatory = $true)][string]$SiteUrl #URL del sito
)


try
{
    if ($SiteUrl[-1] -eq '/')
    {
        $SiteUrl = $SiteUrl.TrimEnd('/')
    }

    if ($SiteUrl.ToLower().Contains('digitaldocumentsc'))
    {
        $listType = 'CD'
        $listName = '/Lists/ClientDocumentList'
    }
    elseif ($SiteUrl.ToLower().Contains('digitaldocuments'))
    {
        $listType = 'DD'
        $listName = '/Lists/DocumentList'
    }
    elseif ($SiteUrl.ToLower().Contains('vdm'))
    {
        $listType = 'VD'
        $listName = '/Lists/VendorDocumentsList'
    }
    elseif ($SiteUrl.ToLower().EndsWith('ddwave2'))
    {
        $listType = 'DD'
        $listName = '/Lists/DocumentList'
    }
    elseif ($SiteUrl.ToLower().EndsWith('ddwave2C'))
    {
        $listType = 'CD'
        $listName = '/Lists/ClientDocumentList'
    }
    else
    {
        Write-Host '[ERROR] URL non valido' -ForegroundColor Red
        exit
    }

    #Connessione al sito
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue
    $siteCode = (Get-PnPWeb).Title.Split(' ')[0]

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $filePath = "$($PSScriptRoot)\Log\$($siteCode)-$($listType)-$($ExecutionDate).csv";
    if (!(Test-Path -Path $filePath)) { New-Item $filePath -Force -ItemType File | Out-Null }

    Write-Host "Caricamento '$($listName.Split('/')[-1])'..." -ForegroundColor Cyan
    $listItems = Get-PnPListItem -List $listName -PageSize 5000 | ForEach-Object {
        if ($listType -eq 'DD' -or $listType -eq 'CD')
        {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID          = $($_['ID'])
                TCM_DN      = $($_['Title'])
                Rev         = $($_['IssueIndex'])
                DocPath     = $($_['DocumentsPath'])
                Dep_Code    = $($_['DepartmentCode'])
                Doc_Class   = $($_['DocumentClassification'])
                Doc_Type    = $($_['DocumentTypology'])
                DocPathCalc = ''
                IsCalcPath  = ''
                IsDepCode   = ''
                IsDocClass  = ''
                IsDocType   = ''
                PathExists  = ''
            }
        }
        elseif ($listType -eq 'VD')
        {
            $item = New-Object -TypeName PSCustomObject -Property @{
                ID      = $($_['ID'])
                TCM_DN  = $($_['VD_DocumentNumber'])
                Rev     = $($_['VD_RevisionNumber'])
                DocPath = $($_['VD_DocumentsPath'])
            }
        }
        $item
    }

    Write-Host 'Inizio controllo in corso...' -ForegroundColor Magenta
    $counter = 0
    foreach ($item in $listItems)
    {

        $counter++
        Write-Progress -Activity 'Controllo valori' -Status "Processing Item number $($counter)/$($listItems.Count),  $($item.TCM_DN) Rev.$($item.Rev)" -PercentComplete (($counter / $($listItems.Count)) * 100)
        $DepCodeCalc = $item.TCM_DN[5]
        $DocClassCalc = $item.TCM_DN[6]
        $DocTypeCalc = $item.TCM_DN.Substring(8, 2)

        $DocPathCalc = ($SiteUrl, $DepCodeCalc, $DocClassCalc, $DocTypeCalc, $item.TCM_DN, $item.Rev) -join '/'

        If ($item.DocPath)
        {
            $DocFolder = Get-PnPFolder -Url $item.DocPath -ErrorAction SilentlyContinue
            if ($null -eq $DocFolder)
            {
                $item.PathExists = $false
            }
            else
            {
                $item.PathExists = $true
            }

            $IsCalcPath = ($DocPathCalc -eq $item.DocPath)
            $IsDepCode = ($DepCodeCalc -eq $item.Dep_Code)
            $IsDocClass = ($DocClassCalc -eq $item.Doc_Class[1])
            $IsDocType = ($DocTypeCalc -eq $item.Doc_Type)

            $item.DocPathCalc = $DocPathCalc
            $item.IsCalcPath = $IsCalcPath
            $item.IsDepCode = $IsDepCode
            $item.IsDocClass = $IsDocClass
            $item.IsDocType = $IsDocType
        }
    }

    $listItems | Export-Csv -Path $filePath -Delimiter ';' -NoTypeInformation

    Write-Host "[SUCCESS] Log generato nel percorso $($filePath)" -ForegroundColor Green
    Write-Progress -Activity 'Controllo valori' -Completed

}
catch
{
    Write-Progress -Activity 'Controllo valori' -Completed
    Throw
}