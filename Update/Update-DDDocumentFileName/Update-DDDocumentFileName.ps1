<# !QUESTO SCRIPT È UNA BOZZA SENZA PARAMETRI E CON VALORI HARD CODED

! add csv log export
! add progress bar
#>

$CSV = Import-Csv -Path 'C:\Users\ST-442\Downloads\INC0798842_4191-Document File Name Change List.csv' -Delimiter ';'

Connect-PnPOnline -UseWebLogin -Url 'https://tecnimont.sharepoint.com/sites/4191DigitalDocuments'

$DocumentList = Get-PnPListItem -List 'DocumentList' -PageSize 5000 | ForEach-Object {
    $item = New-Object -TypeName PSCustomObject -Property @{
        ID      = $($_['ID'])
        TCM_DN  = $($_['Title'])
        Rev     = $($_['IssueIndex'])
        DocPath = $($_['DocumentsPath'])
    }
    $item
}

Try {
    ForEach ($Doc in $CSV) {
        [array]$DocToRename = $DocumentList | Where-Object { $_.TCM_DN -eq $Doc.TCM_DN -and $_.Rev -eq $Doc.Rev }
        If ($DocToRename.Count -eq 1) {
            Write-Host ('{0} - {1}' -f $Doc.TCM_DN, $Doc.Rev) -ForegroundColor Cyan

            $DocPathServerRelativeUrl = $DocToRename.DocPath -replace 'https://tecnimont.sharepoint.com/sites/4191DigitalDocuments', ''
            $FolderItems = Get-PnPFolderItem -FolderSiteRelativeUrl $DocPathServerRelativeUrl -ItemType File -Recursive | Where-Object { $_.Name -like "*$($Doc.TCM_DN)*" }
            If ($FolderItems.Count -eq 0) {
                Write-Host ("{0} - {1} - No files found in `n{2}" -f $Doc.TCM_DN, $Doc.Rev, $DocPathServerRelativeUrl) -ForegroundColor Red
                Write-Host ''
            }
            else {
                ForEach ($File in $FolderItems) {
                    $FileURL = $File.ServerRelativeUrl -replace '/sites/4191DigitalDocuments', ''
                    $NewFileName = $File.Name -replace $Doc.TCM_DN, $Doc.Doc_FileName
                    Write-Host ('{0}' -f $FileURL) -ForegroundColor Gray
                    Write-Host ("{0} `nrenamed in `n{1}" -f $File.Name, $NewFileName) -ForegroundColor Green
                    Rename-PnPFile -SiteRelativeUrl $FileURL -TargetFileName $NewFileName -Force
                    Write-Host ''
                }
            }

        }
        else {
            Write-Host ('{0} occurrences found of {1} - {2}' -f $DocToRename.Count, $Doc.TCM_DN, $Doc.Rev)
        }
        Write-Host ''
    }
}
catch {
    Write-Host ($_ | Out-String) -ForegroundColor Red
}