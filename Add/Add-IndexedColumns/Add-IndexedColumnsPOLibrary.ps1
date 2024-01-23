<#
    Il CSV deve avere la colonna SiteUrl
#>

#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2.0" }

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [String]$Code = $siteCode
    )

    $ExecutionDate = Get-Date -Format 'yyyy_MM_dd'
    $logPath = "$($PSScriptRoot)\logs\$($Code)-$($ExecutionDate).csv";

    if (!(Test-Path -Path $logPath)) {
        $newLog = New-Item $logPath -Force -ItemType File
        Add-Content $newLog 'Timestamp;Type;ListName;ID/Doc;Action;Key;Value;OldValue'
    }
    $FormattedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Message.Contains('[SUCCESS]')) { Write-Host $Message -ForegroundColor Green }
    elseif ($Message.Contains('[ERROR]')) { Write-Host $Message -ForegroundColor Red }
    elseif ($Message.Contains('[WARNING]')) { Write-Host $Message -ForegroundColor Yellow }
    else {
        Write-Host $Message -ForegroundColor Cyan
        return
    }
    $Message = $Message.Replace(' - List: ', ';').Replace(' - Library: ', ';').Replace(' - Column: ', ';').Replace(' - ', ';')
    Add-Content $logPath "$FormattedDate;$Message"
}

$VendorList = 'Vendors'

Try {
    # Caricamento ID / CSV / Documento
    $CSVPath = (Read-Host -Prompt 'SiteUrl or CSV Path').Trim('"')
    If ($CSVPath.ToLower().Contains('.csv')) { $csv = Import-Csv -Path $CSVPath -Delimiter ';' }
    Else {
        $csv = [PSCustomObject] @{
            SiteUrl = $CSVPath
            Count   = 1
        }
    }

    $columnsToIndex = (Read-Host -Prompt 'Columns (,)').Split(',')

    $rowCounter = 0
    Write-Log 'Inizio operazione...'
    ForEach ($row in $csv) {
        if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }
        $mainConn = Connect-PnPOnline -Url $row.SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
        $siteCode = $row.SiteUrl.Split('/')[-1]

        Write-Host "Site: $($siteCode)" -ForegroundColor Blue

        Write-Log "Caricamento '$($VendorList)'..."
        $Vendors = Get-PnPListItem -List $VendorList -PageSize 5000 -Connection $mainConn | ForEach-Object {
            [PSCustomObject]@{
                ID      = $_['ID']
                Name    = $_['Title']
                SiteUrl = $_['VD_SiteUrl']
            }
        }
        Write-Log 'Caricamento lista completato.'

        $vendorCounter = 0
        ForEach ($vendor in $vendors) {
            $vendorCounter++
            Write-Log "$($vendorCounter)/$($Vendors.Length) - Vendor: $($vendor.Name)"
            $subConn = Connect-PnPOnline -Url $vendor.SiteUrl -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop -ReturnConnection
            $POLists = Get-PnPList -Connection $subConn | Where-Object -FilterScript { $_.Title -match '^\d{10}$' }

            ForEach ($library in $POLists.Title) {
                $fields = Get-PnPField -List $library -Connection $subConn
                $indexedColumns = $fields | Where-Object { $_.Indexed -eq $true }

                # Check indexed column limit
                if ($indexedColumns.Count -lt 20) {
                    # Process each column
                    ForEach ($column in $columnsToIndex) {
                        If ($indexedColumns.InternalName -notcontains $column) {
                            Try {
                                Set-PnPField -List $library -Identity $column -Values @{ Indexed = $true } -Connection $subConn
                                Write-Log "[SUCCESS] - Library: $($library) - Column: $($column) - INDEXED"
                            }
                            Catch { Write-Log "[ERROR] - Library: $($library) - Column: $($column) - $($_)" }
                        }
                        Else { Write-Log "[WARNING] - Library: $($library) - Column: $($column) - ALREADY INDEXED" }
                    }
                }
                Else { Write-Log "[ERROR] - Library: $($library) - Column: $($column) - TOO MANY INDEXED" }
            }
        }
    }
    if ($csv.Count -gt 1) { Write-Progress -Activity 'Aggiornamento' -Status "$($rowCounter+1)/$($csv.Count)" -PercentComplete (($rowCounter++ / $csv.Count) * 100) }
    Write-Log 'Operazione completata.'
}
catch { Throw }