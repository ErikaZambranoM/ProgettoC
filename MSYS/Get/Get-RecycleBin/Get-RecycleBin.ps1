#Requires -Version 7 -Modules @{ ModuleName = "PnP.PowerShell"; ModuleVersion = "2.2" }

<#
.SYNOPSIS
    Export a SharePoint Online site's Recycle Bin to a CSV file.

.DESCRIPTION
    This script exports the Recycle Bin of a SharePoint Online site or subsite to a CSV file.

.PARAMETER SiteURL
    Mandatory parameter. Specifies the URL of the SharePoint Online site or subsite.

.PARAMETER SecondStage
        Optional switch. Specifies whether to export the second stage Recycle Bin instead of the first stage Recycle Bin.
        By default, the first stage Recycle Bin is exported.

.EXAMPLE
    PS C:\> .\Get-RecycleBin.ps1 -SiteURL "https://contoso.sharepoint.com/sites/contoso" -SecondStage
    This example exports the second stage Recycle Bin of the "https://contoso.sharepoint.com/sites/contoso" site.

.OUTPUTS
    The script exports the Recycle Bin to a CSV file in the "Logs" folder of the script.
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            # Match a SharePoint Main Site or Sub Site URL
            If ($_ -match '^https://[a-zA-Z0-9-]+\.sharepoint\.com/Sites/[\w-]+(/[\w-]+)?/?$') {
                $True
            }
            Else {
                Throw "`n'$($_)' is not a valid SharePoint Online site or subsite URL."
            }
        })]
    [String]
    $SiteURL,

    [Switch]
    $SecondStage
)

Begin {
    Try {
        # Connect to SharePoint Online
        $SiteURL = $SiteURL.TrimEnd('/')
        Connect-PnPOnline -Url $SiteURL -UseWebLogin -ValidateConnection -WarningAction SilentlyContinue -ErrorAction Stop

        # Get site code and compose log file path
        $SiteCode = (Get-PnPWeb).Title.Split(' ')[0]
        $LogRootPath = "$($PSScriptRoot)\Logs"
        $CSVExportPath = "$($LogRootPath)\$($SiteCode)-RecycleBin-$((Get-Date).ToString('dd-MM-yyyy_HH-mm-ss')).csv"
    }
    Catch {
        $ScriptResult = 'ERROR'
        Throw
    }
}

Process {
    Try {
        # Download Recycle Bin
        Write-Host 'Downloading Recycle Bin...' -ForegroundColor Cyan
        $RecycleBin = $SecondStage ? (Get-PnPRecycleBinItem -SecondStage) : (Get-PnPRecycleBinItem -FirstStage)

        # Create log folder if it doesn't exist
        If (!(Test-Path -Path $LogRootPath -PathType Container)) {
            New-Item -Path $LogRootPath -ItemType Directory | Out-Null
        }

        # Export Recycle Bin
        $RecycleBin | Select-Object -Property Title, ItemType, Size, ItemState, DirName, DeletedByName, DeletedDate, Id |
            Export-Csv -Path $CSVExportPath -NoTypeInformation -Delimiter ';' | Out-Null

        # Manually add the BOM
        $CSV_Content = Get-Content -Path $CSVExportPath -Raw
        $CSV_Content = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetPreamble()) + $CSV_Content
        Set-Content -Path $CSVExportPath -Value $CSV_Content -Encoding UTF8

        $ScriptResult = 'SUCCESS'
    }
    Catch {
        $ScriptResult = 'ERROR'
        Throw
    }
}

End {
    If ($ScriptResult -eq 'SUCCESS') {
        Write-Host "[$ScriptResult] Recycle Bin exported in:`n$($CSVExportPath)" -ForegroundColor Green
    }
    Else {
        Write-Host "[$ScriptResult] Recycle Bin export failed." -ForegroundColor Red
    }
}