# Function that converts a CSV file to an Excel file and adds hyperlinks to the cells containing URLs.
Function Convert-CSVToExcel {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
        [String]
        $CSVPath
    )

    Try {
        # Determine the Excel path based on the CSV path
        $ExcelPath = [System.IO.Path]::ChangeExtension($CSVPath, 'xlsx')
        $CSVFileName = [System.IO.Path]::GetFileName($CSVPath)
        $ExcelFileName = [System.IO.Path]::GetFileName($ExcelPath)

        # Create Excel COM object
        $ExcelApp = New-Object -ComObject Excel.Application

        # Make Excel invisible
        $ExcelApp.Visible = $false

        # Disable alerts to suppress overwrite confirmation
        $ExcelApp.DisplayAlerts = $false

        # Open CSV
        $Workbook = $ExcelApp.Workbooks.Open($CSVPath)
        $Worksheet = $Workbook.Worksheets.Item(1)

        # Get the Range of data and create a table
        $Range = $Worksheet.UsedRange
        $Worksheet.ListObjects.Add(1, $Range, $null, 1) | Out-Null

        # Get the Columns with names containing 'Link'
        $LinkColumns = @()
        For ($i = 1; $i -le $Worksheet.UsedRange.Columns.Count; $i++) {
            If ($Worksheet.Cells.Item(1, $i).Text -like '*Link*') {
                $LinkColumns += $i
            }
        }

        # Find the last Row with data
        $LastRow = $Worksheet.UsedRange.Rows.Count

        # Iterate through link Columns and turn Cell Values into hyperlinks
        ForEach ($Column in $LinkColumns) {
            $ColumnName = $Worksheet.Cells.Item(1, $Column).Text # Getting the column name
            For ($Row = 2; $Row -le $LastRow; $Row++) {
                # Starting from Row 2 to skip the header
                $Cell = $Worksheet.Cells.Item($Row, $Column)
                $Value = $Cell.Text
                If (
                    $Value -match 'https?://\S+' -or
                    $Value -match 'http?://\S+' -or
                    $Value -match 'www\.\S+' -or
                    $Value -match 'mailto:\S+'
                ) {
                    If ($Value.Length -le 2083) {
                        $HyperLink = $Worksheet.Hyperlinks.Add($Cell, $Value)
                        Switch ($ColumnName) {
                            'Link to Tickets' {
                                # Customizing hyperlink text based on pattern
                                $LinkText = ((($Value -split 'numberIN')[1] -split '%255EORDERBY' | Select-Object -First 1) -split '%252C') -join ', '
                                $Hyperlink.TextToDisplay = $LinkText
                                Break
                            }
                            Default
                            {}
                        }
                    }
                }
            }
        }

        # Autofit the Columns and Wrap the Text
        $Range.WrapText = $false
        $Range.EntireColumn.AutoFit() | Out-Null

        # Attempt to save the file with retries
        $Success = $false
        While (-not $Success) {
            Try {
                # Save the Workbook
                $Workbook.SaveAs($ExcelPath, 51) # 51 = xlsx format
                $Success = $true
            }
            Catch {
                Write-Host ''
                Write-Warning ('Failed to save the Excel file. It may be in use by another process.')

                $ChoiceTitle = 'Excel file in use during CSV to XLSX conversion'
                $ChoiceMessage = ("File '{0}' is in use by another process.{1}Do you want to manually close the file and try again?" -f $ExcelFileName, "`n")
                $RetryChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Retry', ('{0}Retry saving after manually closing the file:{0}{1}{0}{0}' -f "`n", $ExcelPath)
                $TerminateChoice = New-Object System.Management.Automation.Host.ChoiceDescription '&Abort', ('{0}Abort the operation. It can be run later using this command:{0}{1}{0}' -f "`n", $MyInvocation.Line )
                $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($RetryChoice, $TerminateChoice)

                $Result = $Host.UI.PromptForChoice($ChoiceTitle, $ChoiceMessage, $Choices, 0)

                Switch ($Result) {
                    # Retry
                    0 {
                        Write-Host ''
                    }

                    # Abort
                    1 {
                        Write-Host ''
                        Write-Warning ('You choosed to abort the save process. To try again later, you can run this command:{0}{1}' -f "`n", $MyInvocation.Line)
                        Return $false
                    }
                }
            }

        }

        # Close Excel
        $ExcelApp.Quit()

        Write-Host ("{0}File '{1}' converted in '{2}'" -f "`n", $CSVFileName, $ExcelFileName) -ForegroundColor Green
        Return $true
    }
    Catch {
        Throw
    }
    Finally {
        # Clean up by releasing all COM objects created
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null

        # Force garbage collection to clean up any lingering objects
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
