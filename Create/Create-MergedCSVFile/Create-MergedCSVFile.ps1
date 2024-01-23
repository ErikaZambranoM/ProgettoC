# Comment-Based Help
<#
.SYNOPSIS
    Merges all CSV files in a specified folder.
.DESCRIPTION
    The script scans the folder for CSV files and merges them into a single CSV file.
.PARAMETER FolderPath
    The path to the folder containing the CSV files to merge.
.PARAMETER OutputPath
    The path where the merged CSV file will be saved.
#>

# Function to Merge CSV files
function Create-MergedCSVFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'Enter the folder path containing CSV files to merge.')]
        [string]$FolderPath,

        [Parameter(Mandatory = $true, HelpMessage = 'Enter the path where the merged CSV file should be saved.')]
        [string]$OutputPath
    )
    Begin {
        # Validate Folder Path
        if (-Not (Test-Path $FolderPath)) {
            Throw 'The specified folder path does not exist.'
        }
    }
    Process {
        # Initialize merged data variable
        $mergedData = @()

        # Get CSV files from folder
        $csvFiles = Get-ChildItem -Path $FolderPath -Filter *.csv

        # Validate if files found
        if ($csvFiles.Count -eq 0) {
            Throw 'No CSV files found in the specified folder.'
        }

        # Initialize Progress Counter
        $currentIndex = 0
        $totalItems = $csvFiles.Count

        # Loop through each CSV file
        foreach ($csv in $csvFiles) {
            # Progress bar
            $currentIndex++
            Write-Progress -PercentComplete (($currentIndex / $totalItems) * 100) -Status 'Processing' -Activity 'Merging CSV files into one...' -CurrentOperation "$currentIndex out of $totalItems"

            # Read CSV data
            $data = Import-Csv -Path $csv.FullName

            # Add data to merged data
            $mergedData += $data
        }
    }
    End {
        # Export merged data to new CSV file
        $mergedData | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Progress -Activity 'Merging CSV files into one...' -Completed
        Write-Host "CSV files have been merged successfully. The merged file is saved at $OutputPath."
    }
}

# Example usage
Create-MergedCSVFile -FolderPath 'C:\Users\ST-442\Downloads\CSVMergeTest\C' -OutputPath 'C:\Users\ST-442\Downloads\CSVMergeTest\Merged Client Only\MergedFile.csv'
