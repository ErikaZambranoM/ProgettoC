$Global:ProgressBarsIds = @{}
$Global:ProgressBarsIds[0] = $true

# PreventiveAction ScriptBlock
Function Function-Activity1
{
    $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0] # 0
    $CurrentProgressBarId = $ParentProgressBarId + 1 # 1
    $Global:ProgressBarsIds[$CurrentProgressBarId] = $true # 1 = $true

    $Activity1_Count = 1..5
    $Activity1_Count | ForEach-Object {
        Write-Host ("Item of Activity 1: {0}" -f $_) -ForegroundColor Cyan
        $PercentComplete = ($_ / $Activity1_Count[-1] * 100)
        $ProgressBarSplatting = @{
            Activity        = ("Activity 1")
            Status          = ("Status {0}..." -f $_)
            PercentComplete = $PercentComplete
            ParentId        = $ParentProgressBarId
            Id              = $CurrentProgressBarId
        }
        Write-Progress @ProgressBarSplatting
        Function-Activity2
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Start-Sleep -Milliseconds 50
    }

}

# Function used inside PreventiveAction ScriptBlock
Function Function-Activity2
{
    $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0] # 1
    $CurrentProgressBarId = $ParentProgressBarId + 1 # 2
    $Global:ProgressBarsIds[$CurrentProgressBarId] = $true # 2 = $true

    $Activity2_Count = 1..5
    $Activity2_Count | ForEach-Object {
        Write-Host ("Item of Activity 2: {0}" -f $_) -ForegroundColor Cyan
        $PercentComplete = ($_ / $Activity2_Count[-1] * 100)
        $ProgressBarSplatting = @{
            Activity        = ("Activity 2")
            Status          = ("Status {0}..." -f $_)
            PercentComplete = $PercentComplete
            ParentId        = $ParentProgressBarId
            Id              = $CurrentProgressBarId
        }
        Write-Progress @ProgressBarSplatting
        Function-Activity3
        $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
        Start-Sleep -Milliseconds 50
    }
}

# Nested function used inside function inside PreventiveAction ScriptBlock

Function Function-Activity3
{
    $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0] # 2
    $CurrentProgressBarId = $ParentProgressBarId + 1 # 3
    $Global:ProgressBarsIds[$CurrentProgressBarId] = $true # 3 = $true

    $Activity3_Count = 1..5
    $Activity3_Count | ForEach-Object {
        Write-Host ("Item of Activity 3: {0}" -f $_) -ForegroundColor Cyan
        $PercentComplete = ($_ / $Activity3_Count[-1] * 100)
        $ProgressBarSplatting = @{
            Activity        = ("Activity 3")
            Status          = ("Status {0}..." -f $_)
            PercentComplete = $PercentComplete
            ParentId        = $ParentProgressBarId
            Id              = $CurrentProgressBarId
        }
        Write-Progress @ProgressBarSplatting
        Start-Sleep -Milliseconds 1000
    }
}

# Main (PATicket List)
$Count = 1..10
Foreach ($item in $Count)
{
    Write-Host "Item of Activity 0: $item" -ForegroundColor Green
    $PercentComplete = ($item / $Count[-1] * 100)
    $ProgressBarSplatting = @{
        Activity        = ("Activity 0")
        Status          = ("Status {0}..." -f $item)
        PercentComplete = [Math]::Round($PercentComplete)
        Id              = ([Array]$Global:ProgressBarsIds.Keys)[0] # 0
    }
    Write-Progress @ProgressBarSplatting

    # Run Remediation
    Function-Activity1
    $Global:ProgressBarsIds.Remove($Global:ProgressBarsIds.Count - 1)
}