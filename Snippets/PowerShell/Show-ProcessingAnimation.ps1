Function Show-ProcessingAnimation {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, HelpMessage = 'The script block to execute.')]
        [ScriptBlock]
        $ScriptBlock,

        [HashTable]
        $ScriptBlockParams,

        [Parameter(HelpMessage = 'Where to display the animation.')]
        [ValidateSet('Console', 'WindowTitle', 'Both')]
        [string]
        $AnimationPosition = 'Console'
    )

    # Store the original console title to restore it later
    $originalTitle = $Host.UI.RawUI.WindowTitle

    Try {
        [Console]::CursorVisible = $false

        $Counter = 0
        $Frames = '|', '/', '-', '\'
        $Job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ScriptBlockParams

        While ($Job.State -eq 'Running') {
            $Frame = $Frames[$Counter % $Frames.Length]

            if ($AnimationPosition -eq 'Console' -or $AnimationPosition -eq 'Both') {
                Try {
                    # Dynamically update cursor position based on current console size
                    $CursorTop = [Math]::Min([Console]::CursorTop, [Console]::BufferHeight - 1)

                    [Console]::SetCursorPosition(0, $CursorTop)
                    Write-Host "$Frame" -NoNewline
                }
                Catch {
                    # Even if we fail to set the cursor position, continue the animation
                    Write-Host "$Frame" -NoNewline
                }
            }

            if ($AnimationPosition -eq 'WindowTitle' -or $AnimationPosition -eq 'Both') {
                # Append the frame to the current window title
                $Host.UI.RawUI.WindowTitle = "$originalTitle - Processing  $Frame"
            }

            $Counter += 1
            Start-Sleep -Milliseconds 125
        }

        # Receive job results or errors
        $Result = Receive-Job -Job $Job
        Remove-Job -Job $Job

        # Only needed if you use multiline frames in console mode
        if ($AnimationPosition -eq 'Console' -or $AnimationPosition -eq 'Both') {
            Write-Host "`r" + ($Frames[0] -replace '[^\s+]', ' ') -NoNewline
        }
    }
    Catch {
        Write-Host "An error occurred: $_"
    }
    Finally {
        Try {
            if ($AnimationPosition -eq 'Console' -or $AnimationPosition -eq 'Both') {
                [Console]::SetCursorPosition(0, [Console]::CursorTop)
            }

            if ($AnimationPosition -eq 'WindowTitle' -or $AnimationPosition -eq 'Both') {
                # Restore the original console title
                $Host.UI.RawUI.WindowTitle = $originalTitle
            }
        }
        Catch {
            Write-Host "An error occurred while resetting: $_"
        }

        [Console]::CursorVisible = $true
    }

    Return $Result
}

# Sample usage with script block parameters
$ScriptBlockParams = @{
    Param1_Name = 'Param1_Value'
    Param2_Name = 'Param2_Value'
}

$Scriptblock = Show-ProcessingAnimation -ScriptBlock {
    Param ($Params)
    # Use $Params['ParamName'] to reference the parameters
    $Counter = 0
    $TestOutput1 = @()
    $TestOutput2 = @()
    $TestOutput3 = @()
    Do {
        $Counter++
        $TestOutput1 += "TestOutput: $Counter"
        $TestOutput2 += "TestOutput: $($Params['Param1_Name'])"
        $TestOutput3 += "TestOutput: $($Params['Param2_Name'])"

    } While ($Counter -lt 50)

    # Name and value of the variable to return as property of the returned object
    Return [Ordered]@{TestOutput1 = $TestOutput1; TestOutput2 = $TestOutput2; TestOutput3 = $TestOutput3 }

} -ScriptBlockParams $ScriptBlockParams
$Scriptblock