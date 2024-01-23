# Function that returns a Hyperlink clickable in the PowerShell console
Function New-PowerShellHyperlink
{
    Param(
        [Parameter(Mandatory = $true)]
        [String]
        $LinkURL,

        [Parameter(Mandatory = $true)]
        [String]
        $LinkDisplayText
    )

    Try
    {
        if ($LinkURL -match '^https?://')
        {
            $PowerShellHyperlink = ("`e]8;;{0}`e\{1}`e]8;;`e\" -f $LinkURL, $LinkDisplayText)
        }
        else
        {
            $PowerShellHyperlink = ("`e]8;;file:///{0}`e\{1}`e]8;;`e\" -f $LinkURL, $LinkDisplayText)
        }
        Return $PowerShellHyperlink
    }
    Catch
    {
        Throw
    }
}