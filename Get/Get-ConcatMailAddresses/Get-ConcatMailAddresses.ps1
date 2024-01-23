$MailAddresses = @'
Canducci Alessandro <A.Canducci@tecnimont.it>
Lattuada Francesco <F.Lattuada@tecnimont.it>
Del Gesso Angela Piera <A.DelGesso@tecnimont.it>
Ulatowska Anna Barbara <A.Ulatowska@tecnimont.it>
Gentile Francesco (PROEN) <F.Gentile2@tecnimont.it>
Zaoralkova Lenka <L.Zaoralkova@tecnimont.it>

'@

Function IsValidEmail($MailAddress)
{
    $MailAddress -Match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
}

# Remove HTML tags

# Regex pattern for matching email addresses
$EmailPattern = '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'

# Extract valid email addresses
$ParsedMailAddressesArray = [regex]::Matches($MailAddresses, $EmailPattern) | ForEach-Object { $_.Value }


#! Remove HTML tags (Old way)
<#
    $MailAddresses = $MailAddresses -replace '[^\w\.\-+@]', ''
    $MailAddresses = $MailAddresses -replace '<|>', ''
    $MailAddresses = $MailAddresses -replace '<[^>]+>', ' '
    $MailAddresses = $MailAddresses -replace 'mailto:', ' '
    $MailAddresses = $MailAddresses -replace '>', ' '

    # Remove all separators
    $MailAddresses = $MailAddresses -replace '"', ''
    $MailAddresses = $MailAddresses -replace "`r`n", ' '
    $MailAddresses = $MailAddresses -replace ';', ' '
    $MailAddresses = $MailAddresses -replace ',', ' '

    # Create array of MailAddresses
    $MailAddressesArray = $MailAddresses.Split(' ')
    $ValidMailAddressesArray = @()
    $ParsedMailAddressesArray = @()
    $ConcatMailAddress = @()
    $ConcatMail = ''

    ForEach ($Mail in $MailAddressesArray)
    {
        If (!($Mail.Contains('@')) -or '' -ne $ConcatMail)
        {
            $ConcatMail += $Mail
            If ($ConcatMail.Contains('@'))
            {
                $ParsedMailAddressesArray += $ConcatMail
            $ConcatMailAddress += $ConcatMail
            $ConcatMail = ''
        }
    }
    Else
    {
        $ParsedMailAddressesArray += $Mail
    }
}
#>

Write-Host "$($ParsedMailAddressesArray.Count) MailAddresses found!" -ForegroundColor Cyan
Write-Host ''

$UniqueMailAddresses = $ParsedMailAddressesArray | Select-Object -Unique
$InvalidMailAddress = $ParsedMailAddressesArray | Where-Object -FilterScript { !(IsValidEmail $_) }
$DuplicatedMailAddresses = (Compare-Object -ReferenceObject $ParsedMailAddressesArray -DifferenceObject $UniqueMailAddresses).InputObject | Where-Object -FilterScript { $_ -notin $InvalidMailAddress }
$ValidMailAddressesArray = $UniqueMailAddresses | Where-Object -FilterScript { $_ -notin $InvalidMailAddress }


$UniqueMailAddressesString = ($ValidMailAddressesArray -join ';')
Set-Clipboard -Value $UniqueMailAddressesString

<# Only for old version
If ($ConcatMailAddress.Count -gt 0)
{
    Write-Host "$($ConcatMailAddress.Count) of $($ParsedMailAddressesArray.Count) parsed MailAddresses contained empty spaces and have been automatically concatenated:" -ForegroundColor Yellow
    Write-Host ''
    Write-Host ($ConcatMailAddress | Out-String).Trim() -ForegroundColor Yellow
    Write-Host ''
}
#>

Write-Host "$($ValidMailAddressesArray.Count) unique and valid MailAddresses found and copied to clipboard:" -ForegroundColor Green
Write-Host ''
Write-Host $UniqueMailAddressesString -ForegroundColor Green
Write-Host ''

If ($DuplicatedMailAddresses.Count -gt 0)
{
    Write-Host "$($DuplicatedMailAddresses.Count) duplicated but valid MailAddresses found and have not been copied to clipboard:" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host ($DuplicatedMailAddresses | Group-Object | Select-Object -Property @{L = 'MailAddress'; E = { $_.Name } }, @{L = 'Occurrences'; E = { $_.Count + 1 } } | Out-String).Trim() -ForegroundColor DarkGray
    Write-Host ''
}

If ($InvalidMailAddress.Count -gt 0)
{
    Write-Host "$($InvalidMailAddress.Count) invalid MailAddresses found:" -ForegroundColor Red
    Write-Host ''
    Write-Host ($InvalidMailAddress | Out-String).Trim() -ForegroundColor Red
}