param (
    [Parameter(Mandatory = $true)][String]$Folder
)

Get-ChildItem -Path $Folder -Directory | ForEach-Object {
    Get-ChildItem -Path $_.FullName -Directory | Where-Object -FilterScript { $_.Name.ToLower().Contains('attachment') } | Select-Object FullName
}
