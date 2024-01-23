$FormulaMappingCSV = "$pwd\CondtionalFormulaCsv.csv"
$SiteUrl = 'https://tecnimont.sharepoint.com/sites/TIMS_B4_Telco'

###################

Function Set-ConditionalFormula($ListDisplayName, $SiteColumnInternalName, $Formula) {

	$Field = Get-PnPField -List $ListDisplayName -Identity $SiteColumnInternalName
	$Field.ClientValidationFormula = $Formula
	$Field.Update()
	Invoke-PnPQuery

}

$ImportedCsv = Import-Csv -Path $FormulaMappingCSV -Delimiter ';'
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

ForEach ($Row in $ImportedCsv) {

	Try {
		Set-ConditionalFormula $($Row.ListDisplayName) $($Row.ColumnInternalName) $($Row.Formula)
		Write-Host "Column `"$($Row.ColumnInternalName)`" updated on `"$($Row.ListDisplayName)`"" -ForegroundColor Green
	}
	Catch {
		Write-Host "FAILED UPDATING column `"$($Row.ColumnInternalName)`" on `"$($Row.ListDisplayName)`"" -ForegroundColor Red
	}
}