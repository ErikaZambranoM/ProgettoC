#DA TESTARE

#Chiedo all'utente di inserire il codice di progetto e la tipologia di sito
Param (
    [parameter(Mandatory = $true)]
    [String]$ProjectCode, #Inserire codice di progetto
    [parameter(Mandatory = $true)]
    [validateset('VD', 'DD', 'DDClient')]
    [String]$SiteType #Inserire tipologia sito (VD/DD/DDClient)
)

#In base al SiteType selezionato ricostruisco il SiteUrl e altri parametri che mi serviranno più avanti per la modifica di valori sulla lista MD Settings
try {
    $SiteType = $SiteType.ToUpper() # Convert SiteType to uppercase

    switch ($SiteType) {
        'VD' {
            $SiteUrl = "https://tecnimont.sharepoint.com/sites/vdm_$($ProjectCode)"
            $DocumentListName = 'Vendor Documents List'
            $DocumentListServerRelativeUrl = "/sites/vdm_$($ProjectCode)/Lists/VendorDocumentsList"
        }
        'DD' {
            $SiteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocuments"
            $DocumentListName = 'DocumentList'
            $DocumentListServerRelativeUrl = "/sites/$($ProjectCode)DigitalDocuments/Lists/DocumentList"
        }
        'DDCLIENT' {
            $SiteUrl = "https://tecnimont.sharepoint.com/sites/$($ProjectCode)DigitalDocumentsC"
            $DocumentListName = 'Client Document List'
            $DocumentListServerRelativeUrl = "/sites/$($ProjectCode)DigitalDocumentsC/Lists/ClientDocumentList"
            $SiteType = 'DDClient'
        }
        Default {
            Write-Host -ForegroundColor Red 'SiteType not recognized, must be one of VD/DD/DDClient' -ErrorAction Stop
        }
    }
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ValidateConnection -ErrorAction Stop
    Write-Host -ForegroundColor Green "Connessione a $($SiteUrl) effettuata"

    <#
###################################### RIMOZIONE PERMESSI NON NECESSARI ######################################
        ####### NOT WORKING #######
$ListTitles = "MD Basket Areas", "MD Download Sessions","MDMassiveDownloadTemp" #liste su cui voglio operare
$MDSettings_List = Get-PnPList | Where-Object {$ListTitles -contains $_.Title}
$excludedGroupNames = "MassiveDownload","$($ProjectCode) - Vendor Documents Owners" #gruppi che voglio mantenere, aggiungere scelta automatica e non manuale DD/VD

foreach($item in $MDSettings_List)
{
    $groups = Get-PnPGroup -List $item
    foreach ($group in $groups) {
        $groupName = $group.Title
        if ($excludedGroupNames -notcontains $groupName) {
            Remove-PnPGroup -Identity $group
            Write-Host "Deleted permission group: $groupName"}
        }
        # aggiungere elseif per cambiare permessi to MT-Contributors di MassiveDownload
    }

#>
    ###################################### AGGIORNAMENTO PARAMETRI LISTA MD Settings ###################################### TESTED & WORKING

    #Dichiaro il nome della lista da modificare
    $ListToModifyName = 'MD Settings'

    #Salvo e parametrizzo la lista dichiarata
    $ToModify_List = Get-PnPListItem -List $ListToModifyName | ForEach-Object { [PSCustomObject] @{
            ID    = $_['ID']
            Key   = $_['DDMDKey']
            Value = $_['DDMDValue']
        }
    }
    Write-Host -ForegroundColor Cyan "Lista $($ListToModifyName) caricata"

    #Elenco le chiavi che voglio modificare nella lista dichiarata
    $KeysToFilter = @(
        'SiteType',
        'RestrictedGroupName',
        'DocumentListName',
        'DocumentListServerRelativeUrl',
        'DownloadArea'
    )

    $SettingsToUpdate = $ToModify_List | Where-Object -FilterScript { $_.Key -in $KeysToFilter }

    #vengono popolati i campi in $KeysToFilter in base al progetto e tipo di sito specificato con l'inserimento dei parametri all'inizio
    foreach ($row in $SettingsToUpdate) {
        switch ($row.Key) {
            'SiteType' {
                $values = @{
                    DDMDValue = $SiteType
                }
                Set-PnPListItem -List $ListToModifyName -Identity $row.ID -Values $values | Out-Null
                Write-Host -ForegroundColor Green '[ SUCCESS ] - SiteType updated'
                Break
            }
            'RestrictedGroupName' {
                $values = @{
                    DDMDValue = 'MassiveDownload'
                }
                Set-PnPListItem -List $ListToModifyName -Identity $row.ID -Values $values | Out-Null
                Write-Host -ForegroundColor Green '[ SUCCESS ] - RestrictedGroupName updated'
                Break
            }
            'DocumentListName' {
                $values = @{
                    DDMDValue = $DocumentListName
                }
                Set-PnPListItem -List $ListToModifyName -Identity $row.ID -Values $values | Out-Null
                Write-Host -ForegroundColor Green '[ SUCCESS ] - DocumentListName updated'
                Break
            }
            'DocumentListServerRelativeUrl' {
                $values = @{
                    DDMDValue = $DocumentListServerRelativeUrl
                }
                Set-PnPListItem -List $ListToModifyName -Identity $row.ID -Values $values | Out-Null
                Write-Host -ForegroundColor Green '[ SUCCESS ] - DocumentListServerRelativeUrl updated'
                Break
            }
            'DownloadArea' {
                $values = @{
                    DDMDValue = 'Temporary Documents Downloads'
                }
                Set-PnPListItem -List $ListToModifyName -Identity $row.ID -Values $values | Out-Null
                Write-Host -ForegroundColor Green '[ SUCCESS ] - DownloadArea updated'
                Break
            }
            Default {
                Write-Host -ForegroundColor Yellow ('Key {0} not found' -f $row.Key)
            }
        }
    }

}
catch {
    throw
}
