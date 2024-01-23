<# Module manifest file (AMS-Utilities.psd1)
    AMS-Utilities.psm1 gets imported for functions, the AMS-ClassesImporter.ps1 gets called to import classes from AMS-Classes.psm1
#>

@{
    # Module information
    ModuleVersion        = '1.0.0'
    GUID                 = '7850fdb7-9875-4568-a9e1-cc007d7032af'
    Author               = 'Federico Barone, Davide De Luca'
    Description          = 'Module created on top of PnP.PowerShell module to help AMS Doc Support team in their daily tasks.'

    # Root module
    RootModule           = 'AMS-Utilities.psm1'
    ScriptsToProcess     = @('.\Classes\ClassesImporter.ps1')

    # Module dependencies
    PowerShellVersion    = '7.3.4'
    CompatiblePSEditions = @('Core')
    RequiredModules      = @(
        @{
            ModuleName    = 'PnP.PowerShell'
            ModuleVersion = '2.2.0'
        }
    )

    # Exported functions
    FunctionsToExport    = '*'
}
