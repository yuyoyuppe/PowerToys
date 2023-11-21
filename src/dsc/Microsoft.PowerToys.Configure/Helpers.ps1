# Make sure that PowerShell 7 is installed
# Make sure that Install-Module -Name PSDesiredStateConfiguration -RequiredVersion 2.0.6 invoked

$env:PSModulePath += ";$pwd"
# Get-DSCResource | grep PowerToysConfigure # Make sure it's loaded
# ipmo .\Microsoft.PowerToys.Configure.psd1 # Force-import the module
# Invoke-DscResource -Name PowerToysConfigure -Method Set -ModuleName Microsoft.PowerToys.Configure -Property @{ BackupFile = '<PathToAConfigFile>' } # Invoke a method manually
# winget configure .\configuration.dsc.yaml