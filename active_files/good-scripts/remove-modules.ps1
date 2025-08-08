# Clean slate
Uninstall-Module -Name ExchangeOnlineManagement -AllVersions -Force
#Uninstall-Module -Name Az.Accounts -AllVersions -Force

# Install compatible versions
Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
#Install-Module -Name Az.Storage -RequiredVersion 5.7.0 -Force

# Connect
# Connect-AzAccount
