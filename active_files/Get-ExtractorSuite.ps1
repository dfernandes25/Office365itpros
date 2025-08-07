

$OutputDir = 'C:\scripts\mext'


function Check-Modules {
    param (
        [Parameter(Mandatory)]
        [string[]]$ModuleNames
    )

    $logDir = "C:\scripts\logs"
    $logFile = Join-Path $logDir "ModuleInstallLog.txt"

    # Ensure log directory exists
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    # Function to write to log
    function Write-Log {
        param (
            [string]$Message,
            [string]$Level = "INFO"
        )
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp [$Level] $Message"
        Add-Content -Path $logFile -Value $logEntry
        Write-Host $logEntry
    }

    foreach ($module in $ModuleNames) {
        # Check if the module is already installed
        $installed = Get-Module -ListAvailable -Name $module

        if (-not $installed) {
            Write-Log "Module '$module' not found. Attempting to install..." "WARN"
            try {
                Install-Module -Name $module -Force -Scope CurrentUser -AllowClobber
                Write-Log "Module '$module' installed successfully." "SUCCESS"
            } catch {
                Write-Log "Failed to install module '$module': $_" "ERROR"
                continue
            }
        } else {
            Write-Log "Module '$module' is already installed." "INFO"
        }

        <# Import the module
        try {
            Import-Module -Name $module -Force
            Write-Log "Module '$module' imported successfully." "SUCCESS"
        } catch {
            Write-Log "Failed to import module '$module': $_" "ERROR"
        } #>
    }
}

Check-Modules -ModuleName @('ExchangeOnlineManagement',
                            'Az', 
                            'Microsoft.Graph',
                            'Microsoft.Graph.Beta',
                            'Microsoft-Extractor-Suite'
                            )


Connect-ExchangeOnline
Connect-MgGraph
Connect-AzureAZ


Connect-MgGraph -Scopes ("AuditLog.Read.All",
                        "Application.Read.All",
                        "Device.Read.All",
                        "Directory.Read.All",
                        "Group.Read.All",
                        "Organization.Read.All",
                        "Policy.Read.All",
                        "Policy.Read.ConditionalAccess",
                        "User.Read.All",
                        "UserAuthenticationMethod.Read.All",
                        "IdentityRiskyUser.Read.All",
                        "SecurityEvents.Read.All"
                        )




## ENTRA ##
Get-MFA -OutputDir $OutputDir 
Get-Users -OutputDir $OutputDir 
Get-AdminUsers -OutputDir $OutputDir 
Get-AllRoleActivity -OutputDir $OutputDir 
Get-GraphEntraSignInLogs -OutputDir $OutputDir -startDate 2025-08-01 #json output
Get-GraphEntraAuditLogs -OutputDir $OutputDir -startDate 2025-08-01 #json output
Get-ConditionalAccessPolicies -OutputDir $OutputDir
Get-OAuthPermissionsGraph -OutputDir $OutputDir
Get-SecurityAlerts -OutputDir $OutputDir -DaysBack 180

Get-Groups -OutputDir $OutputDir
Get-DynamicGroups -OutputDir $OutputDir
Get-GroupMembers -OutputDir $OutputDir

Get-Licenses -OutputDir $OutputDir
Get-LicenseCompatibility -OutputDir $OutputDir
Get-EntraSecurityDefaults -OutputDir $OutputDir
Get-LicensesByUser -OutputDir $OutputDir
# Get-ProductLicenses -OutputDir $OutputDir # not working

Get-Devices -OutputDir $OutputDir

<## need appropriate licensing
Get-RiskyUsers -OutputDir $OutputDir 
Get-RiskyDetections -OutputDir $OutputDir 
##>