## https://learn.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide
## https://microsoft-365-extractor-suite.readthedocs.io/en/latest/index.html
## https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-schema#auditlogrecordtype

<## INSTALL MODULES
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.9.0
Install-Module -Name Microsoft.Graph -Force # -RequiredVersion 2.32.0
Install-Module -Name Microsoft.Graph.Beta -Force # -RequiredVersion 2.32.0
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
Install-Module -Name MicrosoftTeams -Force
Install-Module -Name Microsoft-Extractor-Suite -Force
Install-Module -Name ImportExcel -Force


Import-Module -Name ExchangeOnlineManagement # -RequiredVersion 3.9.0
Import-Module -Name Microsoft.Graph # -RequiredVersion 2.32.0
Import-Module -Name Microsoft.Graph.Beta # -RequiredVersion 2.32.0
Import-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
Import-Module -Name MicrosoftTeams -Force
Import-Module -Name Microsoft-Extractor-Suite -Force
Import-Module -Name ImportExcel -Force
##>

Connect-MgGraph -Scopes ("AuditLog.Read.All",
                        "AuditLogsQuery.Read.All",
                        "AuditLogsQuery-Exchange.Read.All",
                        "Application.Read.All",
                        "Device.Read.All",
                        "Directory.Read.All",
                        "Group.ReadWrite.All",
                        "Organization.Read.All",
                        "Policy.Read.All",
                        "Policy.Read.ConditionalAccess",
                        "User.Read.All",
                        "UserAuthenticationMethod.Read.All",
                        "SecurityEvents.Read.All")

Connect-ExchangeOnline
Connect-AzAccount



# function
$OutputDir = 'C:\scripts\mext'
function Reset-Folder {
    param (
        [string]$ParentPath = "C:\Scripts",
        [string]$FolderName = "Mext"
    )

    $fullPath = Join-Path $ParentPath $FolderName

    if (Test-Path $fullPath) {
        Write-Host "Folder '$fullPath' exists. Clearing contents..." -ForegroundColor Yellow
        try {
            Get-ChildItem -Path $fullPath -Recurse -Force | Remove-Item -Recurse -Force
            Write-Host "Contents of '$fullPath' deleted successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error deleting contents of '$fullPath': $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Folder '$fullPath' does not exist. Creating it..." -ForegroundColor Cyan
        try {
            New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
            Write-Host "Folder '$fullPath' created successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error creating folder '$fullPath': $_" -ForegroundColor Red
        }
    }
}
Reset-Folder

## ENTRA ##
$OutputDir = 'C:\scripts\Mext\Entra'
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}
Get-MFA -OutputDir $OutputDir 
Get-Users -OutputDir $OutputDir 
Get-AdminUsers -OutputDir $OutputDir 
Get-AllRoleActivity -OutputDir $OutputDir 
Get-ConditionalAccessPolicies -OutputDir $OutputDir
Get-OAuthPermissionsGraph -OutputDir $OutputDir
Get-SecurityAlerts -OutputDir $OutputDir -DaysBack 7

$OutputDir = 'C:\scripts\Mext\Groups'
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}
Get-Groups -OutputDir $OutputDir
Get-DynamicGroups -OutputDir $OutputDir
Get-GroupMembers -OutputDir $OutputDir
#Get-GroupOwners -OutputDir $OutputDir
#Get-GroupSettings -OutputDir $OutputDir

$OutputDir = 'C:\scripts\Mext\Licenses'
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}
Get-Licenses -OutputDir $OutputDir
Get-LicenseCompatibility -OutputDir $OutputDir
Get-EntraSecurityDefaults -OutputDir $OutputDir
Get-LicensesByUser -OutputDir $OutputDir
#Get-ProductLicenses -OutputDir $OutputDir # not working

$OutputDir = 'C:\scripts\Mext\Devices'
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}
Get-Devices -OutputDir $OutputDir


$OutputDir = 'C:\scripts\Mext\EOL'
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}
## EXCHANGE ONLINE ##
Get-MailboxRules -OutputDir $OutputDir 
Get-TransportRules -OutputDir $OutputDir
Get-MailboxAuditStatus -OutputDir $OutputDir
Get-MailboxPermissions -OutputDir $OutputDir
#Get-Sessions -StartDate 2025-08-05 -EndDate 2025-08-06 -OutputDir $OutputDir # takes long time and limit is 5k


## AUDIT ##
$startDate = (Get-Date).AddDays(-14)
$endDate = (Get-Date)
$OutputDir = 'C:\scripts\Mext\Audit'

if (!(Test-Path $OutputDir)) {New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null}

## Unified Audit Log All Activities ##
Get-UAL -OutputDir $OutputDir -StartDate $startDate -EndDate (Get-Date) -Output CSV -MergeOutput
#Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate $endDate -RecordType * | Export-Csv -Path "$OutputDir\ualtest.csv" -nti -Force 

## UAL AAD Activities ##
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType AzureActiveDirectory | Export-Csv -Path "$OutputDir\aad.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType AzureActiveDirectoryAccountLogon | Export-Csv -Path "$OutputDir\login.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType AzureActiveDirectoryStsLogon | Export-Csv -Path "$OutputDir\sts.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType AadRiskDetection | Export-Csv -Path "$OutputDir\Risky.csv" -nti -Force

## UAL Storae Activities ##
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType OneDrive | Export-Csv -Path "$OutputDir\onedrive.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType SharePoint | Export-Csv -Path "$OutputDir\spo.csv" -nti -Force 

## UAL Exchange Activities ##
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType ExchangeAdmin | Export-Csv -Path "$OutputDir\exoadm.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType ExchangeItem | Export-Csv -Path "$OutputDir\exoitem.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType ExchangeItemAggregated | Export-Csv -Path "$OutputDir\exoagg.csv" -nti -Force 
Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType ExchangeAggregatedOperation | Export-Csv -Path "$OutputDir\exoaggop.csv" -nti -Force 
Get-MessageTraceV2 -ResultSize 5000 -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) | Export-Csv -Path "$OutputDir\mtrace.csv" -nti -Force


Disconnect-AzAccount
Disconnect-ExchangeOnline
Disconnect-MgGraph
