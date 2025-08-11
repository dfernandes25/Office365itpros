<#
.SYNOPSIS
    <Short summary of what the script does>

.DESCRIPTION
    <Detailed description of the script's purpose, logic, and behavior>

.AUTHOR
    Don <your full name or initials if preferred>

.DATE CREATED
    2025-06-26

.LAST MODIFIED
    2025-06-26

.VERSION
    1.0.0

.REQUIRED MODULES
    - Microsoft.Graph
    - ExchangeOnlineManagement
    - ImportExcel

.DEPENDENCIES
    - PowerShell 5.1 or later
    - Administrator privileges (if applicable)
    - Valid credentials to Microsoft 365 with appropriate roles
    - Network access to Microsoft Graph and Exchange Online endpoints

.PARAMETERS
    <List and describe script parameters here if applicable>

.NOTES
    Run `Install-Module ModuleName -Scope CurrentUser -Force` as needed.
    Ensure multi-factor authentication (MFA) is configured if required.

.USAGE
    1. Launch PowerShell as Administrator (if necessary)
    2. Execute the script: `.\MyScriptName.ps1`
    3. Review output at: <output location or file path>

.EXAMPLE
    PS> .\Generate-TenantReport.ps1
    Generates a tenant-wide inventory workbook including users, groups, and more.

.LICENSE
    Proprietary – Internal use only

#>

## THIS MODULE NEEDS TO LOAD AND CONNECT BEFORE MGGRAPH MODULES ##
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName donf@oliverlawfl.com
 
# Ensure required modules are installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Beta)) {
    Install-Module Microsoft.Graph.Beta -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

# Connect to Microsoft 365 services
Connect-MgGraph -Scopes "User.Read.All", 
                        "Group.Read.All",
                        "Team.ReadBasic.All",
                        "Reports.Read.All", 
                        "Organization.Read.All", 
                        "Directory.Read.All", 
                        "Policy.Read.All",
                        "Microsoft.Graph.Reports"
                


# Output folder
$outputPath = "C:\scripts\mgReports"
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

# 1. Entra Users
Get-MgUser -All | Export-Csv "$outputPath\mgUsers.csv" -NoTypeInformation -Force

# 2. Entra Groups
Get-MgGroup -All | Export-Csv "$outputPath\mgGroups.csv" -NoTypeInformation -Force

# 3. Exchange Unified Groups
Get-UnifiedGroup | Export-Csv "$outputPath\mgUnifiedGroups.csv" -NoTypeInformation -Force

# 4. Exchange Distribution Groups
Get-DistributionGroup | Export-Csv "$outputPath\mgDistributionGroups.csv" -NoTypeInformation -Force

# 5. Microsoft Teams
Get-MgTeam -All | Export-Csv "$outputPath\mgTeams.csv" -NoTypeInformation -Force

# 6. Organization Info
Get-MgOrganization | Export-Csv "$outputPath\mgOrganization.csv" -NoTypeInformation -Force

# 7. Entra Devices
Get-MgDevice -All | Export-Csv "$outputPath\mgDevices.csv" -NoTypeInformation -Force

# 8. Mailboxes
Get-Mailbox | Export-Csv "$outputPath\mailboxes.csv" -NoTypeInformation -Force

# Combine all CSVs into a single Excel workbook
$csvFiles = Get-ChildItem -Path $outputPath -Filter *.csv
$excelPath = "$outputPath\M365_Tenant_Report_$(Get-Date -Format yyyyMMdd_HHmm).xlsx"

foreach ($csv in $csvFiles) {
    $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($csv.Name)
    Import-Csv $csv.FullName | Export-Excel -Path $excelPath -WorksheetName $sheetName -AutoSize -Append
}

Write-Host "✅ Report generated: $excelPath"


##################################################################################################################################

<#
.SYNOPSIS
    Generate a comprehensive Microsoft 365 user activity report.

.DESCRIPTION
    Connects to Microsoft Graph (beta) and retrieves user account details, license info,
    Teams and calendar memberships, and recent sign-in metadata (including IP and location).
    Outputs the data to a CSV file: Users_Enriched.csv

.AUTHOR
    Don

.DATE CREATED
    2025-06-26

.REQUIRED MODULES
    - Microsoft.Graph

.NOTES
    Ensure your account has the following delegated Graph API permissions:
    - User.Read.All
    - Calendars.Read
    - Team.ReadBasic.All
    - AuditLog.Read.All
#>

# Switch to beta profile for sign-in activity
# Select-MgProfile beta
Connect-MgGraph -Scopes "User.Read.All", 
                        "Calendars.Read", 
                        "Team.ReadBasic.All", 
                        "AuditLog.Read.All"

# Build lookup tables for SKUs and service plans
$skuMap = @{}
Get-MgSubscribedSku -All | ForEach-Object {
    $skuMap[$_.SkuId.Guid] = $_.SkuPartNumber
}

# Retrieve users
$users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,CreatedDateTime,SignInActivity

# Pull recent sign-in logs for IP and location mapping
$signIns = Get-MgAuditLogSignIn -All |
    Sort-Object CreatedDateTime -Descending |
    Group-Object UserPrincipalName -AsHashTable -AsString

# Collect enriched user data
$flattened = foreach ($user in $users) {
    $signIn = $signIns[$user.UserPrincipalName]
    $lastSignIn = $signIn | Select-Object -First 1

    # Get license detail
    $licenseDetail = try {
        Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop
    } catch { @() }

    $assignedLicenses = ($licenseDetail | ForEach-Object { $_.SkuPartNumber }) -join ';'

    # Get calendars and joined Teams
    $calendars = try {
        Get-MgUserCalendar -UserId $user.Id -ErrorAction Stop | Select-Object -ExpandProperty Name
    } catch { @() }

    $teams = try {
        Get-MgUserJoinedTeam -UserId $user.Id -ErrorAction Stop | Select-Object -ExpandProperty DisplayName
    } catch { @() }

    [pscustomobject][ordered]@{
        ID                = $user.Id
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Mail              = $user.Mail
        AccountEnabled    = $user.AccountEnabled
        CreatedDateTime   = $user.CreatedDateTime
        LastLogin         = $user.SignInActivity.LastSignInDateTime
        LastLoginIP       = $lastSignIn.IpAddress
        LastLoginCity     = $lastSignIn.Location.City
        LastLoginState    = $lastSignIn.Location.State
        AssignedLicenses  = $assignedLicenses
        Calendars         = $calendars -join ';'
        JoinedTeams       = $teams -join ';'
    }
}

# Export to CSV
$flattened | Export-Csv -Path "c:\scripts\mgReports\Users_Enriched.csv" -NoTypeInformation -Force
Write-Host "`n✅ User report generated: .\Users_Enriched.csv" -ForegroundColor Green