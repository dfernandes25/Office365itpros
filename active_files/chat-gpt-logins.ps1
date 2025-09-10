<#
.SYNOPSIS
    Export suspicious sign-ins from Microsoft Entra using Graph API (Beta).

.DESCRIPTION
    This script uses Microsoft Graph PowerShell SDK to pull *all* sign-in logs 
    from the Beta endpoint for the last 14 days and export suspicious events to CSV.
    Displays a rolling counter while processing records.

.REQUIREMENTS
    Install-Module Microsoft.Graph -Scope AllUsers
    Connect-MgGraph -Scopes "AuditLog.Read.All","Directory.Read.All","IdentityRiskEvent.Read.All"

.OUTPUT
    CSV file: SuspiciousSignIns.csv
#>

## THIS MODULE NEEDS TO LOAD AND CONNECT BEFORE MGGRAPH MODULES ##
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement
#Connect-ExchangeOnline -UserPrincipalName donf@oliverlawfl.com
 
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

# Connect to Graph (interactive login if needed)
Connect-MgGraph -Scopes "AuditLog.Read.All","Directory.Read.All","IdentityRiskEvent.Read.All"
Select-MgProfile -Name beta  # Use beta for detailed sign-in fields

# Define output file
$csvPath = "C:\scripts\auditLogs\ChatGPT-SuspiciousSignIns.csv"
$allRecords = @()

Write-Host "Fetching ALL sign-in logs from Graph Beta for the last 14 days..."

# Get sign-ins for the last 14 days
$startDate = (Get-Date).AddDays(-14).ToString("o")

# Fetch ALL results with paging
$signIns = Get-MgAuditLogSignIn -All -Filter "createdDateTime ge $startDate"

$counter = 0
foreach ($signIn in $signIns) {
    $counter++
    Write-Progress -Activity "Processing sign-ins..." -Status "Processed $counter sign-ins"

    # Check for suspicious indicators
    $isSuspicious = $false

    if ($signIn.riskState -ne "none" -and $null -ne $signIn.riskState) { $isSuspicious = $true }
    if ($signIn.riskLevelDuringSignIn -ne "none" -and $null -ne $signIn.riskLevelDuringSignIn) { $isSuspicious = $true }
    if ($signIn.status.errorCode -ne 0) { $isSuspicious = $true }
    if ($signIn.riskDetail -and $signIn.riskDetail -ne "none") { $isSuspicious = $true }

    if ($isSuspicious) {
        $record = [PSCustomObject]@{
            UserPrincipalName       = $signIn.userPrincipalName
            DisplayName             = $signIn.userDisplayName
            CreatedDateTime         = $signIn.createdDateTime
            Status                  = $signIn.status.additionalDetails
            ErrorCode               = $signIn.status.errorCode
            IPAddress               = $signIn.ipAddress
            Location                = $signIn.location.city + ", " + $signIn.location.countryOrRegion
            DeviceDetail            = $signIn.deviceDetail.displayName
            Browser                 = $signIn.deviceDetail.browser
            OperatingSystem         = $signIn.deviceDetail.operatingSystem
            RiskState               = $signIn.riskState
            RiskLevelDuringSignIn   = $signIn.riskLevelDuringSignIn
            RiskDetail              = $signIn.riskDetail
            MFARequired             = $signIn.authenticationRequirement
            AppDisplayName          = $signIn.appDisplayName
            ConditionalAccessStatus = $signIn.conditionalAccessStatus
            CorrelationId           = $signIn.correlationId
        }

        $allRecords += $record
    }
}

if ($allRecords.Count -gt 0) {
    $allRecords | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Export complete: $csvPath ($($allRecords.Count) suspicious sign-ins)"
} else {
    Write-Host "ℹ️ No suspicious sign-ins found in the last 14 days."
}
