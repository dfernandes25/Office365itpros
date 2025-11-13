<# 
collects eol, odb, and teams data
# used with Analyze-MailTraffic.ps1 for data collection
# output used for 20250731-wip.rmd which generates a report

Install-Module Microsoft.Graph -RequiredVersion 2.29.1 -AllowClobber -Force
Install-Module Microsoft.Graph.Beta -AllowClobber -Force
Install-Module ExchangeOnlineManagement -AllowClobber -Force

# Connect-MgGraph -Scopes Reports.Read.All, ReportSettings.ReadWrite.All, User.Read.All, AuditLog.Read.All, Directory.Read.All, SignIn.Read.All
#>

Connect-MgGraph -Scopes @(
    "Reports.Read.All",
    "ReportSettings.ReadWrite.All",
    "User.Read.All"
    "AuditLog.Read.All",
    "Directory.Read.All",
    "SignIn.Read.All"
)


$dirPath = "C:\scripts\m365Pros"
if($dirPath){Remove-Item -Path "$dirPath" -Recurse -Force}
Sleep -Seconds 3
mkdir $dirPath -Force


$ObscureFlag = $false
$Uri = "https://graph.microsoft.com/beta/admin/reportSettings"
# Check if the tenant has obscured real names for reports - see https://office365itpros.com/2022/09/09/graph-usage-report-tips/
$DisplaySettings = Invoke-MgGraphRequest -Method Get -Uri $Uri
If ($DisplaySettings['displayConcealedNames'] -eq $true) { # data is obscured, so let's reset it to allow the report to run
   $ObscureFlag = $true
   Write-Host "Setting tenant data concealment for reports to False" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body (@{"displayConcealedNames"= $false} | ConvertTo-Json) 
}

# Specifies  SKU identifers for Office 365 and Microsoft 365 E3. There are other variants of these SKUs
# for government and academic use, so it's important to pass the SKU identifiers in use within the tenant
Write-Host "Finding user accounts to check..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/any(s:s/skuId eq 6fd2c87f-b296-42f0-b197-1e91e994b900) `
    or assignedLicenses/any(s:s/skuid eq c7df2760-2c81-4ef7-b578-5b5392b571df) `
    or assignedLicenses/any(s:s/skuid eq 05e9a617-0261-4cee-bb44-138d3ef5d965) `
    or assignedLicenses/any(s:s/skuid eq 06ebc4ee-1bb5-47dd-8120-11324bc54e06)" `
    -ConsistencyLevel Eventual -CountVariable Licenses -All -Sort 'displayName' `
    -Property Id, displayName, signInActivity, userPrincipalName -PageSize 999

Write-Host "Fetching usage data for Teams, Exchange, and OneDrive for Business..."
# Get Teams user activity detail for the last 30 days
$Uri = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D30')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath "$dirPath\teams.csv"

# Get Email activity data
$Uri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D30')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath "$dirPath\email.csv"

# Get OneDrive data 
$Uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserDetail(period='D30')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath "$dirPath\onedrive.csv"

# Get Apps detail
$Uri = "https://graph.microsoft.com/v1.0/reports/getM365AppUserDetail(period='D30')"
Invoke-mgGraphRequest -Uri $Uri -Method GET -OutputFilePath "$dirPath\apps.csv"


# Fetch recent sign-in logs
$signIns = Get-MgAuditLogSignIn -Top 500

# Get suspicious logins
$suspiciousLogins = $signIns | Where-Object {
    $_.Status.ErrorCode -ne 0 -or
    $_.ConditionalAccessStatus -eq "failure"
}

$suspiciousLogins | 
Select-Object UserPrincipalName, CreatedDateTime, IPAddress, 
    @{Name="City";Expression={$_.Location.City}}, @{Name="Country";Expression={$_.Location.CountryOrRegion}}, 
    @{Name="FailureReason";Expression={$_.Status.FailureReason}}, ConditionalAccessStatus | Export-csv "$dirPath\suspiciousLogins.csv" -nti -Force
    

# Switch the tenant report obscure data setting back if necessary
If ($ObscureFlag -eq $True) {
    Write-Host "Resetting tenant data concealment for reports to True" -foregroundcolor red
    Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' `
     -Body (@{"displayConcealedNames"= $true} | ConvertTo-Json) 
     }

