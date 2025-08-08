# Install Microsoft Graph SDK if not present (uncomment if needed)
# Update-Module Microsoft.Graph -Force

# Connect to Microsoft Graph with required scopes
Connect-MgGraph -Scopes "AuditLog.Read.All","Directory.Read.All","Policy.Read.ConditionalAccess"

$ObscureFlag = $false
$Uri = "https://graph.microsoft.com/beta/admin/reportSettings"
# Check if the tenant has obscured real names for reports - see https://office365itpros.com/2022/09/09/graph-usage-report-tips/
$DisplaySettings = Invoke-MgGraphRequest -Method Get -Uri $Uri
If ($DisplaySettings['displayConcealedNames'] -eq $true) { # data is obscured, so let's reset it to allow the report to run
   $ObscureFlag = $true
   Write-Host "Setting tenant data concealment for reports to False" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body (@{"displayConcealedNames"= $false} | ConvertTo-Json) 
}

# Retrieve Entra (Azure AD) sign-in logs (adjust filters as needed for deep analysis)
$signins = Get-MgAuditLogSignIn -All

# Unnest relevant object properties for clean CSV export
$results = $signins | ForEach-Object {
    [PSCustomObject]@{
        SignInDate      = $_.CreatedDateTime
        UserDisplayName = $_.UserDisplayName
        UserPrincipal   = $_.UserPrincipalName
        Status          = if ($_.Status.ErrorCode -eq 0) { "Success" } else { "Failed" }
        IPAddress       = $_.IpAddress
        Location        = "$($_.Location.City), $($_.Location.State), $($_.Location.CountryOrRegion)"
        DeviceName      = $_.DeviceDetail.DisplayName
        Browser         = $_.DeviceDetail.Browser
        OperatingSystem = $_.DeviceDetail.OperatingSystem
        RiskDetail      = $_.RiskDetail
        RiskState       = $_.RiskState
        ConditionalAccessStatus = $_.ConditionalAccessStatus
        AppliedCAPolicies = ($_.AppliedConditionalAccessPolicies | ForEach-Object { $_.DisplayName }) -join ", "
    }
}

# Export to CSV (results will be unnested – each property flat in the CSV)
$exportPath = "C:\scripts\m365Pros\EntraSignIns.csv"
$results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -Force

Write-Host "Export complete: $exportPath"

# Switch the tenant report obscure data setting back if necessary
If ($ObscureFlag -eq $True) {
    Write-Host "Resetting tenant data concealment for reports to True" -foregroundcolor red
    Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' `
     -Body (@{"displayConcealedNames"= $true} | ConvertTo-Json) 
     }

Disconnect-MgGraph
