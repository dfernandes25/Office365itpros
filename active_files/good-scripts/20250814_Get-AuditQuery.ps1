<# https://practical365.com/auditlog-query-api-deeper-look/ #>

# Connect to Microsoft 365 services
Connect-MgGraph -Scopes "User.Read.All", 
                        "Group.Read.All",
                        "Team.ReadBasic.All",
                        "Reports.Read.All", 
                        "Organization.Read.All", 
                        "Directory.Read.All", 
                        "Policy.Read.All",
                        "Microsoft.Graph.Reports",
                        "AuditLogsQuery.Read.All"

##### AUDIT QUERY PARAMETERS
$AuditJobName = ("Full audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date).AddHours(1)
$StartDate = (Get-Date $EndDate).AddDays(-1)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)

##### CREATE AUDIT JOB
$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters

##### WAIT FOR JOB TO COMPLETE
[int]$i = 1
[int]$SleepSeconds = 20
$SearchFinished = $false; [int]$SecondsElapsed = 20
Write-Host "Checking audit query status..."
Start-Sleep -Seconds 30
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}" -f $AuditJob.id)
$AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get

While ($SearchFinished -eq $false) {
    $i++
    Write-Host ("Waiting for audit search to complete. Check {0} after {1} seconds. Current state {2}" -f $i, $SecondsElapsed, $AuditQueryStatus.status)
    If ($AuditQueryStatus.status -eq 'succeeded') {
        $SearchFinished = $true
    } Else {
        Start-Sleep -Seconds $SleepSeconds
        $SecondsElapsed = $SecondsElapsed + $SleepSeconds
        $AuditQueryStatus = Invoke-MgGraphRequest -Uri $Uri -Method Get
    }
}

##### FETCH AUDIT RECORDS
$AuditRecords = [System.Collections.Generic.List[string]]::new()
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$top=999" -f $AuditJob.Id)
[array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET
[array]$AuditRecords = $AuditSearchRecords.value

$NextLink = $AuditSearchRecords.'@Odata.NextLink'
While ($null -ne $NextLink) {
    $AuditSearchRecords = $null
    [array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $NextLink -Method GET 
    $AuditRecords += $AuditSearchRecords.value
    Write-Host ("{0} audit records fetched so far..." -f $AuditRecords.count)
    $NextLink = $AuditSearchRecords.'@odata.NextLink' 
}

Write-Host ("Audit query {0} returned {1} records" -f $AuditJobName, $AuditRecords.Count)

$AuditRecords | Group-Object Operation -NoElement | Sort-Object Count -Descending | 
Select-Object Count, Name | Export-Csv "C:\scripts\auditLogs\ops-count.csv" -nti -Force  #Format-Table -AutoSize

##### JOB WITH FILTERS

$AuditJobName = ("Audit job created at {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm'))
$EndDate = (Get-Date).AddHours(1)
$StartDate = (Get-Date $EndDate).AddDays(-10)
$AuditQueryStart = (Get-Date $StartDate -format s)
$AuditQueryEnd = (Get-Date $EndDate -format s)
[array]$AuditOperationFilters = "FileModified", 
                                "FileDeleted", 
                                "FileUploaded", 
                                "FileRenamed",
                                "FileRecycled"

#[array]$AuditobjectIdFilters = "https://office365itpros.sharepoint.com/sites/productcreation/*", "https://office365itpros.sharepoint.com/sites/Office365Adoption/*"
#[array]$AuditUserPrincipalNameFilters = "Ken.Bowers@office365itpros.com", "Lotte.Vetler@office365itpros.com", "tony.redmond@office365itpros.com"

$AuditQueryParameters = @{}
$AuditQueryParameters.Add("@odata.type","#microsoft.graph.security.auditLogQuery")
$AuditQueryParameters.Add("displayName", $AuditJobName)
$AuditQueryParameters.Add("OperationFilters", $AuditOperationFilters)
$AuditQueryParameters.Add("filterStartDateTime", $AuditQueryStart)
$AuditQueryParameters.Add("filterEndDateTime", $AuditQueryEnd)
#$AuditQueryParameters.Add("userPrincipalNameFilters", $AuditUserPrincipalNameFilters)
#$AuditQueryParameters.Add("objectIdFilters", $AuditobjectIdFilters)

$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$AuditJob = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $AuditQueryParameters


##### PARSE AuditData
# This section assumes you have a CSV file with an AuditData column that contains JSON strings.
# Adjust the path to your CSV file as needed.
$data = Import-Csv -Path "C:\Users\cbc\Downloads\fileOperations.csv"
$parsed = @()
foreach ($row in $data) {
    try {
        $audit = $row.AuditData | ConvertFrom-Json
        $parsed += $audit
    } catch {}
}
$parsed | Export-Csv -Path "C:\scripts\auditLogs\parsed-auditdata.csv" -NoTypeInformation


##### FIND EXISTING AUDIT JOBS
$Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
$Data = Invoke-MgGraphRequest -Uri $Uri -Method GET
If ($Data) {
    Write-Output "Audit Jobs found"
    $Data.Value | ForEach-Object {
        Write-Host ("{0} {1}" -f $_.id, $_.displayName)
    }
} Else {
    Write-Output "No audit jobs found"
}

##### FETCH AUDIT RECORDS COUNT FOR JOB
$AuditJob = "70715017-a51a-4fb0-86ed-1176928e0919" #sample job id from above
$Uri = ("https://graph.microsoft.com/beta/security/auditLog/queries/{0}/records?`$top=999" -f $AuditJob)
[array]$AuditSearchRecords = Invoke-MgGraphRequest -Uri $Uri -Method GET

