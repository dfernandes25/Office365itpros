


# Load and parse the JSON file created from 
$jsonPath = "C:\scripts\mext\20250807145621-DirectoryAuditLogs.json"
$auditLogs = Get-Content $jsonPath | ConvertFrom-Json

# Prepare an array to hold parsed entries
$parsedEntries = @()

foreach ($entry in $auditLogs) {
    $parsedEntries += [PSCustomObject]@{
        id                             = $entry.id
        category                       = $entry.category
        activityDisplayName            = $entry.activityDisplayName
        activityDateTime               = $entry.activityDateTime
        result                         = $entry.result
        resultReason                   = $entry.resultReason
        initiatedBy_userPrincipalName = $entry.initiatedBy.user.userPrincipalName
        initiatedBy_appDisplayName    = $entry.initiatedBy.app.displayName
        targetResource_userPrincipalName = $entry.targetResources[0].userPrincipalName
        targetResource_type           = $entry.targetResources[0].type
    }
}

# Export to CSV
$csvPath = "C:\scripts\mext\DirAuditLog_Parsed.csv"
$parsedEntries | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8




## UNIFIED AUDIT LOG ##
# Define the path to the JSON file
$jsonPath = "C:\scripts\mext\UnifiedAuditLog.json"

# Load the JSON content
$auditLogs = Get-Content $jsonPath | ConvertFrom-Json

# Prepare an array to hold parsed entries
$parsedEntries = @()

foreach ($entry in $auditLogs) {
    $auditData = $entry.auditData

    $parsedEntries += [PSCustomObject]@{
        Id                  = $entry.id
        CreatedDateTime     = $entry.createdDateTime
        Operation           = $entry.operation
        Service             = $entry.service
        UserPrincipalName   = $entry.userPrincipalName
        ClientIPAddress     = $auditData.ClientIPAddress
        ResultStatus        = $auditData.ResultStatus
        MailboxOwnerUPN     = $auditData.MailboxOwnerUPN
        MailAccessType      = ($auditData.OperationProperties | Where-Object { $_.Name -eq "MailAccessType" }).Value
        FolderPath          = ($auditData.Folders | Select-Object -First 1).Path
        ItemSizeInBytes     = ($auditData.Folders[0].FolderItems[0].SizeInBytes)
        InternetMessageId   = ($auditData.Folders[0].FolderItems[0].InternetMessageId)
    }
}

# Export to CSV
$csvPath = "C:\scripts\mext\UAL_parsed.csv"
$parsedEntries | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Audit log data exported to $csvPath"

