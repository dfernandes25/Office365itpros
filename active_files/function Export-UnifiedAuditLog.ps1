function Export-UnifiedAuditLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Last 1 Day", "Last 7 Days", "Last 30 Days")]
        [string]$DateRange,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    # Calculate date range
    switch ($DateRange) {
        "Last 1 Day"   { $startDate = (Get-Date).AddDays(-1) }
        "Last 7 Days"  { $startDate = (Get-Date).AddDays(-7) }
        "Last 30 Days" { $startDate = (Get-Date).AddDays(-30) }
    }
    $endDate = Get-Date

    Write-Host "Searching Unified Audit Log from $startDate to $endDate..."

    # Search the audit log
    $results = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -ResultSize 5000

    if ($results.Count -eq 0) {
        Write-Warning "No audit log entries found for the specified date range."
        return
    }

    # Export to CSV
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "Export complete. File saved to: $OutputPath"
}
Export-UnifiedAuditLog -DateRange "Last 7 Days" -OutputPath "C:\scripts\mext\UAL_7Day.csv"
