function Get-AllUsersOneDriveSharingReport {
    param (
        [string]$DomainName,
        [datetime]$StartDate = (Get-Date).AddDays(-30),
        [datetime]$EndDate = (Get-Date)
    )

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All", "User.Read.All"

    # Get all users in the tenant
    $users = Get-MgUser -All

    $report = @()

    foreach ($user in $users) {
        Write-Host "Processing OneDrive for $($user.UserPrincipalName)..."

        try {
            $drive = Get-MgUserDrive -UserId $user.Id -ErrorAction SilentlyContinue
            $files = Get-MgDriveRootChild -DriveId $drive.Id -All

            foreach ($file in $files) {
                $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -ItemId $file.Id -ErrorAction SilentlyContinue

                $sharingStatus = "Private"
                $sharedWith = @()
                $sharedDate = $null

                foreach ($perm in $permissions) {
                    if ($perm.Link) {
                        if ($perm.Link.Scope -eq "anonymous") {
                            $sharingStatus = "Shared Externally"
                            $sharedWith += "Anonymous Link"
                        } elseif ($perm.Link.Scope -eq "organization") {
                            $sharingStatus = "Shared Internally"
                            $sharedWith += "Org Link"
                        }
                        $sharedDate = $perm.CreatedDateTime
                    } elseif ($perm.Grantee) {
                        $email = $perm.Grantee.User.Email
                        if ($email -like "*@$DomainName") {
                            $sharingStatus = "Shared Internally"
                        } else {
                            $sharingStatus = "Shared Externally"
                        }
                        $sharedWith += $email
                        $sharedDate = $perm.CreatedDateTime
                    }
                }

                $entry = [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    FileName          = $file.Name
                    FilePath          = $file.WebUrl
                    SizeMB            = [math]::Round($file.Size / 1MB, 2)
                    LastModified      = $file.LastModifiedDateTime
                    LastModifiedBy    = $file.LastModifiedBy?.User?.DisplayName
                    SharingStatus     = $sharingStatus
                    SharedWith        = ($sharedWith -join ", ")
                    SharedDate        = $sharedDate
                }

                $report += $entry
            }
        } catch {
            Write-Warning "Failed to access OneDrive for $($user.UserPrincipalName): $_"
        }
    }

    # Output or export the report
    $report | Export-Csv -Path "C:\scripts\mgReports\mgOneDriveSharing.csv" -NoTypeInformation -Force
    Write-Host "Report saved to AllUsers_OneDriveSharingReport.csv"
}

Get-AllUsersOneDriveSharingReport -DomainName "oliverlawfl.com"
