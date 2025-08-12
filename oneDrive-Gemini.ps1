function Get-AllUsersOneDriveSharingReport {
    param (
        [string]$DomainName,
        [datetime]$StartDate = (Get-Date).AddDays(-30),
        [datetime]$EndDate = (Get-Date)
    )

    # Connect to Microsoft Graph with necessary scopes
    # Note: 'Sites.Read.All', 'Files.Read.All', and 'User.Read.All' are required.
    Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All", "User.Read.All"

    # Get all users in the tenant
    $users = Get-MgUser -All

    $report = @()

    foreach ($user in $users) {
        Write-Host "Processing OneDrive for $($user.UserPrincipalName)..."

        try {
            $drive = Get-MgUserDrive -UserId $user.Id -ErrorAction SilentlyContinue
        
            # IMPORTANT MODIFICATION:
            # Replaced 'Get-MgDriveRootChild' with a direct API call to the search endpoint.
            # This is a workaround to include hidden files, as the standard 'Get-MgDriveRootChild'
            # cmdlet and its underlying API endpoint do not have a parameter for this.
            # The search endpoint with an empty query (q='') returns all items recursively,
            # which effectively includes hidden files that would otherwise be excluded.

            # Construct the API URI for the search endpoint
            $uri = "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/root/search(q='')?$select=name,webUrl,size,lastModifiedDateTime,lastModifiedBy,permissions,fileSystemInfo"

            # Invoke the Graph API request to get all files, including hidden ones
            $allDriveItems = Invoke-MgGraphRequest -Uri $uri -Method GET

            # The response contains a 'value' property which holds the list of items
            $files = $allDriveItems.value

            foreach ($file in $files) {
                # Get permissions for the file
                # The search endpoint response already includes permissions as a nested property.
                # We will check for permissions directly on the file object.
                $permissions = $file.permissions

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
                    IsHidden          = $file.FileSystemInfo.IsHidden # New property to show if file is hidden
                }

                $report += $entry
            }
        } catch {
            Write-Warning "Failed to access OneDrive for $($user.UserPrincipalName): $_"
        }
    }

    # Output or export the report
    $report | Export-Csv -Path "C:\scripts\mgReports\mgOneDriveSharing.csv" -NoTypeInformation -Force
    Write-Host "Report saved to C:\scripts\mgReports\mgOneDriveSharing.csv"
}

Get-AllUsersOneDriveSharingReport -DomainName "oliverlawfl.com"
