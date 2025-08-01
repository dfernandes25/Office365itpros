# Smart Connect Script for Exchange Online
# Author: BTGB Cloud Team

Write-Host "Attempting to connect to Exchange Online..." -ForegroundColor Cyan

# Try modern authentication with fallback options
try {
    # Preferred method: Device Authentication (no UI dependency)
    Connect-ExchangeOnline -UseDeviceAuthentication -ErrorAction Stop
    Write-Host "✅ Connected using Device Authentication." -ForegroundColor Green
}
catch {
    Write-Warning "Device Authentication failed. Trying Web Login..."

    try {
        # Fallback: Web Login (opens browser)
        Connect-ExchangeOnline -UseWebLogin -ErrorAction Stop
        Write-Host "✅ Connected using Web Login." -ForegroundColor Green
    }
    catch {
        Write-Error "❌ All connection methods failed. Please check your network, credentials, or module version."
        Write-Host "Tip: You can disable WAM by running:" -ForegroundColor Yellow
        Write-Host '[System.Environment]::SetEnvironmentVariable("MSAL_PREFER_WAM", "false", "User")'
    }
}
