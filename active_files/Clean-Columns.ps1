<#
.SYNOPSIS
    Cleans multiple CSV files in a folder by removing unwanted columns.

.DESCRIPTION
    - Iterates through all CSV files in a folder.
    - Removes columns if:
        - All row values are 0, "0", null, or empty.
        - Any row contains unnested objects (e.g., "[]", "System.Object[]", "@{…}", JSON-like structures).
        - All row values are identical.
        - Any row value contains the term 'object' (case-insensitive).
        - Column header is 'Additional Properties' (case-insensitive).
    - Saves cleaned CSV files with "_cleaned" appended to the original filename.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath
)

# Define patterns indicating unnested or unexpanded object values
$UnnestedPatterns = @(
    '^\s*\[\s*\]\s*$',            # empty array literal: []
    '^\s*@\{.*\}\s*$',            # PowerShell hashtable: @{Key=Value}
    '^\s*\{.*\}\s*$',             # JSON-like object: {...}
    '^\s*\[.*\]\s*$',             # JSON-like array: [...]
    '^\s*System\.[\w\.\[\]]+\s*$',# e.g., System.Object[]
    '^\s*Microsoft\.[\w\.\[\]]+\s*$' # e.g., Microsoft.Graph.* types
)

# Function to check if a value matches any unnested pattern
function Test-IsUnnested {
    param([object]$Value)
    if ($null -eq $Value) { return $false }
    $s = $Value.ToString().Trim()
    foreach ($pattern in $UnnestedPatterns) {
        if ($s -match $pattern) { return $true }
    }
    return $false
}

# Get all CSV files in the folder
$csvFiles = Get-ChildItem -Path $FolderPath -Filter *.csv

foreach ($file in $csvFiles) {
    Write-Host "`n📄 Processing file: $($file.FullName)"

    # Import CSV data
    $data = Import-Csv -Path $file.FullName
    if (-not $data) {
        Write-Host "  ⚠️ Empty or invalid CSV. Skipping."
        continue
    }

    # Get column headers
    $headers = $data[0].PSObject.Properties.Name
    $columnsToKeep = @()

    foreach ($header in $headers) {
        $values = $data | ForEach-Object { $_.$header }

        # Condition 1: Remove column if header is 'Additional Properties' (case-insensitive)
        if ($header -ieq 'Additional Properties') {
            Write-Host "  🗑️ Removing column (header is 'Additional Properties'): $header"
            continue
        }

        # Condition 2: Remove column if all values are 0, "0", null, or empty
        $allEmptyOrZero = $true
        foreach ($v in $values) {
            if ($v -ne $null -and $v.ToString().Trim() -ne "" -and $v.ToString().Trim() -ne "0") {
                $allEmptyOrZero = $false
                break
            }
        }
        if ($allEmptyOrZero) {
            Write-Host "  🗑️ Removing column (all empty/zero): $header"
            continue
        }

        # Condition 3: Remove column if any value is unnested
        $hasUnnested = $false
        foreach ($v in $values) {
            if (Test-IsUnnested $v) {
                $hasUnnested = $true
                break
            }
        }
        if ($hasUnnested) {
            Write-Host "  🗑️ Removing column (unnested values detected): $header"
            continue
        }

        # Condition 4: Remove column if all values are identical
        $uniqueValues = $values | Select-Object -Unique
        if ($uniqueValues.Count -eq 1) {
            Write-Host "  🗑️ Removing column (all values identical): $header"
            continue
        }

        # Condition 5: Remove column if any value contains 'object' (case-insensitive)
        $containsObject = $false
        foreach ($v in $values) {
            if ($v -ne $null -and $v.ToString().ToLower().Contains("object")) {
                $containsObject = $true
                break
            }
        }
        if ($containsObject) {
            Write-Host "  🗑️ Removing column (contains 'object'): $header"
            continue
        }

        # Keep the column
        $columnsToKeep += $header
    }

    # Select only the kept columns
    $cleanedData = $data | Select-Object $columnsToKeep

    # Save cleaned CSV
    $outPath = Join-Path $file.DirectoryName ($file.BaseName + "_cleaned.csv")
    $cleanedData | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8

    Write-Host "  ✅ Cleaned CSV saved to: $outPath"
}

