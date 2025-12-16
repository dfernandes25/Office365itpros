
param(
    [Parameter(Mandatory=$true)]
    [string]$InputFolder,

    [Parameter(Mandatory=$true)]
    [string]$OutputFile,

    [switch]$Recurse,

    [switch]$EnableJsonL
)

# --- Guard: local-only ---
# Fail early if path looks like UNC or OneDrive/SharePoint
if ($InputFolder -match '^(\\\\|/)' -or $InputFolder -match 'OneDrive|SharePoint') {
    Write-Error "InputFolder appears to be a network/cloud path. Please copy files to a local folder (e.g., C:\Data\Logs) and re-run."
    exit 1
}

if (-not (Test-Path -LiteralPath $InputFolder -PathType Container)) {
    Write-Error "InputFolder not found: $InputFolder"
    exit 1
}

$searchArgs = @{ LiteralPath = $InputFolder; Filter = '*.json' }
if ($Recurse) { $searchArgs.Recurse = $true }

$files = Get-ChildItem @searchArgs | Sort-Object FullName
if ($files.Count -eq 0) {
    Write-Warning "No JSON files found in: $InputFolder"
    exit 0
}

# Helpers
function Try-ConvertFromJson([string]$text) {
    try { return ($text | ConvertFrom-Json -ErrorAction Stop) }
    catch { return $null }
}

function FirstArrayInObject([object]$root) {
    # Breadth-first search for the first array inside a PSCustomObject
    if ($root -isnot [pscustomobject]) { return $null }
    $q = New-Object System.Collections.Generic.Queue[object]
    $q.Enqueue($root)
    while ($q.Count -gt 0) {
        $cur = $q.Dequeue()
        foreach ($p in $cur.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [object[]] -or $v -is [System.Collections.ArrayList]) { return $v }
            if ($v -is [pscustomobject]) { $q.Enqueue($v) }
        }
    }
    return $null
}

function ParseJsonL([string]$path) {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($line in [System.IO.File]::ReadLines($path)) {
        $t = $line.Trim()
        if ($t.Length -eq 0) { continue }
        $obj = Try-ConvertFromJson $t
        if ($null -ne $obj) { $items.Add($obj) }
    }
    return $items.ToArray()
}

# Main
$all = New-Object System.Collections.Generic.List[object]

foreach ($f in $files) {
    Write-Host "Reading $($f.FullName) ..." -ForegroundColor Cyan

    # Local, non-network read
    $text = $null
    try {
        $text = [System.IO.File]::ReadAllText($f.FullName)
    } catch {
        Write-Warning "Failed to read '$($f.FullName)'. Is it locked or unavailable? Skipping."
        continue
    }

    $root = Try-ConvertFromJson $text

    if ($null -ne $root) {
        if ($root -is [object[]] -or $root -is [System.Collections.ArrayList]) {
            foreach ($i in $root) { $all.Add($i) }
        }
        elseif ($root -is [pscustomobject]) {
            $arr = FirstArrayInObject $root
            if ($null -ne $arr) {
                foreach ($i in $arr) { $all.Add($i) }
            } else {
                Write-Warning "No array found in '$($f.Name)'; skipping."
            }
        }
        else {
            Write-Warning "Unknown JSON root type in '$($f.Name)'; skipping."
        }
    }
    elseif ($EnableJsonL) {
        $lineItems = ParseJsonL $f.FullName
        if ($lineItems.Count -gt 0) {
            foreach ($i in $lineItems) { $all.Add($i) }
        } else {
            Write-Warning "Failed to parse '$($f.Name)' as JSONL; skipping."
        }
    }
    else {
        Write-Warning "ConvertFrom-Json failed for '$($f.Name)'. If it is JSON Lines, re-run with -EnableJsonL."
    }
}

# Write consolidated output as a JSON array
$depth = 100
try {
    $json = $all | ConvertTo-Json -Depth $depth
    [System.IO.File]::WriteAllText($OutputFile, $json, [System.Text.Encoding]::UTF8)
    Write-Host "Done. Merged $($all.Count) items into: $OutputFile" -ForegroundColor Green
} catch {
    Write-Error "Failed to write output to '$OutputFile': $($_.Exception.Message)"
    exit 1
}


