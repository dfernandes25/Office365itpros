


## used with mext script to copy all csv files to central location

$OutputDir = 'C:\scripts\data'
if (!(Test-Path $OutputDir)) {mkdir $OutputDir -Force}



## copy mext csv files to a data folder ##
function Reset-Folder {
    param (
        [string]$ParentPath = "C:\Scripts",
        [string]$FolderName = "Mext"
    )

    $fullPath = Join-Path $ParentPath $FolderName

    if (Test-Path $fullPath) {
        Write-Host "Folder '$fullPath' exists. Clearing contents..." -ForegroundColor Yellow
        try {
            Get-ChildItem -Path $fullPath -Recurse -Force -Filter '*.csv*'  | Copy-Item -Destination $OutputDir  -Recurse -Force
            Write-Host "Contents of '$fullPath' successfully copied." -ForegroundColor Green
        } catch {
            Write-Host "Error copying contents of '$fullPath': $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Folder '$fullPath' does not exist. Creating it..." -ForegroundColor Cyan
        try {
            New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
            Write-Host "Folder '$fullPath' created successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error creating folder '$fullPath': $_" -ForegroundColor Red
        }
    }
    }

    Reset-Folder



## if Unified Access Logs were pulled, combine them into a single csv ##

$OutputFile = "$OutputDir\UAL-Combined.csv"
$files = Get-ChildItem -Path $OutputDir -Filter "UAL-*.csv"


# Get first file’s header, then append remaining files without header
$First = $true
Remove-Item $OutputFile -ErrorAction SilentlyContinue

foreach ($File in $Files) {
    if ($First) {
        Get-Content $File.FullName | Out-File $OutputFile
        $First = $false
    } else {
        Get-Content $File.FullName | Select-Object -Skip 1 | Out-File $OutputFile -Append
    }
}

## Remove the raw UAL log files ##
Get-ChildItem -Path $OutputDir -Filter "UAL-2*.csv" | Remove-Item -Force