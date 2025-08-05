

$files = gci C:\scripts\dmarc -Filter *.csv

# Initialize an empty array to store imported CSV data
$csvData = @()

# Import each CSV file and add it to the array
foreach ($file in $files) {
    $csvData += Import-Csv -Path $file.FullName
}

# Combine the CSV files by column
$mergedData = $csvData[0] # Start with the first file's data
for ($i = 1; $i -lt $csvData.Count; $i++) {
    $mergedData = $mergedData | ForEach-Object {
        $row = $_
        $csvData[$i] | ForEach-Object {
            $row += $_
        }
        $row
    }
}

# Export the merged data to a new CSV file
$mergedData | Export-Csv C:\scripts\dmarc\merged.csv -NoTypeInformation
