Add-Type -AssemblyName System.IO.Compression.FileSystem
$xlpath = 'C:\Users\WilliamMorris\OneDrive - Gurner TM\Projects - 189 Kent Street, Sydney\04. Legals\05. Off the Plan Contract of Sale\COS Tracker\189Kent_CoSTracker.xlsx'
$tmppath = [System.IO.Path]::GetTempFileName() + '.xlsx'
[System.IO.File]::Copy($xlpath, $tmppath, $true)
$zip = [System.IO.Compression.ZipFile]::OpenRead($tmppath)

# Get shared string #304 (last one, 0-indexed)
$ssEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }
if ($ssEntry) {
    $reader = New-Object System.IO.StreamReader($ssEntry.Open())
    $ss = $reader.ReadToEnd()
    $reader.Close()
    # Get last 500 chars to find string #304
    Write-Output "=== END OF SHARED STRINGS ==="
    Write-Output $ss.Substring([Math]::Max(0, $ss.Length - 1000))
}

$zip.Dispose()
