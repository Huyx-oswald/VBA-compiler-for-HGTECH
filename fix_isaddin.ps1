# fix_isaddin.ps1 - Add isAddin="true" to workbook.xml inside .xlam ZIP
# Usage: powershell -ExecutionPolicy Bypass -File fix_isaddin.ps1 <xlamPath>
# WPS SaveAs format 55 doesn't write isAddin, so we patch the ZIP directly

param(
    [Parameter(Mandatory=$true)]
    [string]$XlamPath
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path $XlamPath)) {
    Write-Error "File not found: $XlamPath"
    exit 1
}

try {
    $zip = [System.IO.Compression.ZipFile]::Open($XlamPath, [System.IO.Compression.ZipArchiveMode]::Update)
} catch {
    Write-Error "Cannot open ZIP: $_"
    exit 1
}

$entry = $zip.GetEntry("xl/workbook.xml")
if ($entry -eq $null) {
    Write-Error "xl/workbook.xml not found in ZIP"
    $zip.Dispose()
    exit 1
}

# Read current content
$reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
$content = $reader.ReadToEnd()
$reader.Close()

if ($content -match 'isAddin="true"') {
    Write-Output "isAddin already present, no change needed"
    $zip.Dispose()
    exit 0
}

# Add isAddin="true" to <workbookPr
if ($content -match '<workbookPr') {
    $content = $content -replace '<workbookPr', '<workbookPr isAddin="true"'
} else {
    Write-Error "workbookPr element not found"
    $zip.Dispose()
    exit 1
}

# Delete old entry and create new one with modified content
$entry.Delete()
$newEntry = $zip.CreateEntry("xl/workbook.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$stream = $newEntry.Open()
$writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
$writer.Write($content)
$writer.Close()
$stream.Close()
$zip.Dispose()

Write-Output "OK: isAddin=true added to workbook.xml"
