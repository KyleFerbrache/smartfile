<#
.SYNOPSIS
Ensures XLSX files in the locally-synced SharePoint library
contain indexable metadata corresponding to their file name

.DESCRIPTION
- Scans a configured directory recursively to identify XLSX candidate files
- Filters candidates based on lastModifiedDateTime (optional)
- Updates or creates Title metadata inside the document using OpenXML structures
- Uses direct XML manipulation for performance

.PARAMETER RootPath
Root Directory to Scan.
This should be "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources"

.PARAMETER Days
Only process files modified in the last X days (optional)
If not provided, the script will evaluate all files in the RootPath

.EXAMPLE
.\Set-SharePointIndex.ps1 -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" -Days 7

.EXAMPLE
.\Set-SharePointIndex.ps1 - RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources"
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$RootPath,

    [int]$Days
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a new OpenXML content body for `docProps/core.xml`
function New-CoreXmlContent {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Title
    )
    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<cp:coreProperties 
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>$escapedTitle</dc:title>
</cp:coreProperties>
"@
}

# Determine the earliest lastModifiedDateTime (if -Days is provided)
$cutoffDate = $null
if ($PSBoundParameters.ContainsKey("Days")) {
    $cutoffDate = (Get-Date).AddDays(-$Days)
    Write-Host "Scanning *.xlsx Files in $RootPath - Modified After: $cutoffDate"
}
else {
    Write-Host "Scanning *.xlsx Files in $RootPath"
}

# Phase 1 - File Discovery
$files = Get-ChildItem -Path $RootPath -Recurse -Filter *.xlsx -File | Where-Object {
    if (-not $cutoffDate) { return $true }
    $_.LastWriteTime -ge $cutoffDate
}

$totalFiles = $files.Count

if ($totalFiles -eq 0) {
    Write-Host "No Files Found Matching Provided Criteria" -ForegroundColor Yellow
    return
}

Write-Host "Found $totalFiles eligible files. Processing..." -ForegroundColor Cyan

# Processing Counters
$updated = 0
$created = 0
$skipped = 0
$errored = 0

# Phase 2 - Conditional Processing
$index = 0

foreach ($file in $files) {
    $index++

    Write-Progress `
        -Activity "Processing Eligible Files" `
        -Status "$index / $totalFiles" `
        -PercentComplete (($index/$totalFiles) * 100)

    $filePath = $file.FullName
    $baseName = $file.$baseName

    try {
        $zip = [System.IO.Compression.ZipFile]::Open($filePath, 'Update')
        $entry = $zip.GetEntry("docProps/core.xml")

        # Create XML if does not exist
        if (-not $entry) {
            $entry = $zip.CreateEntry("docProps/core.xml")
            $xmlContent = New-CoreXmlContent -Title $baseName
            $writer = New-Object System.IO.StreamWriter($stream)
            $writer.Write($xmlContent)
            $writer.close()
            
            $created++
        }

        # Evaluate XML if exists
        else {
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $xmlContent = $reader.ReadToEnd()
            $reader.Close()

            [xml]$xml = $xmlContent
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNameSpace("dc", "http://purl.org/dc/elements/1.1")
            $ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
            
            $titleNode = $xml.SelectSingleNode("//dc:title", $ns)

            # If Title is Blank/Null - Write New Tag
            if ($null -eq $titleNode) {
                $titleNode = $xml.CreateElement("dc", "title", $ns.LookupNamespace("dc"))
                $xml.DocumentElement.AppendChild($titleNode) | Out-Null
                $titleNode.InnerText = $baseName
                $updated++
            }

            # If Title otherwise does not equal filename - Overwrite Tag Value
            elseif ($titleNode -ne $baseName) {
                $titleNode.InnerText = $baseName
                $updated++
            }

            # Title matches filename - Do not edit
            else {
                $skipped++
                $zip.Dispose()
                continue
            }

            # Write Back XML
            $stream = $entry.Open()
            $writer = New-Object System.IO.StreamWriter($stream)
            $writer.BaseStream.SetLength(0)
            $xml.Save($writer)
            $writer.Close()
        }
        $zip.Dispose()
    }
    catch {
        $errors++
    }
}

# Complete the progress bar
Write-Progress -Activity "Processing Eligible Files" -Completed

# Write Change Summary Logs
Write-Host "`n==== Change Summary ====" -ForegroundColor Cyan
Write-Host "Files Scanned: $totalFiles"
Write-Host "Files Updated: $updated"
Write-Host "Files Created: $created"
Write-Host "Files Skipped: $skipped"
Write-Host "Files Errored: $errored"
