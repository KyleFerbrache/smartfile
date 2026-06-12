<#
.SYNOPSIS
Hard-repairs XLSX OpenXML core properties metadata.

.DESCRIPTION
This script repairs XLSX files by rebuilding the package and replacing the
core properties metadata with valid OpenXML metadata.

It:
- Recursively scans XLSX files under a locally synced SharePoint path
- Rebuilds each XLSX into a temporary package
- Removes existing docProps/core.xml
- Removes duplicate/stale core-properties content type overrides
- Removes duplicate/stale core-properties package relationships
- Writes one clean docProps/core.xml
- Sets dc:title equal to the XLSX filename without extension
- Sets created and modified timestamps to current UTC time
- Optionally creates .bak backups
- Supports dry-run mode with -DoNotApply

.PARAMETER RootPath
Root directory relative to $HOME.

Example:
"Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources"

.PARAMETER Days
Optional. Only process files modified in the last X days.
If omitted, all XLSX files under RootPath are processed.

.PARAMETER DoNotApply
Dry run. Reports what would be repaired without modifying files.

.PARAMETER Backup
Creates a .bak copy before replacing each XLSX.

.PARAMETER Creator
Value for dc:creator. Defaults to "SAP UI5 Document Export".

.PARAMETER Keywords
Value for cp:keywords. Defaults to "SAP UI5 EXPORT".

.PARAMETER LastModifiedBy
Value for cp:lastModifiedBy. Defaults to the current Windows user display fallback.

.EXAMPLE
.\Repair-XlsxCoreProperties.ps1 `
  -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" `
  -DoNotApply

.EXAMPLE
.\Repair-XlsxCoreProperties.ps1 `
  -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" `
  -Backup

.EXAMPLE
.\Repair-XlsxCoreProperties.ps1 `
  -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" `
  -Days 7 `
  -Backup
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RootPath,

    [ValidateRange(0, 3650)]
    [int]$Days,

    [switch]$DoNotApply,

    [switch]$Backup,

    [string]$Creator = "SAP UI5 Document Export",

    [string]$Keywords = "SAP UI5 EXPORT",

    [string]$LastModifiedBy = $env:USERNAME
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($PSVersionTable.PSVersion.Major -ge 7) {
    $PSStyle.Progress.View = 'Minimal'
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Resolve root path relative to user profile/home.
$ResolvedRootPath = Join-Path $HOME $RootPath

# UTF-8 without BOM, matching the Excel-repaired sample.
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

# OpenXML namespaces and content types.
$ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types"
$RelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships"

$CoreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
$CoreContentType = "application/vnd.openxmlformats-package.core-properties+xml"

$CorePropertiesPath = "docProps/core.xml"
$ContentTypesPath = "[Content_Types].xml"
$PackageRelsPath = "_rels/.rels"

function Get-SafeTempPath {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $guid = [System.Guid]::NewGuid().ToString("N")
    return Join-Path $File.DirectoryName ".$($File.Name).repair-$guid.tmp"
}

function Get-ZipEntryText {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)

    if (-not $entry) {
        return $null
    }

    $stream = $entry.Open()

    try {
        $reader = [System.IO.StreamReader]::new($stream, $true)

        try {
            return [string]$reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Set-ZipEntryText {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName,

        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()

    try {
        $writer = [System.IO.StreamWriter]::new($stream, $Utf8NoBom)

        try {
            $writer.Write($Text)
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function ConvertTo-XmlDocument {
    param (
        [Parameter(Mandatory = $true)]
        [string]$XmlText
    )

    if ([string]::IsNullOrWhiteSpace($XmlText)) {
        throw "XML text is empty."
    }

    $doc = [System.Xml.XmlDocument]::new()
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($XmlText)

    return ,$doc
}

function ConvertFrom-XmlDocument {
    param (
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Doc
    )

    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = $Utf8NoBom
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $false
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::None

    $memoryStream = [System.IO.MemoryStream]::new()

    try {
        $writer = [System.Xml.XmlWriter]::Create($memoryStream, $settings)

        try {
            $Doc.Save($writer)
        }
        finally {
            $writer.Dispose()
        }

        return [string]$Utf8NoBom.GetString($memoryStream.ToArray())
    }
    finally {
        $memoryStream.Dispose()
    }
}

function New-CorePropertiesXml {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Creator,

        [Parameter(Mandatory = $true)]
        [string]$Keywords,

        [Parameter(Mandatory = $true)]
        [string]$LastModifiedBy
    )

    # Use current UTC time for both created and modified, per request.
    $now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
    $escapedCreator = [System.Security.SecurityElement]::Escape($Creator)
    $escapedKeywords = [System.Security.SecurityElement]::Escape($Keywords)
    $escapedLastModifiedBy = [System.Security.SecurityElement]::Escape($LastModifiedBy)

    # This mirrors the Excel-repaired structure:
    # - dc:title
    # - dc:creator
    # - cp:keywords
    # - cp:lastModifiedBy
    # - dcterms:created xsi:type="dcterms:W3CDTF"
    # - dcterms:modified xsi:type="dcterms:W3CDTF"
    return [string]@"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>$escapedTitle</dc:title><dc:creator>$escapedCreator</dc:creator><cp:keywords>$escapedKeywords</cp:keywords><cp:lastModifiedBy>$escapedLastModifiedBy</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified></cp:coreProperties>
"@
}

function Get-RepairedContentTypesXml {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OriginalXml
    )

    $doc = ConvertTo-XmlDocument -XmlText $OriginalXml

    if (-not $doc.DocumentElement) {
        throw "Invalid [Content_Types].xml. Missing document element."
    }

    if ($doc.DocumentElement.LocalName -ne "Types") {
        throw "Invalid [Content_Types].xml. Expected root element 'Types'."
    }

    if ($doc.DocumentElement.NamespaceURI -ne $ContentTypesNs) {
        throw "Invalid [Content_Types].xml namespace: $($doc.DocumentElement.NamespaceURI)"
    }

    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("ct", $ContentTypesNs)

    # Remove all existing core.xml overrides so we can add exactly one clean copy.
    $existingCoreOverrides = @(
        $doc.SelectNodes("/ct:Types/ct:Override[@PartName='/docProps/core.xml']", $ns)
    )

    foreach ($node in $existingCoreOverrides) {
        [void]$node.ParentNode.RemoveChild($node)
    }

    $override = $doc.CreateElement("Override", $ContentTypesNs)
    $override.SetAttribute("PartName", "/docProps/core.xml")
    $override.SetAttribute("ContentType", $CoreContentType)

    [void]$doc.DocumentElement.AppendChild($override)

    return [string](ConvertFrom-XmlDocument -Doc $doc)
}

function New-DefaultPackageRelsXml {
    $doc = [System.Xml.XmlDocument]::new()
    $doc.PreserveWhitespace = $false

    $declaration = $doc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
    [void]$doc.AppendChild($declaration)

    $root = $doc.CreateElement("Relationships", $RelationshipsNs)
    [void]$doc.AppendChild($root)

    return ,$doc
}

function Get-RepairedPackageRelsXml {
    param (
        [AllowNull()]
        [string]$OriginalXml
    )

    if ([string]::IsNullOrWhiteSpace($OriginalXml)) {
        $doc = New-DefaultPackageRelsXml
    }
    else {
        $doc = ConvertTo-XmlDocument -XmlText $OriginalXml
    }

    if (-not $doc.DocumentElement) {
        throw "Invalid _rels/.rels. Missing document element."
    }

    if ($doc.DocumentElement.LocalName -ne "Relationships") {
        throw "Invalid _rels/.rels. Expected root element 'Relationships'."
    }

    if ($doc.DocumentElement.NamespaceURI -ne $RelationshipsNs) {
        throw "Invalid _rels/.rels namespace: $($doc.DocumentElement.NamespaceURI)"
    }

    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("rel", $RelationshipsNs)

    # Remove relationships pointing to core.xml or using the core-properties relationship type.
    $existingCoreRelationships = @(
        $doc.SelectNodes(
            "/rel:Relationships/rel:Relationship[@Type='$CoreRelationshipType' or @Target='docProps/core.xml' or @Target='/docProps/core.xml']",
            $ns
        )
    )

    foreach ($node in $existingCoreRelationships) {
        [void]$node.ParentNode.RemoveChild($node)
    }

    # Collect used rIds to avoid collisions.
    $usedIds = @{}

    foreach ($rel in $doc.SelectNodes("/rel:Relationships/rel:Relationship", $ns)) {
        $idValue = $rel.GetAttribute("Id")

        if (-not [string]::IsNullOrWhiteSpace($idValue)) {
            $usedIds[$idValue] = $true
        }
    }

    $newId = "rIdCoreProps"
    $counter = 1

    while ($usedIds.ContainsKey($newId)) {
        $newId = "rIdCoreProps$counter"
        $counter++
    }

    $newRel = $doc.CreateElement("Relationship", $RelationshipsNs)
    $newRel.SetAttribute("Id", $newId)
    $newRel.SetAttribute("Type", $CoreRelationshipType)
    $newRel.SetAttribute("Target", "docProps/core.xml")

    [void]$doc.DocumentElement.AppendChild($newRel)

    return [string](ConvertFrom-XmlDocument -Doc $doc)
}

function Get-ZipEntryCountByName {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName
    )

    return @(
        $Zip.Entries | Where-Object {
            $_.FullName -ieq $EntryName
        }
    ).Count
}

function Test-RepairedPackage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

    try {
        $coreCount = Get-ZipEntryCountByName -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesCount = Get-ZipEntryCountByName -Zip $zip -EntryName $ContentTypesPath
        $relsCount = Get-ZipEntryCountByName -Zip $zip -EntryName $PackageRelsPath

        if ($coreCount -ne 1) {
            throw "Validation failed. Expected exactly one $CorePropertiesPath entry; found $coreCount."
        }

        if ($contentTypesCount -ne 1) {
            throw "Validation failed. Expected exactly one $ContentTypesPath entry; found $contentTypesCount."
        }

        if ($relsCount -ne 1) {
            throw "Validation failed. Expected exactly one $PackageRelsPath entry; found $relsCount."
        }

        $coreXml = Get-ZipEntryText -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesXml = Get-ZipEntryText -Zip $zip -EntryName $ContentTypesPath
        $relsXml = Get-ZipEntryText -Zip $zip -EntryName $PackageRelsPath

        $coreDoc = ConvertTo-XmlDocument -XmlText $coreXml
        $contentTypesDoc = ConvertTo-XmlDocument -XmlText $contentTypesXml
        $relsDoc = ConvertTo-XmlDocument -XmlText $relsXml

        if ($coreDoc.DocumentElement.LocalName -ne "coreProperties") {
            throw "Validation failed. $CorePropertiesPath root is not coreProperties."
        }

        if ($coreDoc.DocumentElement.NamespaceURI -ne "http://schemas.openxmlformats.org/package/2006/metadata/core-properties") {
            throw "Validation failed. $CorePropertiesPath has invalid core-properties namespace."
        }

        $ctNs = [System.Xml.XmlNamespaceManager]::new($contentTypesDoc.NameTable)
        $ctNs.AddNamespace("ct", $ContentTypesNs)

        $coreOverride = $contentTypesDoc.SelectSingleNode(
            "/ct:Types/ct:Override[@PartName='/docProps/core.xml' and @ContentType='$CoreContentType']",
            $ctNs
        )

        if (-not $coreOverride) {
            throw "Validation failed. Missing valid core properties override in $ContentTypesPath."
        }

        $relNs = [System.Xml.XmlNamespaceManager]::new($relsDoc.NameTable)
        $relNs.AddNamespace("rel", $RelationshipsNs)

        $coreRel = $relsDoc.SelectSingleNode(
            "/rel:Relationships/rel:Relationship[@Type='$CoreRelationshipType' and @Target='docProps/core.xml']",
            $relNs
        )

        if (-not $coreRel) {
            throw "Validation failed. Missing valid core properties relationship in $PackageRelsPath."
        }

        return $true
    }
    finally {
        $zip.Dispose()
    }
}

function Repair-XlsxPackage {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $filePath = $File.FullName
    $baseName = $File.BaseName
    $tempPath = Get-SafeTempPath -File $File
    $backupPath = "$filePath.bak"

    if (Test-Path $tempPath) {
        Remove-Item -Path $tempPath -Force
    }

    $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($filePath)

    try {
        $contentTypesXml = Get-ZipEntryText -Zip $sourceZip -EntryName $ContentTypesPath

        if ([string]::IsNullOrWhiteSpace($contentTypesXml)) {
            throw "Missing $ContentTypesPath. This does not appear to be a valid XLSX package."
        }

        $packageRelsXml = Get-ZipEntryText -Zip $sourceZip -EntryName $PackageRelsPath

        $repairedContentTypesXml = Get-RepairedContentTypesXml -OriginalXml $contentTypesXml
        $repairedPackageRelsXml = Get-RepairedPackageRelsXml -OriginalXml $packageRelsXml

        $newCoreXml = New-CorePropertiesXml `
            -Title $baseName `
            -Creator $Creator `
            -Keywords $Keywords `
            -LastModifiedBy $LastModifiedBy

        $targetZip = [System.IO.Compression.ZipFile]::Open(
            $tempPath,
            [System.IO.Compression.ZipArchiveMode]::Create
        )

        try {
            $copiedEntries = @{}

            foreach ($entry in $sourceZip.Entries) {
                $name = $entry.FullName

                # Skip folder/directory entries. File paths imply folders in zip packages.
                if ($name.EndsWith("/")) {
                    continue
                }

                # Remove these so they can be written back exactly once.
                if ($name -ieq $CorePropertiesPath) {
                    continue
                }

                if ($name -ieq $ContentTypesPath) {
                    continue
                }

                if ($name -ieq $PackageRelsPath) {
                    continue
                }

                # Defensive skip for duplicate zip entries.
                $nameKey = $name.ToLowerInvariant()

                if ($copiedEntries.ContainsKey($nameKey)) {
                    continue
                }

                $copiedEntries[$nameKey] = $true

                $newEntry = $targetZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)

                $inputStream = $entry.Open()
                $outputStream = $newEntry.Open()

                try {
                    $inputStream.CopyTo($outputStream)
                }
                finally {
                    $outputStream.Dispose()
                    $inputStream.Dispose()
                }
            }

            Set-ZipEntryText -Zip $targetZip -EntryName $ContentTypesPath -Text $repairedContentTypesXml
            Set-ZipEntryText -Zip $targetZip -EntryName $PackageRelsPath -Text $repairedPackageRelsXml
            Set-ZipEntryText -Zip $targetZip -EntryName $CorePropertiesPath -Text $newCoreXml
        }
        finally {
            $targetZip.Dispose()
        }
    }
    finally {
        $sourceZip.Dispose()
    }

    # Validate rebuilt package before replacing original.
    [void](Test-RepairedPackage -Path $tempPath)

    if ($Backup -and -not (Test-Path $backupPath)) {
        Copy-Item -Path $filePath -Destination $backupPath
    }

    Move-Item -Path $tempPath -Destination $filePath -Force
}

function Get-RepairPreview {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $zip = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)

    try {
        $coreCount = Get-ZipEntryCountByName -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesCount = Get-ZipEntryCountByName -Zip $zip -EntryName $ContentTypesPath
        $relsCount = Get-ZipEntryCountByName -Zip $zip -EntryName $PackageRelsPath

        [PSCustomObject]@{
            File = $File.FullName
            NewTitle = $File.BaseName
            CoreXmlEntries = $coreCount
            ContentTypesEntries = $contentTypesCount
            PackageRelsEntries = $relsCount
        }
    }
    finally {
        $zip.Dispose()
    }
}

Write-Host "Scanning XLSX files in: $ResolvedRootPath" -ForegroundColor Cyan

if (-not (Test-Path $ResolvedRootPath)) {
    throw "Root path does not exist: $ResolvedRootPath"
}

if ($PSBoundParameters.ContainsKey("Days")) {
    $cutoffDate = (Get-Date).ToUniversalTime().Date.AddDays(-$Days)

    $files = @(
        Get-ChildItem -Path $ResolvedRootPath -Recurse -Filter *.xlsx -File |
            Where-Object {
                $_.LastWriteTimeUtc -ge $cutoffDate -and
                $_.Name -notlike "~$*"
            }
    )

    Write-Host "Modified after UTC: $cutoffDate" -ForegroundColor Cyan
}
else {
    $files = @(
        Get-ChildItem -Path $ResolvedRootPath -Recurse -Filter *.xlsx -File |
            Where-Object {
                $_.Name -notlike "~$*"
            }
    )
}

if ($files.Count -eq 0) {
    Write-Host "No XLSX files found." -ForegroundColor Yellow
    return
}

Write-Host "Files found: $($files.Count)" -ForegroundColor Cyan

if ($DoNotApply) {
    Write-Host "Dry run only. No files will be changed." -ForegroundColor Yellow
}
elseif ($Backup) {
    Write-Host "Backups enabled. Existing files will be copied to .bak before replacement." -ForegroundColor Yellow
}
else {
    Write-Host "Backups disabled. Files will be replaced after successful validation." -ForegroundColor Yellow
}

$scanned = 0
$repaired = 0
$wouldRepair = 0
$errored = 0

foreach ($file in $files) {
    $scanned++

    Write-Progress `
        -Activity "Repairing XLSX Core Properties" `
        -Status "$scanned / $($files.Count)" `
        -PercentComplete (($scanned / $files.Count) * 100)

    try {
        if ($DoNotApply) {
            $preview = Get-RepairPreview -File $file

            Write-Host "Would repair: $($preview.File)" -ForegroundColor Yellow
            Write-Host "  New dc:title: $($preview.NewTitle)"
            Write-Host "  Current docProps/core.xml entries: $($preview.CoreXmlEntries)"
            Write-Host "  Current [Content_Types].xml entries: $($preview.ContentTypesEntries)"
            Write-Host "  Current _rels/.rels entries: $($preview.PackageRelsEntries)"

            $wouldRepair++
            continue
        }

        Repair-XlsxPackage -File $file
        $repaired++

        Write-Host "Repaired: $($file.FullName)" -ForegroundColor Green
    }
    catch {
        $errored++
        Write-Warning "Failed: $($file.FullName) :: $($_.Exception.Message)"
    }
}

Write-Progress -Activity "Repairing XLSX Core Properties" -Completed

Write-Host "`n==== Repair Summary ====" -ForegroundColor Cyan
Write-Host "Files Scanned: $scanned"

if ($DoNotApply) {
    Write-Host "Files That Would Be Repaired: $wouldRepair"
}
else {
    Write-Host "Files Repaired: $repaired"
}

Write-Host "Files Errored: $errored"