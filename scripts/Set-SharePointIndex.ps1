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

.PARAMETER DoNotApply
Dry-run mode. Reports what would be updated without modifying the files.

.EXAMPLE
.\Set-SharePointIndex.ps1 -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources" -Days 1

.EXAMPLE
.\Set-SharePointIndex.ps1 - RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/01 - Data Sources"
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RootPath,

    [ValidateRange(0, 3650)]
    [int]$Days = 1,

    [switch]$DoNotApply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Use compact progress when running in PowerShell 7+
# This usually renders as a bottom/status-line style progress display.
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $PSStyle.Progress.View = 'Minimal'
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$BaseDir = $HOME
$ResolvedRootPath = Join-Path $BaseDir $RootPath
$LastModifiedBy = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "SmartFile" } else { $env:USERNAME }
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)


# XML Namespace/Schema Declarations
$DcTermsNS = "http://purl.org/dc/terms/"
$DcmiTypeNS = "http://purl.org/dc/dcmitype/"
$DcElementsNS = "http://purl.org/dc/elements/1.1/"
$XMLSchemaNS = "http://www.w3.org/2001/XMLSchema-instance"
$CoreContentType = "application/vnd.openxmlformats-package.core-properties+xml"
$ContentTypesNS = "http://schemas.openxmlformats.org/package/2006/content-types"
$RelationshipsNS = "http://schemas.openxmlformats.org/package/2006/relationships"
$CorePropertiesNS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
$CoreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

# Excel ZIP Schema Declarations
$PackageRelsPath = "_rels/.rels"
$ContentTypesPath = "[Content_Types].xml"
$CorePropertiesPath = "docProps/core.xml"


function Get-SafeTempPath {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )
    $guid = [System.Guid]::NewGuid().ToString("N")
    return Join-Path $File.DirectoryName ".$($File.Name).sharepoint-index-$guid.tmp"
}

function Get-ZipEntryText {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName
    )
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }

    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new($stream, $true)
        try { return [string]$reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Set-ZipEntryText {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName,

        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [DateTimeOffset]$LastWriteTime
    )
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $entry.LastWriteTime = $LastWriteTime

    $stream = $entry.Open()
    try {
        $writer = [System.IO.StreamWriter]::new($stream, $Utf8NoBom)
        try { $writer.Write($Text) }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
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

function ConvertTo-XmlDocument {
    param (
        [Parameter(Mandatory = $true)]
        [string]$XmlText
    )
    if ([string]::IsNullOrWhiteSpace($XmlText)) { throw "XML Text is Empty." }

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
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::None

    $memoryStream = [System.IO.MemoryStream]::new()

    try {
        $writer = [System.Xml.XmlWriter]::Create($memoryStream, $settings)
        try { $Doc.Save($writer) }
        finally { $writer.Dispose() }

        return [string]$Utf8NoBom.GetString($memoryStream.ToArray())
    }
    finally { $memoryStream.Dispose() }
}

function New-CorePropertiesXml {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [DateTime]$CreatedTimeStampUTC,

        [Parameter(Mandatory = $true)]
        [DateTime]$ModifiedTimeStampUTC
    )

    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
    $escapedModifier = [System.Security.SecurityElement]::Escape($LastModifiedBy)
    $escapedCreator = [System.Security.SecurityElement]::Escape("SmartFile CLI Utility")
    $escapedKeywords = [System.Security.SecurityElement]::Escape("SmartFile.Set-SharePointIndex")
    $CreatedTimeStampText = $CreatedTimeStampUTC.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $ModifiedTimeStampText = $ModifiedTimeStampUTC.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties
xmlns:cp="$CorePropertiesNS" xmlns:dc="$DcElementsNS" xmlns:dcterms="$DcTermsNS"
xmlns:dcmitype="$DcmiTypeNS" xmlns:xsi="$XMLSchemaNS">
<dc:title>$escapedTitle</dc:title><dc:creator>$escapedCreator</dc:creator>
<cp:keywords>$escapedKeywords</cp:keywords><cp:lastModifiedBy>$escapedModifier</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">$CreatedTimeStampText</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">$ModifiedTimeStampText</dcterms:modified>
</cp:coreProperties>
"@
    return [string]$xml
}

function Get-CoreTitle {
    param (
        [AllowNull()]
        [string]$CoreXml
    )
    if ([string]::IsNullOrWhiteSpace($CoreXml)) { return $null }

    try {
        $doc = ConvertTo-XmlDocument -XmlText $CoreXml
        if (-not $doc.DocumentElement) { return $null } 
        if ($doc.DocumentElement.LocalName -ne "coreProperties") { return $null }
        if ($doc.DocumentElement.NamespaceURI -ne $CorePropertiesNS) { return $null }

        $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
        $ns.AddNamespace("cp", $CorePropertiesNS)
        $ns.AddNamespace("dc", $DcElementsNS)

        $titleNode = $doc.SelectSingleNode("/cp:coreProperties/dc:title", $ns)
        if ($titleNode) { return [string]$titleNode.InnerText } else { return $null }
    }
    catch { return $null }
}

function Get-CoreCreated {
    param (
        [AllowNull()]
        [string]$CoreXml
    )
    if ([string]::IsNullOrWhiteSpace($CoreXml)) { return $null }

    try {
        $doc = ConvertTo-XmlDocument -XmlText $CoreXml
        if (-not $doc.DocumentElement) { return $null } 
        if ($doc.DocumentElement.LocalName -ne "coreProperties") { return $null }
        if ($doc.DocumentElement.NamespaceURI -ne $CorePropertiesNS) { return $null }

        $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
        $ns.AddNamespace("cp", $CorePropertiesNS)
        $ns.AddNamespace("dcterms", $DcTermsNS)

        [DateTime]$createdValue = [DateTime]::MinValue

        $createdNode = $doc.SelectSingleNode("/cp:coreProperties/dcterms:created", $ns)
        if ($createdNode) {
            if (
               [DateTime]::TryParse(
                    $createdNode.InnerText,
                    [Globalization.CultureInfo]::InvariantCulture,
                    [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal,
                    [ref]$createdValue
                )) { return $createdValue.ToUniversalTime() }
            } else { return $null }
    } catch { return $null }
}

function Get-RepairedContentTypesXml {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OriginalXml
    )
    $doc = ConvertTo-XmlDocument -XmlText $OriginalXml
    if (-not $doc.DocumentElement) { throw "Invalid $ContentTypesPath - Missing ElementTree" }
    if ($doc.DocumentElement.LocalName -ne "Types") { throw "Invalid $ContentTypesPath - Missing Root Element: 'types'"}
    if ($doc.DocumentElement.NamespaceURI -ne $ContentTypesNS) { throw "Invalid $ContentTypesPath Namespace: '$($doc.DocumentElement.NamespaceURI)'" }

    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("ct", $ContentTypesNS)

    $existing = $doc.SelectNodes("/ct:Types/ct:Override[@PartName='/docProps/core.xml']", $ns)
    foreach ($node in $existing) { [void]$node.ParentNode.RemoveChild($node) }

    $override = $doc.CreateElement("Override", $ContentTypesNS)
    $override.SetAttribute("PartName", "/docProps/core.xml")
    $override.SetAttribute("ContentType", $CoreContentType)

    [void]$doc.DocumentElement.AppendChild($override)

    return ConvertFrom-XmlDocument -Doc $doc
}

function New-DefaultPackageRelsXml {
    $doc = [System.Xml.XmlDocument]::new()
    $doc.PreserveWhitespace = $false
    $declaration = $doc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
    $root = $doc.CreateElement("Relationships", $RelationshipsNS)
    [void]$doc.AppendChild($declaration)
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
    } else { $doc = ConvertTo-XmlDocument -XmlText $OriginalXml }
    
    if (-not $doc.DocumentElement) { throw "Invalid $PackageRelsPath - Missing ElementTree" }
    if ($doc.DocumentElement.LocalName -ne "Relationships") { throw "Invalid $PackageRelsPath - Missing Root Element: 'Relationships'"}
    if ($doc.DocumentElement.NamespaceURI -ne $RelationshipsNS) { throw "Invalid $PackageRelsPath Namespace: '$($doc.DocumentElement.NamespaceURI)'" }

    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("rel", $RelationshipsNS)

    $existing = $doc.SelectNodes("/rel:Relationships/rel:Relationship[@Type='$CoreRelationshipType' or @Target='$CorePropertiesPath' or @Target='/$CorePropertiesPath']", $ns)
    foreach ($node in $existing) { [void]$node.ParentNode.RemoveChild($node) }

    $usedIds = @{}
    foreach ($rel in $doc.SelectNodes("/rel:Relationships/rel:Relationship", $ns)) {
        $idValue = $rel.GetAttribute("Id")
        if (-not [string]::IsNullOrWhiteSpace($idValue)) { $usedIds[$idValue] = $true }
    }

    $newId = "rIdCoreProps"
    $counter = 1
    while ($usedIds.ContainsKey($newId)) {
        $newId = "rIdCoreProps$counter"
        $counter++
    }

    $newRel = $doc.CreateElement("Relationship", $RelationshipsNS)
    $newRel.SetAttribute("Id", $newId)
    $newRel.SetAttribute("Type", $CoreRelationshipType)
    $newRel.SetAttribute("Target", $CorePropertiesPath)

    [void]$doc.DocumentElement.AppendChild($newRel)

    return ConvertFrom-XmlDocument -Doc $doc
}

function Test-CorePropertiesXmlValid {
    param (
        [AllowNull()]
        [string]$CoreXml,
        
        [Parameter(Mandatory = $true)]
        [string]$ExpectedTitle
    )
    if ([string]::IsNullOrWhiteSpace($CoreXml)) { return $false }
    
    try {
        $doc = ConvertTo-XmlDocument -XmlText $CoreXml
        if (-not $doc.DocumentElement) { return $false }
        if ($doc.DocumentElement.LocalName -ne "coreProperties") { return $false }
        if ($doc.DocumentElement.NamespaceURI -ne $CorePropertiesNS) { return $false }

        $titleNode = Get-CoreTitle -CoreXml $CoreXml
        if (-not $titleNode -or $titleNode -ne $ExpectedTitle) { return $false } else { return $true }
    }
    catch { return $false }
}

function Test-ContentTypesHasValidCoreOverride {
    param (
        [AllowNull()]
        [string]$ContentTypesXml
    )

    if ([string]::IsNullOrWhiteSpace($ContentTypesXml)) { return $false } 

    try {
        $doc = ConvertTo-XmlDocument -XmlText $ContentTypesXml
        $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
        $ns.AddNamespace("ct", $ContentTypesNs)
        $matches = $doc.SelectNodes(
            "/ct:Types/ct:Override[@PartName='/docProps/core.xml' and @ContentType='$CoreContentType']",
            $ns
        )

        return ($matches.Count -eq 1)
    } catch { return $false }
}

function Test-PackageRelsHasValidCoreRelationship {
    param (
        [AllowNull()]
        [string]$PackageRelsXml
    )
    if ([string]::IsNullOrWhiteSpace($PackageRelsXml)) { return $false }

    try {
        $doc = ConvertTo-XmlDocument -XmlText $PackageRelsXml
        $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
        $ns.AddNamespace("rel", $RelationshipsNs)
        $matches = $doc.SelectNodes(
            "/rel:Relationships/rel:Relationship[@Type='$CoreRelationshipType' and @Target='docProps/core.xml']",
            $ns
        )

        return ($matches.Count -eq 1)
    } catch { return $false }
}


function Get-XlsxUpdateAssessment {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )
    $expectedTitle = $File.BaseName
    $zip = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)

    try {
        $coreCount = Get-ZipEntryCountByName -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesCount = Get-ZipEntryCountByName -Zip $zip -EntryName $ContentTypesPath
        $relsCount = Get-ZipEntryCountByName -Zip $zip -EntryName $PackageRelsPath

        $coreXml = Get-ZipEntryText -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesXml = Get-ZipEntryText -Zip $zip -EntryName $ContentTypesPath
        $packageRelsXml = Get-ZipEntryText -Zip $zip -EntryName $PackageRelsPath

        $reasons = New-Object System.Collections.Generic.List[string]

        if ($coreCount -ne 1) { $reasons.Add("Expected exactly one $CorePropertiesPath entry; found $coreCount.") }
        if ($contentTypesCount -ne 1) { $reasons.Add("Expected exactly one $ContentTypesPath entry; found $contentTypesCount.") }
        if ($relsCount -ne 1) { $reasons.Add("Expected exactly one $PackageRelsPath entry; found $relsCount.") }

        if (-not (Test-CorePropertiesXmlValid -CoreXml $coreXml -ExpectedTitle $expectedTitle)) {
            $currentTitle = Get-CoreTitle -CoreXml $coreXml

            if ([string]::IsNullOrWhiteSpace($currentTitle)) { $reasons.Add("Core properties XML is missing, malformed, or missing required metadata.") }
            elseif ($currentTitle -ne $expectedTitle) { $reasons.Add("Title mismatch. Current title '$currentTitle' should be '$expectedTitle'.") }
            else { $reasons.Add("Core properties XML title is correct, but metadata structure is incomplete or invalid.") }
        }

        if (-not (Test-ContentTypesHasValidCoreOverride -ContentTypesXml $contentTypesXml)) { $reasons.Add("Missing or duplicate core-properties override in $ContentTypesPath.") }
        if (-not (Test-PackageRelsHasValidCoreRelationship -PackageRelsXml $packageRelsXml)) { $reasons.Add("Missing or duplicate core-properties relationship in $PackageRelsPath.") }

        return [PSCustomObject]@{
            File = $File.FullName
            FileInfo = $File
            ExpectedTitle = $expectedTitle
            NeedsUpdate = ($reasons.Count -gt 0)
            Reasons = @($reasons)
        }
    } finally { $zip.Dispose() }
}


function Test-RebuiltPackage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedTitle
    )
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

    try {
        $coreCount = Get-ZipEntryCountByName -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesCount = Get-ZipEntryCountByName -Zip $zip -EntryName $ContentTypesPath
        $relsCount = Get-ZipEntryCountByName -Zip $zip -EntryName $PackageRelsPath

        if ($coreCount -ne 1) { throw "Validation failed. Expected exactly one $CorePropertiesPath entry; found $coreCount." }
        if ($contentTypesCount -ne 1) { throw "Validation failed. Expected exactly one $ContentTypesPath entry; found $contentTypesCount." }
        if ($relsCount -ne 1) { throw "Validation failed. Expected exactly one $PackageRelsPath entry; found $relsCount." }

        $coreXml = Get-ZipEntryText -Zip $zip -EntryName $CorePropertiesPath
        $contentTypesXml = Get-ZipEntryText -Zip $zip -EntryName $ContentTypesPath
        $packageRelsXml = Get-ZipEntryText -Zip $zip -EntryName $PackageRelsPath

        if (-not (Test-CorePropertiesXmlValid -CoreXml $coreXml -ExpectedTitle $ExpectedTitle)) { throw "Validation failed. $CorePropertiesPath is invalid or title does not match." }
        if (-not (Test-ContentTypesHasValidCoreOverride -ContentTypesXml $contentTypesXml)) { throw "Validation failed. $ContentTypesPath does not contain exactly one valid core-properties override." }
        if (-not (Test-PackageRelsHasValidCoreRelationship -PackageRelsXml $packageRelsXml)) { throw "Validation failed. $PackageRelsPath does not contain exactly one valid core-properties relationship." }

        return $true
    } finally { $zip.Dispose() }
}


function Update-XlsxPackage {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $filePath = $File.FullName
    $baseName = $File.BaseName
    $tempPath = Get-SafeTempPath -File $File
    $timestampUtc = (Get-Date).ToUniversalTime()
    $zipTimestamp = [DateTimeOffset]::new($timestampUtc)

    if (Test-Path $tempPath) { Remove-Item -Path $tempPath -Force }
    
    try {
        $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($filePath)
        $coreXml = Get-ZipEntryText -Zip $sourceZip -EntryName $CorePropertiesPath
        $CreatedTimestampUTC = Get-CoreCreated -CoreXml $coreXml
        if ($null -eq $CreatedTimestampUTC) { $CreatedTimestampUTC = $timestampUtc}

        $contentTypesXml = Get-ZipEntryText -Zip $sourceZip -EntryName $ContentTypesPath

        if ([string]::IsNullOrWhiteSpace($contentTypesXml)) { throw "Missing $ContentTypesPath. This does not appear to be a valid XLSX package." }

        $packageRelsXml = Get-ZipEntryText -Zip $sourceZip -EntryName $PackageRelsPath
        $repairedContentTypesXml = Get-RepairedContentTypesXml -OriginalXml $contentTypesXml
        $repairedPackageRelsXml = Get-RepairedPackageRelsXml -OriginalXml $packageRelsXml
        $newCoreXml = New-CorePropertiesXml -Title $baseName -CreatedTimeStampUTC $CreatedTimestampUTC -ModifiedTimeStampUTC $timestampUtc

        $targetZip = [System.IO.Compression.ZipFile]::Open(
            $tempPath,
            [System.IO.Compression.ZipArchiveMode]::Create
        )

        try {
            $copiedEntries = @{}

            foreach ($entry in $sourceZip.Entries) {
                $name = $entry.FullName

                if ($name.EndsWith("/")) { continue }
                if ($name -ieq $CorePropertiesPath) { continue }
                if ($name -ieq $ContentTypesPath) { continue }
                if ($name -ieq $PackageRelsPath) { continue }

                $nameKey = $name.ToLowerInvariant()
                if ($copiedEntries.ContainsKey($nameKey)) { continue }

                $copiedEntries[$nameKey] = $true
                $newEntry = $targetZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $newEntry.LastWriteTime = $zipTimestamp

                $inputStream = $entry.Open()
                $outputStream = $newEntry.Open()

                try { $inputStream.CopyTo($outputStream) }
                finally {
                    $outputStream.Dispose()
                    $inputStream.Dispose()
                }
            }

            Set-ZipEntryText `
                -Zip $targetZip `
                -EntryName $ContentTypesPath `
                -Text $repairedContentTypesXml `
                -LastWriteTime $zipTimestamp

            Set-ZipEntryText `
                -Zip $targetZip `
                -EntryName $PackageRelsPath `
                -Text $repairedPackageRelsXml `
                -LastWriteTime $zipTimestamp

            Set-ZipEntryText `
                -Zip $targetZip `
                -EntryName $CorePropertiesPath `
                -Text $newCoreXml `
                -LastWriteTime $zipTimestamp
        } finally { $targetZip.Dispose() }
    } finally { $sourceZip.Dispose() }

    Test-RebuiltPackage -Path $tempPath -ExpectedTitle $baseName | Out-Null
    Move-Item -Path $tempPath -Destination $filePath -Force
    [System.IO.File]::SetLastWriteTimeUtc($filePath, $timestampUtc)
}

# Main Workflow

Write-Host "Scanning XLSX files in: $ResolvedRootPath" -ForegroundColor Cyan

if (-not (Test-Path $ResolvedRootPath)) {
    throw "Root path does not exist: $ResolvedRootPath"
}

$cutoffDate = (Get-Date).ToUniversalTime().Date.AddDays(-$Days)

Write-Host "Modified after UTC: $cutoffDate" -ForegroundColor Cyan

$files = @(
    Get-ChildItem -Path $ResolvedRootPath -Recurse -Filter *.xlsx -File |
        Where-Object {
            $_.LastWriteTimeUtc -ge $cutoffDate -and
            $_.Name -notlike "~$*" -and
            $_.Name -notlike "*.tmp" -and
            $_.Name -notlike "*.bak"
        }
)

if ($files.Count -eq 0) {
    Write-Host "No files found matching provided criteria." -ForegroundColor Yellow
    return
}

Write-Host "Eligible files found: $($files.Count)" -ForegroundColor Cyan
Write-Host "Assessment pass starting. No files are modified during this phase." -ForegroundColor Cyan

$assessments = New-Object System.Collections.Generic.List[object]
$assessmentErrors = 0
$assessmentIndex = 0

foreach ($file in $files) {
    $assessmentIndex++

    Write-Progress `
        -Activity "Assessing Eligible XLSX Files" `
        -Status "$assessmentIndex / $($files.Count)" `
        -PercentComplete (($assessmentIndex / $files.Count) * 100)

    try {
        $assessment = Get-XlsxUpdateAssessment -File $file
        $assessments.Add($assessment)

        if ($assessment.NeedsUpdate) {
            Write-Host "Needs update: $($assessment.File)" -ForegroundColor Yellow

            foreach ($reason in $assessment.Reasons) {
                Write-Host "  - $reason" -ForegroundColor DarkYellow
            }
        }
    }
    catch {
        $assessmentErrors++
        Write-Warning "Assessment failed: $($file.FullName) :: $($_.Exception.Message)"
    }
}

Write-Progress -Activity "Assessing Eligible XLSX Files" -Completed

$changePlan = @(
    $assessments | Where-Object {
        $_.NeedsUpdate
    }
)

$skipped = $assessments.Count - $changePlan.Count

Write-Host "`n==== Assessment Summary ====" -ForegroundColor Cyan
Write-Host "Files Assessed: $($assessments.Count)"
Write-Host "Files Needing Update: $($changePlan.Count)"
Write-Host "Files Already Valid: $skipped"
Write-Host "Assessment Errors: $assessmentErrors"

if ($DoNotApply) {
    Write-Host "`nDoNotApply was provided. Exiting after assessment pass." -ForegroundColor Yellow
    return
}

if ($changePlan.Count -eq 0) {
    Write-Host "`nNo updates required." -ForegroundColor Green
    return
}

Write-Host "`nUpdate pass starting. Only files in the change plan will be modified." -ForegroundColor Cyan

$updated = 0
$updateErrors = 0
$updateIndex = 0

foreach ($item in $changePlan) {
    $updateIndex++

    Write-Progress `
        -Activity "Updating XLSX Metadata" `
        -Status "$updateIndex / $($changePlan.Count)" `
        -PercentComplete (($updateIndex / $changePlan.Count) * 100)

    try {
        $file = Get-Item -LiteralPath $item.File -ErrorAction Stop

        Update-XlsxPackage -File $file

        $updated++
        Write-Host "Updated: $($file.FullName)" -ForegroundColor Green
    }
    catch {
        $updateErrors++
        Write-Warning "Update failed: $($item.File) :: $($_.Exception.Message)"
    }
}

Write-Progress -Activity "Updating XLSX Metadata" -Completed

Write-Host "`n==== SharePoint Index Metadata Summary ====" -ForegroundColor Cyan
Write-Host "Eligible Files: $($files.Count)"
Write-Host "Assessed Files: $($assessments.Count)"
Write-Host "ReadXML Errors: $assessmentErrors"
Write-Host "Validated XMLs: $($skipped)"
Write-Host "Proposed Edits: $($changePlan.Count)"
Write-Host "Edits Complete: $updated"
Write-Host "Editing Errors: $updateErrors"