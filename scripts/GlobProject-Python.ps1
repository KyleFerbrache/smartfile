<#
.SYNOPSIS
Concatenates Python source files in a project into a single glob file.

.DESCRIPTION
- Scans a configured directory recursively to identify Python candidate files
- Excludes `.venv` and `__pycache__` directories by default
- Accepts additional excluded directory names through the cmdlet
- Writes all eligible `.py` file contents into a single output file
- Uses deterministic sorting based on relative path
- Emits progress during scan and concatenation
- Writes a concise execution summary at completion

.PARAMETER RootPath
Root Directory to Scan.

.PARAMETER AdditionalExcludeDirectories
Additional directory names to exclude from processing.
Default exclusions always include `.venv` and `__pycache__`.

.EXAMPLE
.\GlobProject-Python.ps1 -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/99 - Code Repositories/pdguard/pdguard"

.EXAMPLE
.\GlobProject-Python.ps1 `
    -RootPath "Freeman Mathis and Gary/FMG - Financial Systems - Documents/99 - Code Repositories/pdguard/pdguard" `
    -AdditionalExcludeDirectories ".git", "build", "dist"
#>



[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RootPath,

    [string[]]$AdditionalExcludeDirectories = @()
)


# Use compact progress when running in PowerShell 7+
# This usually renders as a bottom/status-line style progress display.
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $PSStyle.Progress.View = 'Minimal'
}


function Get-RelativePath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)

    if (-not $baseFullPath.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $baseFullPath += [System.IO.Path]::DirectorySeparatorChar
    }

    $baseUri = New-Object System.Uri($baseFullPath)
    $targetUri = New-Object System.Uri($targetFullPath)

    $relativeUri = $baseUri.MakeRelativeUri($targetUri)
    $relativePath = [System.Uri]::UnescapeDataString($relativeUri.ToString())

    return $relativePath -replace '/', [System.IO.Path]::DirectorySeparatorChar
}



$RootPath = "$HOME/$RootPath"
$ResolvedRootPath = (Resolve-Path -Path $RootPath).Path

if (-not (Test-Path -Path $ResolvedRootPath -PathType Container)) {
    throw "RootPath does not exist or is not a directory: $ResolvedRootPath"
}


# Output file is always written to the input root as `glob.py`
$OutputPath = Join-Path -Path $ResolvedRootPath -ChildPath "glob.py"

# Default exclusions plus any user-provided additions
$ExcludedDirectories = @(".venv", "__pycache__") + $AdditionalExcludeDirectories

Write-Host "Scanning *.py Files in $ResolvedRootPath" -ForegroundColor Cyan
Write-Host "Output File: $OutputPath"

if ($ExcludedDirectories.Count -gt 0) {
    Write-Host "Excluded Directories: $($ExcludedDirectories -join ', ')"
}


# Phase 1 - File Discovery
$allFiles = Get-ChildItem -Path $ResolvedRootPath -Recurse -Filter *.py -File

$scannedCount = $allFiles.Count
$discoveredFiles = @()
$scanIndex = 0

foreach ($file in $allFiles) {
    $scanIndex++

    Write-Progress `
        -Id 1 `
        -Activity "Scanning Python Files" `
        -Status "$scanIndex / $scannedCount" `
        -PercentComplete (($scanIndex / [Math]::Max($scannedCount, 1)) * 100)

    $relativePath = Get-RelativePath -BasePath $ResolvedRootPath -TargetPath $file.FullName

    # Exclude the output file itself if it already exists in the root
    if ($file.FullName -eq $OutputPath) {
        continue
    }

    # Exclude files if any path segment matches an excluded directory name
    $pathSegments = $relativePath -split '[\\/]'
    $isExcluded = $false

    foreach ($excludedDirectory in $ExcludedDirectories) {
        if ($pathSegments -contains $excludedDirectory) {
            $isExcluded = $true
            break
        }
    }

    if (-not $isExcluded) {
        $discoveredFiles += [PSCustomObject]@{
            FullPath     = $file.FullName
            RelativePath = $relativePath
        }
    }
}

Write-Progress -Id 1 -Activity "Scanning Python Files" -Completed

$files = $discoveredFiles | Sort-Object RelativePath
$totalFiles = $files.Count

if ($totalFiles -eq 0) {
    Write-Host "No Files Found Matching Provided Criteria" -ForegroundColor Yellow
    return
}

Write-Host "Found $totalFiles eligible files. Concatenating..." -ForegroundColor Cyan


# Processing Counters
$linesWritten = 0
$errored = 0


# Phase 2 - Concatenation
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$writer = New-Object System.IO.StreamWriter($OutputPath, $false, $utf8NoBom)

try {
    $index = 0

    foreach ($file in $files) {
        $index++

        Write-Progress `
            -Id 2 `
            -Activity "Concatenating Python Files" `
            -Status "$index / $totalFiles" `
            -PercentComplete (($index / $totalFiles) * 100)

        try {
            $content = Get-Content -Path $file.FullPath -Raw -Encoding UTF8
            $trimmedContent = $content.TrimEnd()

            if ([string]::IsNullOrEmpty($trimmedContent)) {
                $fileLineCount = 0
            }
            else {
                $fileLineCount = ($trimmedContent -split "`r`n|`n|`r").Count
            }

            $writer.WriteLine("# ======== $($file.RelativePath) ======== #")
            $writer.WriteLine($trimmedContent)
            $writer.WriteLine("")

            $linesWritten += $fileLineCount
        }
        catch {
            $errored++
        }
    }
}
finally {
    $writer.Close()
}

Write-Progress -Id 2 -Activity "Concatenating Python Files" -Completed


# Write Change Summary Logs
Write-Host "`n==== Change Summary ====" -ForegroundColor Cyan
Write-Host "Files Scanned: $scannedCount"
Write-Host "Files Written: $totalFiles"
Write-Host "Lines Written: $linesWritten"
Write-Host "Files Errored: $errored"
Write-Host "Output File: $OutputPath"

