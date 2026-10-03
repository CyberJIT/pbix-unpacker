<#
.SYNOPSIS
    Unpacks .pbix files into folders, maintaining internal tree structure and only replacing modified or new files.

.DESCRIPTION
    Takes all .pbix files in a specified folder, treats them as zip archives, and unzips their contents
    into a subfolder named like the base name of the .pbix file located in the same directory.
    Compares existing files using SHA256 hashes and file lengths to avoid redundant writes.
    Supports recursive discovery, file name filtering (include/exclude), folder filtering (include/exclude),
    long path extended prefixes (\\?\), Win32 short 8.3 path fallbacks, and graceful error reporting.

.PARAMETER Path
    Target directory containing .pbix files.

.PARAMETER Recurse
    Recursively search for .pbix files in subdirectories.

.PARAMETER FileInclude
    Array of string fragments or wildcards to include matching pbix file names.

.PARAMETER FileExclude
    Array of string fragments or wildcards to exclude matching pbix file names.

.PARAMETER FolderInclude
    Array of string fragments or wildcards to include matching folder names during recursive search.

.PARAMETER FolderExclude
    Array of string fragments or wildcards to exclude matching folder names during recursive search.

.PARAMETER DryRun
    Simulate execution without modifying the filesystem.

.EXAMPLE
    .\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse
    .\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -FileInclude "Sales*", "*Finance*" -FileExclude "*Draft*"
    .\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse -FolderInclude "2024", "2025" -FolderExclude "Archive"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [switch]$Recurse,

    [Parameter(Mandatory = $false)]
    [string[]]$FileInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FileExclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FolderInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FolderExclude,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Define Kernel32 P/Invoke for Win32 GetShortPathName on Windows systems
$isWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
if ($isWindows) {
    $csharpMethods = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class Kernel32PathHelper {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern uint GetShortPathNameW(string lpszLongPath, [Out] StringBuilder lpszShortPath, uint cchBuffer);

    public static string GetShortPath(string path) {
        try {
            StringBuilder sb = new StringBuilder(1024);
            uint res = GetShortPathNameW(path, sb, (uint)sb.Capacity);
            if (res > 0 && res < sb.Capacity) {
                return sb.ToString();
            }
        } catch { }
        return null;
    }
}
"@
    try {
        if (-not ([System.Management.Automation.PSTypeName]'Kernel32PathHelper').Type) {
            Add-Type -TypeDefinition $csharpMethods
        }
    } catch {
        # Ignore if already added or restricted
    }
}

function Format-ExtendedPath {
    param([string]$FilePath)
    if (-not $isWindows) { return $FilePath }
    if ($FilePath.StartsWith("\\?\")) { return $FilePath }
    if ($FilePath.StartsWith("\\")) {
        return "\\?\UNC\" + $FilePath.Substring(2)
    }
    return "\\?\" + $FilePath
}

function New-DirectorySafe {
    param([string]$Dir)
    if ([System.IO.Directory]::Exists($Dir)) { return }
    try {
        [void][System.IO.Directory]::CreateDirectory($Dir)
    }
    catch {
        # Fallback to extended path prefix on Windows
        if ($isWindows) {
            try {
                $ext = Format-ExtendedPath -FilePath $Dir
                [void][System.IO.Directory]::CreateDirectory($ext)
                return
            }
            catch {
                # Fallback to short path if available
                $parent = [System.IO.Path]::GetDirectoryName($Dir)
                if ($parent -and ([System.Management.Automation.PSTypeName]'Kernel32PathHelper').Type) {
                    $shortParent = [Kernel32PathHelper]::GetShortPath($parent)
                    if ($shortParent) {
                        $shortDir = [System.IO.Path]::Combine($shortParent, [System.IO.Path]::GetFileName($Dir))
                        [void][System.IO.Directory]::CreateDirectory($shortDir)
                        return
                    }
                }
            }
        }
        throw
    }
}

function New-FileStreamSafe {
    param(
        [string]$FilePath,
        [System.IO.FileMode]$Mode,
        [System.IO.FileAccess]$Access,
        [System.IO.FileShare]$Share
    )

    try {
        return [System.IO.FileStream]::new($FilePath, $Mode, $Access, $Share)
    }
    catch {
        if ($isWindows) {
            # Try \\?\ prefix
            try {
                $ext = Format-ExtendedPath -FilePath $FilePath
                return [System.IO.FileStream]::new($ext, $Mode, $Access, $Share)
            }
            catch {
                # Try short path of parent directory
                $dir = [System.IO.Path]::GetDirectoryName($FilePath)
                if ($dir -and ([System.Management.Automation.PSTypeName]'Kernel32PathHelper').Type) {
                    $shortDir = [Kernel32PathHelper]::GetShortPath($dir)
                    if ($shortDir) {
                        $altPath = [System.IO.Path]::Combine($shortDir, [System.IO.Path]::GetFileName($FilePath))
                        return [System.IO.FileStream]::new($altPath, $Mode, $Access, $Share)
                    }
                }
            }
        }
        throw
    }
}

function Test-PatternMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string[]]$Patterns
    )

    if (-not $Patterns -or $Patterns.Count -eq 0) {
        return $false
    }

    foreach ($pat in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($pat)) { continue }
        # Check wildcard match or substring match
        if ($Name -like "*$pat*" -or $Name -like $pat) {
            return $true
        }
    }
    return $false
}

function Test-IncludeItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string[]]$IncludePatterns,

        [Parameter(Mandatory = $false)]
        [string[]]$ExcludePatterns
    )

    if ($ExcludePatterns -and (Test-PatternMatch -Name $Name -Patterns $ExcludePatterns)) {
        return $false
    }

    if ($IncludePatterns -and $IncludePatterns.Count -gt 0 -and (-not (Test-PatternMatch -Name $Name -Patterns $IncludePatterns))) {
        return $false
    }

    return $true
}

function Get-StreamSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Stream]$Stream
    )
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha256.ComputeHash($Stream)
        return [System.BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-FileSha256Safe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    $fileStream = New-FileStreamSafe -FilePath $FilePath -Mode ([System.IO.FileMode]::Open) -Access ([System.IO.FileAccess]::Read) -Share ([System.IO.FileShare]::Read)
    try {
        return Get-StreamSha256 -Stream $fileStream
    }
    finally {
        $fileStream.Dispose()
    }
}

function Expand-PbixFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PbixPath,

        [Parameter(Mandatory = $false)]
        [bool]$IsDryRun = $false
    )

    $parentDir = [System.IO.Path]::GetDirectoryName($PbixPath)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($PbixPath)
    $targetDir = [System.IO.Path]::Combine($parentDir, $baseName)

    Write-Host "`nProcessing: $PbixPath" -ForegroundColor Cyan
    Write-Host "Target directory: $targetDir"

    $created = 0
    $updated = 0
    $skipped = 0
    $errors = 0

    $zipArchive = $null
    try {
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($PbixPath)
    }
    catch {
        Write-Host "  [ERROR] Failed to open '$PbixPath' as a zip archive: $_" -ForegroundColor Red
        return [PSCustomObject]@{ Created = 0; Updated = 0; Skipped = 0; Errors = 1 }
    }

    try {
        foreach ($entry in $zipArchive.Entries) {
            # Skip folders
            if ([string]::IsNullOrEmpty($entry.Name) -or $entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) {
                continue
            }

            # Normalise relative paths & prevent Zip Slip
            $entryRelPath = $entry.FullName.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
            if ($entryRelPath.StartsWith("..") -or [System.IO.Path]::IsPathRooted($entryRelPath)) {
                Write-Warning "Skipping unsafe zip entry path: $($entry.FullName)"
                continue
            }

            $destFilePath = [System.IO.Path]::Combine($targetDir, $entryRelPath)
            $destFileDir = [System.IO.Path]::GetDirectoryName($destFilePath)

            try {
                # Calculate hash of incoming entry
                $entryStream = $entry.Open()
                $srcHash = ""
                try {
                    $srcHash = Get-StreamSha256 -Stream $entryStream
                }
                finally {
                    $entryStream.Dispose()
                }

                $fileExists = $false
                $existingLength = -1
                if ([System.IO.File]::Exists($destFilePath)) {
                    $fileExists = $true
                    $existingLength = (New-Object System.IO.FileInfo($destFilePath)).Length
                }
                elseif ($isWindows -and [System.IO.File]::Exists((Format-ExtendedPath -FilePath $destFilePath))) {
                    $fileExists = $true
                    $existingLength = (New-Object System.IO.FileInfo((Format-ExtendedPath -FilePath $destFilePath))).Length
                }

                if ($fileExists) {
                    if ($existingLength -eq $entry.Length) {
                        $dstHash = Get-FileSha256Safe -FilePath $destFilePath
                        if ($dstHash -eq $srcHash) {
                            $skipped++
                            continue
                        }
                    }

                    # File content differs
                    if (-not $IsDryRun) {
                        New-DirectorySafe -Dir $destFileDir
                        $entryStream = $entry.Open()
                        $dstFileStream = New-FileStreamSafe -FilePath $destFilePath -Mode ([System.IO.FileMode]::Create) -Access ([System.IO.FileAccess]::Write) -Share ([System.IO.FileShare]::None)
                        try {
                            $entryStream.CopyTo($dstFileStream)
                        }
                        finally {
                            $dstFileStream.Dispose()
                            $entryStream.Dispose()
                        }
                    }
                    $updated++
                    $tag = if ($IsDryRun) { "[DRY-RUN UPDATED]" } else { "[UPDATED]" }
                    Write-Host "  $tag $entryRelPath" -ForegroundColor Yellow
                }
                else {
                    # New file
                    if (-not $IsDryRun) {
                        New-DirectorySafe -Dir $destFileDir
                        $entryStream = $entry.Open()
                        $dstFileStream = New-FileStreamSafe -FilePath $destFilePath -Mode ([System.IO.FileMode]::Create) -Access ([System.IO.FileAccess]::Write) -Share ([System.IO.FileShare]::None)
                        try {
                            $entryStream.CopyTo($dstFileStream)
                        }
                        finally {
                            $dstFileStream.Dispose()
                            $entryStream.Dispose()
                        }
                    }
                    $created++
                    $tag = if ($IsDryRun) { "[DRY-RUN CREATED]" } else { "[CREATED]" }
                    Write-Host "  $tag $entryRelPath" -ForegroundColor Green
                }
            }
            catch {
                $errors++
                Write-Host "  [SKIPPED - ERROR] Could not unpack '$entryRelPath': $_" -ForegroundColor Red
            }
        }
    }
    finally {
        if ($null -ne $zipArchive) {
            $zipArchive.Dispose()
        }
    }

    $summaryLabel = if ($IsDryRun) { "Dry-run summary" } else { "Unpacked" }
    $errorMsg = if ($errors -gt 0) { ", $errors failed/skipped" } else { "" }
    Write-Host "  -> $summaryLabel : $created created, $updated updated, $skipped unchanged$errorMsg." -ForegroundColor Gray

    return [PSCustomObject]@{
        Created = $created
        Updated = $updated
        Skipped = $skipped
        Errors  = $errors
    }
}

# Resolve and validate root path
$resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (-not (Test-Path -LiteralPath $resolvedPath -PathType Container)) {
    throw "Specified path does not exist or is not a directory: $resolvedPath"
}

# Collect PBIX files
$discoveredFiles = [System.Collections.Generic.List[string]]::new()

if (-not $Recurse) {
    Get-ChildItem -LiteralPath $resolvedPath -File -Filter "*.pbix" | ForEach-Object {
        if (Test-IncludeItem -Name $_.Name -IncludePatterns $FileInclude -ExcludePatterns $FileExclude) {
            $discoveredFiles.Add($_.FullName)
        }
    }
}
else {
    # Recursive search with folder prune support
    $foldersQueue = [System.Collections.Generic.Queue[string]]::new()
    $foldersQueue.Enqueue($resolvedPath)

    while ($foldersQueue.Count -gt 0) {
        $currentFolder = $foldersQueue.Dequeue()

        # Get .pbix files in this folder
        Get-ChildItem -LiteralPath $currentFolder -File -Filter "*.pbix" | ForEach-Object {
            if (Test-IncludeItem -Name $_.Name -IncludePatterns $FileInclude -ExcludePatterns $FileExclude) {
                $discoveredFiles.Add($_.FullName)
            }
        }

        # Traverse subfolders respecting folder inclusion/exclusion
        Get-ChildItem -LiteralPath $currentFolder -Directory | ForEach-Object {
            if (Test-IncludeItem -Name $_.Name -IncludePatterns $FolderInclude -ExcludePatterns $FolderExclude) {
                $foldersQueue.Enqueue($_.FullName)
            }
        }
    }
}

if ($discoveredFiles.Count -eq 0) {
    Write-Host "No matching .pbix files found in '$resolvedPath'." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($discoveredFiles.Count) matching .pbix file(s)." -ForegroundColor Green

$totalCreated = 0
$totalUpdated = 0
$totalSkipped = 0
$totalErrors  = 0

foreach ($file in $discoveredFiles) {
    $res = Expand-PbixFile -PbixPath $file -IsDryRun $DryRun.IsPresent
    $totalCreated += $res.Created
    $totalUpdated += $res.Updated
    $totalSkipped += $res.Skipped
    $totalErrors  += $res.Errors
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  PBIX files processed : $($discoveredFiles.Count)"
Write-Host "  Files created        : $totalCreated"
Write-Host "  Files updated        : $totalUpdated"
Write-Host "  Files unchanged      : $totalSkipped"
if ($totalErrors -gt 0) {
    Write-Host "  Files/items errored  : $totalErrors (skipped gracefully)" -ForegroundColor Red
}
Write-Host "==================================================" -ForegroundColor Cyan
