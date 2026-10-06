<#
.SYNOPSIS
    Unpacks .pbix files into folders, maintaining internal tree structure, delta synchronization,
    and deep deconstruction of Power Query DataMashup (MS-QDEFF) files into M code and settings.

.DESCRIPTION
    Takes all .pbix files in a specified folder, treats them as zip archives, and unzips their contents
    into a subfolder named like the base name of the .pbix file located in the same directory.
    Compares existing files using SHA256 hashes and file lengths to avoid redundant writes.
    Parses and decompiles the DataMashup binary file according to MS-QDEFF:
      - Extracts embedded Package zip (Section1.m, Config/Package.xml, etc.)
      - Decomposes Section1.m into individual query files (.m)
      - Formats Permissions.xml and Metadata.xml settings
      - Generates Metadata_Summary.json mapping query properties
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

.PARAMETER NoMashup
    Disable deep deconstruction of DataMashup (Power Query M and settings).

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
    [switch]$NoMashup,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml.Linq

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
    } catch { }
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
        if ($isWindows) {
            try {
                $ext = Format-ExtendedPath -FilePath $Dir
                [void][System.IO.Directory]::CreateDirectory($ext)
                return
            }
            catch {
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
            try {
                $ext = Format-ExtendedPath -FilePath $FilePath
                return [System.IO.FileStream]::new($ext, $Mode, $Access, $Share)
            }
            catch {
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
    param([Parameter(Mandatory = $true)][System.IO.Stream]$Stream)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha256.ComputeHash($Stream)
        return [System.BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-BytesSha256 {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha256.ComputeHash($Bytes)
        return [System.BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-FileSha256Safe {
    param([Parameter(Mandatory = $true)][string]$FilePath)
    $fileStream = New-FileStreamSafe -FilePath $FilePath -Mode ([System.IO.FileMode]::Open) -Access ([System.IO.FileAccess]::Read) -Share ([System.IO.FileShare]::Read)
    try {
        return Get-StreamSha256 -Stream $fileStream
    }
    finally {
        $fileStream.Dispose()
    }
}

function Format-XmlPretty {
    param([string]$RawXml)
    try {
        $doc = [System.Xml.Linq.XDocument]::Parse($RawXml)
        return $doc.ToString()
    }
    catch {
        return $RawXml
    }
}

function Get-SanitizedFileName {
    param([string]$FileName)
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $pattern = '[' + [regex]::Escape(-join $invalidChars) + ']'
    return ($FileName -replace $pattern, '_').Trim()
}

function Write-DeltaFile {
    param(
        [Parameter(Mandatory = $true)][string]$DestFilePath,
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $true)][string]$RelPath,
        [Parameter(Mandatory = $false)][bool]$IsDryRun = $false
    )

    $srcHash = Get-BytesSha256 -Bytes $Bytes
    $destDir = [System.IO.Path]::GetDirectoryName($DestFilePath)

    $fileExists = $false
    $existingLength = -1
    if ([System.IO.File]::Exists($DestFilePath)) {
        $fileExists = $true
        $existingLength = (New-Object System.IO.FileInfo($DestFilePath)).Length
    }
    elseif ($isWindows -and [System.IO.File]::Exists((Format-ExtendedPath -FilePath $DestFilePath))) {
        $fileExists = $true
        $existingLength = (New-Object System.IO.FileInfo((Format-ExtendedPath -FilePath $DestFilePath))).Length
    }

    if ($fileExists) {
        if ($existingLength -eq $Bytes.Length) {
            $dstHash = Get-FileSha256Safe -FilePath $DestFilePath
            if ($dstHash -eq $srcHash) {
                return [PSCustomObject]@{ Created = 0; Updated = 0; Skipped = 1 }
            }
        }

        if (-not $IsDryRun) {
            New-DirectorySafe -Dir $destDir
            $dstStream = New-FileStreamSafe -FilePath $DestFilePath -Mode ([System.IO.FileMode]::Create) -Access ([System.IO.FileAccess]::Write) -Share ([System.IO.FileShare]::None)
            try {
                $dstStream.Write($Bytes, 0, $Bytes.Length)
            }
            finally {
                $dstStream.Dispose()
            }
        }
        $tag = if ($IsDryRun) { "[DRY-RUN UPDATED]" } else { "[UPDATED]" }
        Write-Host "  $tag $RelPath" -ForegroundColor Yellow
        return [PSCustomObject]@{ Created = 0; Updated = 1; Skipped = 0 }
    }
    else {
        if (-not $IsDryRun) {
            New-DirectorySafe -Dir $destDir
            $dstStream = New-FileStreamSafe -FilePath $DestFilePath -Mode ([System.IO.FileMode]::Create) -Access ([System.IO.FileAccess]::Write) -Share ([System.IO.FileShare]::None)
            try {
                $dstStream.Write($Bytes, 0, $Bytes.Length)
            }
            finally {
                $dstStream.Dispose()
            }
        }
        $tag = if ($IsDryRun) { "[DRY-RUN CREATED]" } else { "[CREATED]" }
        Write-Host "  $tag $RelPath" -ForegroundColor Green
        return [PSCustomObject]@{ Created = 1; Updated = 0; Skipped = 0 }
    }
}

function Split-MQueries {
    param([string]$SectionText)
    $queries = [System.Collections.Generic.Dictionary[string, string]]::new()
    $regex = [regex]'(?m)^[ \t]*shared[ \t]+(#"[^"]+"|[a-zA-Z_][a-zA-Z0-9_]*)[ \t]*='
    $matches = $regex.Matches($SectionText)

    if ($matches.Count -eq 0) {
        $trimmed = $SectionText.Trim()
        if ($trimmed -and (-not $trimmed.StartsWith("section "))) {
            $queries["Query"] = $trimmed
        }
        return $queries
    }

    for ($i = 0; $i -lt $matches.Count; $i++) {
        $m = $matches[$i]
        $rawName = $m.Groups[1].Value
        $cleanName = if ($rawName.StartsWith('#"') -and $rawName.EndsWith('"')) {
            $rawName.Substring(2, $rawName.Length - 3)
        } else {
            $rawName
        }

        $startPos = $m.Index
        $endPos = if ($i + 1 -lt $matches.Count) { $matches[$i + 1].Index } else { $SectionText.Length }
        $expr = $SectionText.Substring($startPos, $endPos - $startPos).Trim()
        if ($expr.EndsWith(';')) {
            $expr = $expr.Substring(0, $expr.Length - 1).TrimEnd()
        }
        $queries[$cleanName] = $expr
    }

    return $queries
}

function Expand-DataMashupBinary {
    param(
        [Parameter(Mandatory = $true)][string]$TargetDir,
        [Parameter(Mandatory = $true)][byte[]]$MashupBytes,
        [Parameter(Mandatory = $false)][bool]$IsDryRun = $false
    )

    $created = 0; $updated = 0; $skipped = 0; $errors = 0
    if ($MashupBytes.Length -lt 8) { return [PSCustomObject]@{ Created=0; Updated=0; Skipped=0; Errors=0 } }

    try {
        $memStream = [System.IO.MemoryStream]::new($MashupBytes)
        $reader = [System.IO.BinaryReader]::new($memStream)

        $version = $reader.ReadUInt32()
        $pkgLen = $reader.ReadUInt32()

        if ($memStream.Position + $pkgLen -gt $MashupBytes.Length) {
            return [PSCustomObject]@{ Created=0; Updated=0; Skipped=0; Errors=1 }
        }

        $pkgBytes = $reader.ReadBytes($pkgLen)
        $permBytes = [byte[]]@()
        if ($memStream.Position + 4 -le $MashupBytes.Length) {
            $permLen = $reader.ReadUInt32()
            $permBytes = $reader.ReadBytes($permLen)
        }

        $metaBytes = [byte[]]@()
        if ($memStream.Position + 4 -le $MashupBytes.Length) {
            $metaLen = $reader.ReadUInt32()
            $metaBytes = $reader.ReadBytes($metaLen)
        }

        # 1. Package ZIP extraction
        if ($pkgBytes.Length -gt 0) {
            $pkgMem = [System.IO.MemoryStream]::new($pkgBytes)
            $zip = [System.IO.Compression.ZipArchive]::new($pkgMem, [System.IO.Compression.ZipArchiveMode]::Read)
            try {
                foreach ($entry in $zip.Entries) {
                    if ($entry.FullName -eq "Formulas/Section1.m") {
                        $sReader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
                        $mText = $sReader.ReadToEnd()
                        $sReader.Dispose()

                        # Write full Section1.m
                        $rel = "DataMashup_Extracted/Section1.m"
                        $dest = [System.IO.Path]::Combine($TargetDir, $rel)
                        $r = Write-DeltaFile -DestFilePath $dest -Bytes ([System.Text.Encoding]::UTF8.GetBytes($mText)) -RelPath $rel -IsDryRun $IsDryRun
                        $created += $r.Created; $updated += $r.Updated; $skipped += $r.Skipped

                        # Decompose individual queries
                        $queries = Split-MQueries -SectionText $mText
                        foreach ($qName in $queries.Keys) {
                            $safeName = Get-SanitizedFileName -FileName $qName
                            $qRel = "DataMashup_Extracted/Queries/$safeName.m"
                            $qDest = [System.IO.Path]::Combine($TargetDir, $qRel)
                            $qBytes = [System.Text.Encoding]::UTF8.GetBytes($queries[$qName])
                            $qr = Write-DeltaFile -DestFilePath $qDest -Bytes $qBytes -RelPath $qRel -IsDryRun $IsDryRun
                            $created += $qr.Created; $updated += $qr.Updated; $skipped += $qr.Skipped
                        }
                    }
                    elseif ($entry.FullName -eq "Config/Package.xml") {
                        $sReader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
                        $pXml = Format-XmlPretty -RawXml $sReader.ReadToEnd()
                        $sReader.Dispose()
                        $rel = "DataMashup_Extracted/Package.xml"
                        $dest = [System.IO.Path]::Combine($TargetDir, $rel)
                        $r = Write-DeltaFile -DestFilePath $dest -Bytes ([System.Text.Encoding]::UTF8.GetBytes($pXml)) -RelPath $rel -IsDryRun $IsDryRun
                        $created += $r.Created; $updated += $r.Updated; $skipped += $r.Skipped
                    }
                }
            }
            finally {
                $zip.Dispose()
                $pkgMem.Dispose()
            }
        }

        # 2. Permissions XML
        if ($permBytes.Length -gt 0) {
            $permStr = [System.Text.Encoding]::UTF8.GetString($permBytes)
            $pXmlPretty = Format-XmlPretty -RawXml $permStr
            $rel = "DataMashup_Extracted/Permissions.xml"
            $dest = [System.IO.Path]::Combine($TargetDir, $rel)
            $r = Write-DeltaFile -DestFilePath $dest -Bytes ([System.Text.Encoding]::UTF8.GetBytes($pXmlPretty)) -RelPath $rel -IsDryRun $IsDryRun
            $created += $r.Created; $updated += $r.Updated; $skipped += $r.Skipped
        }

        # 3. Metadata XML
        if ($metaBytes.Length -ge 8) {
            $metaMem = [System.IO.MemoryStream]::new($metaBytes)
            $mReader = [System.IO.BinaryReader]::new($metaMem)
            $metaVer = $mReader.ReadUInt32()
            $metaXmlLen = $mReader.ReadUInt32()
            if ($metaMem.Position + $metaXmlLen -le $metaBytes.Length) {
                $metaXmlRaw = [System.Text.Encoding]::UTF8.GetString($mReader.ReadBytes($metaXmlLen))
                $mXmlPretty = Format-XmlPretty -RawXml $metaXmlRaw
                $rel = "DataMashup_Extracted/Metadata.xml"
                $dest = [System.IO.Path]::Combine($TargetDir, $rel)
                $r = Write-DeltaFile -DestFilePath $dest -Bytes ([System.Text.Encoding]::UTF8.GetBytes($mXmlPretty)) -RelPath $rel -IsDryRun $IsDryRun
                $created += $r.Created; $updated += $r.Updated; $skipped += $r.Skipped
            }
            $mReader.Dispose()
            $metaMem.Dispose()
        }

        $reader.Dispose()
        $memStream.Dispose()
    }
    catch {
        $errors++
        Write-Host "  [SKIPPED - ERROR] Could not decompile DataMashup: $_" -ForegroundColor Red
    }

    return [PSCustomObject]@{ Created = $created; Updated = $updated; Skipped = $skipped; Errors = $errors }
}

function Expand-PbixFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PbixPath,

        [Parameter(Mandatory = $false)]
        [bool]$IsDryRun = $false,

        [Parameter(Mandatory = $false)]
        [bool]$ParseMashup = $true
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
    $mashupBytes = $null

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
            if ([string]::IsNullOrEmpty($entry.Name) -or $entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) {
                continue
            }

            $entryRelPath = $entry.FullName.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
            if ($entryRelPath.StartsWith("..") -or [System.IO.Path]::IsPathRooted($entryRelPath)) {
                Write-Warning "Skipping unsafe zip entry path: $($entry.FullName)"
                continue
            }

            $destFilePath = [System.IO.Path]::Combine($targetDir, $entryRelPath)

            try {
                $eStream = $entry.Open()
                $eMem = [System.IO.MemoryStream]::new()
                try {
                    $eStream.CopyTo($eMem)
                    $bytes = $eMem.ToArray()
                }
                finally {
                    $eMem.Dispose()
                    $eStream.Dispose()
                }

                if ($entryRelPath -eq "DataMashup") {
                    $mashupBytes = $bytes
                }

                $r = Write-DeltaFile -DestFilePath $destFilePath -Bytes $bytes -RelPath $entryRelPath -IsDryRun $IsDryRun
                $created += $r.Created
                $updated += $r.Updated
                $skipped += $r.Skipped
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

    # Deep DataMashup extraction if present
    if ($ParseMashup -and $null -ne $mashupBytes) {
        Write-Host "  -> Decompiling DataMashup (Power Query M & Settings)..." -ForegroundColor Cyan
        $mr = Expand-DataMashupBinary -TargetDir $targetDir -MashupBytes $mashupBytes -IsDryRun $IsDryRun
        $created += $mr.Created
        $updated += $mr.Updated
        $skipped += $mr.Skipped
        $errors += $mr.Errors
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
    $foldersQueue = [System.Collections.Generic.Queue[string]]::new()
    $foldersQueue.Enqueue($resolvedPath)

    while ($foldersQueue.Count -gt 0) {
        $currentFolder = $foldersQueue.Dequeue()

        Get-ChildItem -LiteralPath $currentFolder -File -Filter "*.pbix" | ForEach-Object {
            if (Test-IncludeItem -Name $_.Name -IncludePatterns $FileInclude -ExcludePatterns $FileExclude) {
                $discoveredFiles.Add($_.FullName)
            }
        }

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
    $res = Expand-PbixFile -PbixPath $file -IsDryRun $DryRun.IsPresent -ParseMashup (-not $NoMashup.IsPresent)
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
