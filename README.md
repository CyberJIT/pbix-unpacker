# PBIX Unpacker (Python & PowerShell)

A lightweight utility to extract `.pbix` (Power BI Desktop) files as ZIP archives into sibling folders named after the report's base name. It preserves internal folder structures, verifies file hashes (SHA-256) to update **only** modified or new files, decompiles the embedded **DataMashup** binary stream into plain-text Power Query (M) code and settings, and supports flexible filtering and recursion.

---

## Features

- **In-Place Sibling Unpacking**: Extracts `./ReportName.pbix` into `./ReportName/`.
- **Preserved Directory Hierarchy**: Maintains nested internal paths (e.g. `Report/Layout`, `StaticResources/...`).
- **DataMashup Deconstruction (Power Query M & Settings)**:
  - Decompiles the Microsoft MS-QDEFF binary package container.
  - Extracts `Formulas/Section1.m` and deconstructs it into individual `.m` files per query under `DataMashup_Extracted/Queries/<QueryName>.m`.
  - Pretty-prints configuration files (`Package.xml`, `Permissions.xml`, `Metadata.xml`).
  - Generates `Metadata_Summary.json` mapping query descriptions, query group IDs, and load-to-model status flags.
- **Content-Aware Delta Updates**: Checks file size and SHA-256 checksums to avoid touching unchanged files.
- **Recursive Directory Traversal**: Optional recursive search across nested folders.
- **Granular Filtering**:
  - Filter in / out `.pbix` files by substring or wildcard.
  - Filter in / out subdirectories during recursion by substring or wildcard.
- **Long Path Support & Short-Name Fallback**: Automatic support for Windows extended-length paths (`\\?\`) and 8.3 short-name workarounds (`GetShortPathNameW`) to bypass `MAX_PATH` limitations.
- **Graceful Error Recovery**: If an individual file fails due to an insurmountable path or permission error, it logs the failure and continues unpacking remaining files instead of aborting the process.
- **Dry-Run Mode**: Preview creations, updates, and unchanged counts without writing to disk.

---

## 1. Python Tool (`pbix_unpacker.py` & `mashup_parser.py`)

Requires Python 3.7+ (pure standard library, zero external dependencies).

### Usage

```bash
# Basic unpack of a single folder (with automatic DataMashup deconstruction)
python3 pbix_unpacker.py -p "/path/to/pbix_folder"

# Recursive scan under subdirectories
python3 pbix_unpacker.py -p "/path/to/pbix_folder" -r

# Filter pbix files (e.g. include 'Finance' or 'Sales', exclude 'Draft')
python3 pbix_unpacker.py -p "/path/to/pbix_folder" -r \
  --file-include "*Finance*" "*Sales*" \
  --file-exclude "*Draft*"

# Filter folders to traverse (e.g. only folders containing '2025', skip 'Archive')
python3 pbix_unpacker.py -p "/path/to/pbix_folder" -r \
  --folder-include "2025" \
  --folder-exclude "Archive" "Old"

# Skip deep DataMashup deconstruction (raw archive extraction only)
python3 pbix_unpacker.py -p "/path/to/pbix_folder" --no-mashup

# Dry run simulation
python3 pbix_unpacker.py -p "/path/to/pbix_folder" -r --dry-run
```

### CLI Parameters

| Flag | Full Option | Description |
| :--- | :--- | :--- |
| `-p` | `--path` | **(Required)** Path to target directory containing `.pbix` files. |
| `-r` | `--recursive` | Recursively search subdirectories for `.pbix` files. |
| | `--file-include` | One or more patterns/substrings to include `.pbix` files. |
| | `--file-exclude` | One or more patterns/substrings to exclude `.pbix` files. |
| | `--folder-include` | One or more patterns/substrings to allow folders during recursion. |
| | `--folder-exclude` | One or more patterns/substrings to prune folders during recursion. |
| | `--no-mashup` | Skip deep parsing and extraction of `DataMashup`. |
| `-n` | `--dry-run` | Preview actions without creating or replacing files. |

---

## 2. PowerShell Tool (`Unpack-Pbix.ps1`)

Compatible with Windows PowerShell 5.1+ and PowerShell 7+ (Core / macOS / Linux). Uses native .NET compression (`System.IO.Compression.ZipArchive`) and cryptographic streams for high performance.

### Usage

```powershell
# Basic unpack (with automatic DataMashup deconstruction)
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports"

# Recursive scan
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse

# Filter pbix files
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse `
  -FileInclude "Finance*", "Sales*" `
  -FileExclude "*Draft*"

# Filter folders during traversal
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse `
  -FolderInclude "2025" `
  -FolderExclude "Archive", "Old"

# Skip deep DataMashup deconstruction
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -NoMashup

# Dry run simulation
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse -DryRun
```

### Script Parameters

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-Path` | `String` | **(Required)** Path to directory containing `.pbix` files. |
| `-Recurse` | `Switch` | Search subdirectories recursively. |
| `-FileInclude` | `String[]` | List of substrings or wildcards to include `.pbix` files. |
| `-FileExclude` | `String[]` | List of substrings or wildcards to skip `.pbix` files. |
| `-FolderInclude` | `String[]` | List of substrings or wildcards of folders to traverse. |
| `-FolderExclude` | `String[]` | List of substrings or wildcards of folders to skip. |
| `-NoMashup` | `Switch` | Skip deep parsing and extraction of `DataMashup`. |
| `-DryRun` | `Switch` | Preview operations without touching disk files. |

---

## Output Structure

When unpacking a report `Sales.pbix`, the resulting directory will be organized as follows:

```text
Sales/
├── [Content_Types].xml
├── Connections
├── DataModel
├── DiagramLayout
├── Settings
├── Version
├── Report/
│   └── Layout
└── DataMashup_Extracted/
    ├── Section1.m                  <- Full Power Query formula document
    ├── Package.xml                 <- Version and culture settings
    ├── Permissions.xml             <- Data source privacy/firewall settings
    ├── Metadata.xml                <- Raw query metadata entries
    ├── Metadata_Summary.json       <- Structured JSON of query descriptions & flags
    └── Queries/                    <- Deconstructed individual query scripts
        ├── Customers.m
        ├── Orders.m
        └── DateTable.m
```

---

## Technical Feasibility & Environment Report

| Capability | Python (3.7+ stdlib) | PowerShell (.NET Framework / Core) |
| :--- | :--- | :--- |
| **MS-QDEFF Binary Decompilation** | Fully supported (`struct`, `io.BytesIO`) | Fully supported (`BinaryReader`, `MemoryStream`) |
| **Package ZIP Extraction** | Fully supported (`zipfile`) | Fully supported (`ZipArchive`) |
| **Section1.m Query Splitting** | Fully supported (`re`) | Fully supported (`[regex]`) |
| **XML Settings Formatting** | Fully supported (`xml.dom.minidom`) | Fully supported (`XDocument`) |
| **External Dependencies Required** | **None** (zero third-party packages) | **None** (standard .NET BCL assemblies) |

---

## Testing

Run unit tests via Python's standard `unittest`:

```bash
python3 -m unittest test_pbix_unpacker.py
```
