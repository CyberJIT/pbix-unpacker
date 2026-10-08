# PBIX & PBIT Unpacker (Python & PowerShell)

A lightweight utility to extract `.pbix` and `.pbit` (Power BI Desktop & Template) files as ZIP archives into sibling folders named after the report's base name. It preserves internal folder structures, verifies file hashes (SHA-256) to update **only** modified or new files, decompiles the embedded **DataMashup** binary stream into plain-text Power Query (M) code and settings, deconstructs **DataModelSchema** (TMSL JSON) into tables, partitions, and DAX measures, and supports flexible filtering and recursion.

---

## Features

- **In-Place Sibling Unpacking**: Extracts `./ReportName.pbix` into `./ReportName/`.
- **Preserved Directory Hierarchy**: Maintains nested internal paths (e.g. `Report/Layout`, `StaticResources/...`).
- **DataMashup Deconstruction (Power Query M & Settings)**:
  - Decompiles the Microsoft MS-QDEFF binary package container.
  - Extracts `Formulas/Section1.m` and deconstructs it into individual `.m` files per query under `DataMashup_Extracted/Queries/<QueryName>.m`.
  - Pretty-prints configuration files (`Package.xml`, `Permissions.xml`, `Metadata.xml`).
  - Generates `Metadata_Summary.json` mapping query descriptions, query group IDs, and load-to-model status flags.
- **DataModelSchema Deconstruction (TMSL JSON & DAX Measures)**:
  - Supports `.pbit` templates and modern packages containing `DataModelSchema`.
  - Pretty-prints `DataModelSchema_Pretty.json`.
  - Decomposes all **DAX measures** into standalone `.dax` files categorized by table under `DataModelSchema_Extracted/Measures/<Table>/<Measure>.dax`.
  - Extracts table metadata, model partitions (embedded M queries), and relationships JSON.
- **Report Visuals & Layout Deconstruction (`pbi-tools` style)**:
  - Formats `Report/Layout` into `Report_Extracted/Layout_Pretty.json`.
  - Deconstructs pages into `Report_Extracted/Pages/<PageName>/page.json`.
  - Extracts individual visual containers into readable JSONs under `Report_Extracted/Pages/<PageName>/Visuals/<Index>_<Type>_<ID>.json`, unwrapping nested escaped JSON strings (`config`, `filters`, `dataTransforms`).
  - Pretty-prints `DiagramLayout_Pretty.json` (diagram visual coordinates).
  - Formats `LinguisticSchema_Pretty.xml` (Q&A natural language definitions).
  - Formats `Settings_Pretty.json` and `Metadata_Pretty.json`.
- **DataModel VertiPaq ABF Bridge & Active Port Discovery**:
  - Automatically identifies the proprietary `DataModel` VertiPaq Analysis Services Backup File (ABF) container and parses the compression header (e.g. `XPress9`).
  - Generates standalone `DataModel_Bridge/DataModel.abf` ready for direct restoration into any local or remote SSAS Tabular developer instance.
  - Generates ready-to-execute `DataModel_Bridge/Restore_Database.xmla` targeting SSAS.
  - Creates `DataModel_Bridge/DataModel_Info.json` outlining model metrics and recommended extraction workflows.
  - Scans for running Power BI Desktop SSAS (`msmdsrv.exe`) instances via `msmdsrv.port.txt` in `AnalysisServicesWorkspace` and prints connection strings for DAX Studio, Tabular Editor, or ADOMD.NET.
- **Content-Aware Delta Updates**: Checks file size and SHA-256 checksums to avoid touching unchanged files.
- **Recursive Directory Traversal**: Optional recursive search across nested folders for `.pbix` and `.pbit` files.
- **Granular Filtering**:
  - Filter in / out files by substring or wildcard.
  - Filter in / out subdirectories during recursion by substring or wildcard.
- **Long Path Support & Short-Name Fallback**: Automatic support for Windows extended-length paths (`\\?\`) and 8.3 short-name workarounds (`GetShortPathNameW`) to bypass `MAX_PATH` limitations.
- **Graceful Error Recovery**: If an individual file fails due to an insurmountable path or permission error, it logs the failure and continues unpacking remaining files instead of aborting the process.
- **Dry-Run Mode**: Preview creations, updates, and unchanged counts without writing to disk.

---

## 1. Python Tool (`pbix_unpacker.py`, `mashup_parser.py`, `schema_parser.py`, `layout_parser.py`, `datamodel_bridge.py`)

Requires Python 3.7+ (pure standard library, zero external dependencies).

### Usage

```bash
# Basic unpack of a folder (automatic DataMashup, DataModelSchema, Report visuals & DataModel bridge)
python3 pbix_unpacker.py -p "/path/to/reports"

# Detect active Power BI Desktop local SSAS ports and print connection strings
python3 pbix_unpacker.py --find-active-ports

# Recursive scan under subdirectories
python3 pbix_unpacker.py -p "/path/to/reports" -r

# Filter reports (e.g. include 'Finance' or 'Sales', exclude 'Draft')
python3 pbix_unpacker.py -p "/path/to/reports" -r \
  --file-include "*Finance*" "*Sales*" \
  --file-exclude "*Draft*"

# Filter folders to traverse (e.g. only folders containing '2025', skip 'Archive')
python3 pbix_unpacker.py -p "/path/to/reports" -r \
  --folder-include "2025" \
  --folder-exclude "Archive" "Old"

# Skip deep deconstruction if raw archive unpacking is desired
python3 pbix_unpacker.py -p "/path/to/reports" --no-mashup --no-schema --no-report --no-datamodel

# Dry run simulation
python3 pbix_unpacker.py -p "/path/to/reports" -r --dry-run
```

### CLI Parameters

| Flag | Full Option | Description |
| :--- | :--- | :--- |
| `-p` | `--path` | Path to target directory containing `.pbix`/`.pbit` files. |
| `-r` | `--recursive` | Recursively search subdirectories for `.pbix`/`.pbit` files. |
| | `--file-include` | One or more patterns/substrings to include files. |
| | `--file-exclude` | One or more patterns/substrings to exclude files. |
| | `--folder-include` | One or more patterns/substrings to allow folders during recursion. |
| | `--folder-exclude` | One or more patterns/substrings to prune folders during recursion. |
| | `--no-mashup` | Skip deep parsing and extraction of `DataMashup`. |
| | `--no-schema` | Skip deep parsing and extraction of `DataModelSchema`. |
| | `--no-report` | Skip deep parsing and deconstruction of `Report` layout, visuals, and diagrams. |
| | `--no-datamodel`| Skip DataModel VertiPaq ABF bridge export and XMLA script generation. |
| | `--find-active-ports` | Scan local machine for active Power BI Desktop SSAS ports and print connection strings. |
| `-n` | `--dry-run` | Preview actions without creating or replacing files. |

---

## 2. PowerShell Tool (`Unpack-Pbix.ps1`)

Compatible with Windows PowerShell 5.1+ and PowerShell 7+ (Core / macOS / Linux). Uses native .NET compression (`System.IO.Compression.ZipArchive`) and cryptographic streams for high performance.

### Usage

```powershell
# Basic unpack (with automatic DataMashup, DataModelSchema, Report visuals & DataModel bridge)
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports"

# Detect active Power BI Desktop local SSAS ports and print connection strings
.\Unpack-Pbix.ps1 -FindActivePorts

# Recursive scan
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse

# Filter reports
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse `
  -FileInclude "Finance*", "Sales*" `
  -FileExclude "*Draft*"

# Filter folders during traversal
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse `
  -FolderInclude "2025" `
  -FolderExclude "Archive", "Old"

# Skip deep deconstruction
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -NoMashup -NoSchema -NoReport -NoDataModel

# Dry run simulation
.\Unpack-Pbix.ps1 -Path "C:\PowerBI\Reports" -Recurse -DryRun
```

### Script Parameters

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-Path` | `String` | Path to directory containing `.pbix`/`.pbit` files. |
| `-Recurse` | `Switch` | Search subdirectories recursively. |
| `-FileInclude` | `String[]` | List of substrings or wildcards to include files. |
| `-FileExclude` | `String[]` | List of substrings or wildcards to skip files. |
| `-FolderInclude` | `String[]` | List of substrings or wildcards of folders to traverse. |
| `-FolderExclude` | `String[]` | List of substrings or wildcards of folders to skip. |
| `-NoMashup` | `Switch` | Skip deep parsing and extraction of `DataMashup`. |
| `-NoSchema` | `Switch` | Skip deep parsing and extraction of `DataModelSchema`. |
| `-NoReport` | `Switch` | Skip deep parsing and extraction of `Report` layout, visuals, and diagrams. |
| `-NoDataModel` | `Switch` | Skip DataModel VertiPaq ABF bridge export and XMLA script generation. |
| `-FindActivePorts` | `Switch` | Scan local machine for active Power BI Desktop SSAS ports and print connection strings. |
| `-DryRun` | `Switch` | Preview operations without touching disk files. |

---

## Output Structure

When unpacking reports, the resulting folders will contain:

```text
Report/
├── [Content_Types].xml
├── Connections
├── DataModel                   <- Raw VertiPaq / Analysis Services backup (in .pbix)
├── DataModelSchema             <- Raw TMSL JSON (in .pbit)
├── DiagramLayout
├── Settings
├── Version
├── Report/
│   └── Layout
├── DataMashup_Extracted/       <- Extracted Power Query M and settings
│   ├── Section1.m
│   ├── Package.xml
│   ├── Permissions.xml
│   ├── Metadata.xml
│   ├── Metadata_Summary.json
│   └── Queries/
│       ├── Customers.m
│       └── Orders.m
├── DataModelSchema_Extracted/  <- Extracted TMSL model definitions
│   ├── DataModelSchema_Pretty.json
│   ├── Relationships.json
│   ├── Tables/
│   │   ├── Sales.json
│   │   └── Date.json
│   ├── Partitions/
│   │   └── Sales_SalesPartition.m
│   └── Measures/
│       └── Sales/
│           ├── Total Sales.dax
│           └── Sales YTD.dax
├── Report_Extracted/           <- Deconstructed visual layout and diagrams
│   ├── Layout_Pretty.json
│   ├── DiagramLayout_Pretty.json
│   ├── LinguisticSchema_Pretty.xml
│   ├── Settings_Pretty.json
│   ├── Metadata_Pretty.json
│   └── Pages/
│       └── 01_Executive Summary/
│           ├── page.json
│           └── Visuals/
│               ├── 01_barChart_visual1.json
│               └── 02_card_visual2.json
└── DataModel_Bridge/           <- VertiPaq ABF bridge & hydration scripts
    ├── DataModel.abf           <- Renamed SSAS backup container
    ├── Restore_Database.xmla   <- Direct XMLA restore script for SSAS Tabular developer instance
    └── DataModel_Info.json     <- Model size, compression (XPress9), and live connection guides
```

---

## Technical Feasibility & Environment Report

| Capability | Python (3.7+ stdlib) | PowerShell (.NET Framework / Core) |
| :--- | :--- | :--- |
| **MS-QDEFF Binary Decompilation** | Fully supported (`struct`, `io.BytesIO`) | Fully supported (`BinaryReader`, `MemoryStream`) |
| **DataModelSchema TMSL JSON** | Fully supported (`json`, UTF-16LE auto-decode) | Fully supported (`ConvertFrom-Json`) |
| **DAX Measures Deconstruction** | Fully supported | Fully supported |
| **Report Layout & Visual Containers** | Fully supported (`layout_parser.py`) | Fully supported (`ConvertFrom-Json`, nested string unwrap) |
| **Diagram Layout & Linguistic Schema** | Fully supported | Fully supported |
| **DataModel VertiPaq ABF Bridge & XMLA** | Fully supported (`datamodel_bridge.py`) | Fully supported (`Expand-DataModelBridge`) |
| **Active SSAS Port Discovery** | Fully supported (`find_active_powerbi_ssas_instances`) | Fully supported (`Get-PowerBiActiveSsasInstances`) |
| **Third-Party Dependencies** | **None** (zero pip installs) | **None** (standard .NET assemblies) |

---

## Testing

Run unit tests via Python's standard `unittest`:

```bash
python3 -m unittest test_pbix_unpacker.py
```
