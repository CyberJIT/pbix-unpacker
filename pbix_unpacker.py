#!/usr/bin/env python3
"""
pbix_unpacker.py

A tool to extract .pbix and .pbit files (treated as zip archives) into subfolders named
after the base name of each file in the same directory, maintaining internal folder
hierarchies, and only replacing files if they are new or have different contents.

Supports recursive scanning, file name filters (include/exclude), folder name filters (include/exclude),
Windows extended-length paths (\\\\?\\) and short 8.3 path fallbacks, graceful error recovery,
deep parsing of DataMashup (MS-QDEFF) into Power Query (M) scripts, and decompilation of
DataModelSchema (TMSL JSON) into measures, tables, partitions, and relationships.
"""

import argparse
import ctypes
import fnmatch
import hashlib
import json
import os
import re
import sys
import zipfile
from typing import List, Optional, Tuple

from mashup_parser import DataMashupParser
from schema_parser import DataModelSchemaParser, decode_schema_bytes


def normalize_long_path(path: str) -> str:
    """
    On Windows, prepend the \\\\?\\ prefix to absolute paths longer than 240 chars
    (or if already prefixed) to bypass MAX_PATH (260 chars) limitations.
    On non-Windows platforms, returns the path unmodified.
    """
    if os.name != "nt":
        return path
    abs_path = os.path.abspath(path)
    if abs_path.startswith("\\\\?\\"):
        return abs_path
    if abs_path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abs_path[2:]
    return "\\\\?\\" + abs_path


def get_short_path_name(long_name: str) -> Optional[str]:
    """Retrieve Windows 8.3 short path equivalent using GetShortPathNameW API."""
    if os.name != "nt":
        return None
    try:
        buffer_size = 512
        buffer = ctypes.create_unicode_buffer(buffer_size)
        res = ctypes.windll.kernel32.GetShortPathNameW(long_name, buffer, buffer_size)
        if 0 < res < buffer_size:
            return buffer.value
    except Exception:
        pass
    return None


def ensure_directory(dir_path: str) -> None:
    """Create directory ensuring support for long paths and short path workarounds."""
    try:
        os.makedirs(dir_path, exist_ok=True)
    except (OSError, FileNotFoundError) as err:
        if os.name == "nt":
            long_p = normalize_long_path(dir_path)
            try:
                os.makedirs(long_p, exist_ok=True)
                return
            except Exception:
                short_p = get_short_path_name(dir_path)
                if short_p and short_p != dir_path:
                    os.makedirs(short_p, exist_ok=True)
                    return
        raise err


def calculate_stream_sha256(stream, chunk_size: int = 65536) -> str:
    """Calculate SHA256 checksum from a file-like stream."""
    hasher = hashlib.sha256()
    while True:
        chunk = stream.read(chunk_size)
        if not chunk:
            break
        hasher.update(chunk)
    return hasher.hexdigest()


def calculate_file_sha256(filepath: str, chunk_size: int = 65536) -> str:
    """Calculate SHA256 checksum from an on-disk file, with long path fallback."""
    try:
        with open(filepath, "rb") as f:
            return calculate_stream_sha256(f, chunk_size=chunk_size)
    except OSError:
        if os.name == "nt":
            extended_p = normalize_long_path(filepath)
            with open(extended_p, "rb") as f:
                return calculate_stream_sha256(f, chunk_size=chunk_size)
        raise


def get_path_size(filepath: str) -> int:
    """Get file size handling long path fallback."""
    try:
        return os.path.getsize(filepath)
    except OSError:
        if os.name == "nt":
            return os.path.getsize(normalize_long_path(filepath))
        raise


def file_exists_safe(filepath: str) -> bool:
    """Check existence handling long path fallback."""
    if os.path.exists(filepath):
        return True
    if os.name == "nt":
        return os.path.exists(normalize_long_path(filepath))
    return False


def matches_patterns(name: str, patterns: Optional[List[str]]) -> bool:
    """Check if a name contains any pattern as a substring or wildcard."""
    if not patterns:
        return False
    name_lower = name.lower()
    for pat in patterns:
        pat_lower = pat.lower()
        if any(c in pat_lower for c in ("*", "?", "[", "]")):
            if fnmatch.fnmatch(name_lower, pat_lower) or fnmatch.fnmatch(name_lower, f"*{pat_lower}*"):
                return True
        else:
            if pat_lower in name_lower:
                return True
    return False


def should_include_name(
    name: str,
    include_patterns: Optional[List[str]],
    exclude_patterns: Optional[List[str]],
) -> bool:
    """Evaluate inclusion/exclusion rules for a given file or folder name."""
    if exclude_patterns and matches_patterns(name, exclude_patterns):
        return False
    if include_patterns and not matches_patterns(name, include_patterns):
        return False
    return True


def write_file_safe(dest_file_path: str, content_bytes: bytes) -> None:
    """Write bytes to destination file with long-path resilience."""
    dest_dir = os.path.dirname(dest_file_path)
    ensure_directory(dest_dir)

    try:
        with open(dest_file_path, "wb") as dst_stream:
            dst_stream.write(content_bytes)
    except OSError as err:
        if os.name == "nt":
            long_p = normalize_long_path(dest_file_path)
            try:
                with open(long_p, "wb") as dst_stream:
                    dst_stream.write(content_bytes)
                return
            except Exception:
                short_dir = get_short_path_name(dest_dir)
                if short_dir:
                    alt_path = os.path.join(short_dir, os.path.basename(dest_file_path))
                    try:
                        with open(alt_path, "wb") as dst_stream:
                            dst_stream.write(content_bytes)
                        return
                    except Exception:
                        pass
        raise err


def write_file_delta(dest_file_path: str, content_bytes: bytes, rel_path: str, dry_run: bool = False) -> Tuple[int, int, int]:
    """Write file only if new or different content. Returns (created, updated, skipped)."""
    src_hash = hashlib.sha256(content_bytes).hexdigest()

    if file_exists_safe(dest_file_path):
        dest_size = get_path_size(dest_file_path)
        if dest_size == len(content_bytes):
            dest_hash = calculate_file_sha256(dest_file_path)
            if dest_hash == src_hash:
                return (0, 0, 1)

        if not dry_run:
            write_file_safe(dest_file_path, content_bytes)
        status_prefix = "[DRY-RUN UPDATED]" if dry_run else "[UPDATED]"
        print(f"  {status_prefix} {rel_path}")
        return (0, 1, 0)
    else:
        if not dry_run:
            write_file_safe(dest_file_path, content_bytes)
        status_prefix = "[DRY-RUN CREATED]" if dry_run else "[CREATED]"
        print(f"  {status_prefix} {rel_path}")
        return (1, 0, 0)


def sanitize_filename(name: str) -> str:
    """Sanitize string for safe usage as a cross-platform filename."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def extract_datamashup_artifacts(
    target_dir: str, mashup_bytes: bytes, dry_run: bool = False
) -> Tuple[int, int, int, int]:
    """Deconstruct DataMashup into Power Query (M) code and settings files."""
    created, updated, skipped, errors = 0, 0, 0, 0
    parser = DataMashupParser(mashup_bytes)
    artifacts = parser.extract_artifacts()
    if not artifacts:
        return (0, 0, 0, 0)

    # 1. Full Section1.m
    if artifacts.get("section1_m"):
        try:
            rel = "DataMashup_Extracted/Section1.m"
            dest = os.path.join(target_dir, rel)
            c, u, s = write_file_delta(dest, artifacts["section1_m"].encode("utf-8"), rel, dry_run=dry_run)
            created += c; updated += u; skipped += s
        except Exception as e:
            errors += 1
            print(f"  [SKIPPED - ERROR] Could not write Section1.m: {e}", file=sys.stderr)

    # 2. Decomposed individual query files
    queries = artifacts.get("queries", {})
    if queries:
        for q_name, q_code in queries.items():
            safe_name = sanitize_filename(q_name)
            rel = f"DataMashup_Extracted/Queries/{safe_name}.m"
            dest = os.path.join(target_dir, rel)
            try:
                c, u, s = write_file_delta(dest, q_code.encode("utf-8"), rel, dry_run=dry_run)
                created += c; updated += u; skipped += s
            except Exception as e:
                errors += 1
                print(f"  [SKIPPED - ERROR] Could not write query '{q_name}': {e}", file=sys.stderr)

    # 3. Settings & Configuration XML
    settings_items = [
        ("Package.xml", artifacts.get("package_xml")),
        ("Permissions.xml", artifacts.get("permissions_xml")),
        ("Metadata.xml", artifacts.get("metadata_xml")),
    ]
    for filename, xml_val in settings_items:
        if xml_val:
            rel = f"DataMashup_Extracted/{filename}"
            dest = os.path.join(target_dir, rel)
            try:
                c, u, s = write_file_delta(dest, xml_val.encode("utf-8"), rel, dry_run=dry_run)
                created += c; updated += u; skipped += s
            except Exception as e:
                errors += 1
                print(f"  [SKIPPED - ERROR] Could not write {filename}: {e}", file=sys.stderr)

    # 4. Metadata summary JSON
    meta_summary = artifacts.get("metadata_summary")
    if meta_summary and meta_summary.get("items"):
        rel = "DataMashup_Extracted/Metadata_Summary.json"
        dest = os.path.join(target_dir, rel)
        try:
            json_bytes = json.dumps(meta_summary, indent=2).encode("utf-8")
            c, u, s = write_file_delta(dest, json_bytes, rel, dry_run=dry_run)
            created += c; updated += u; skipped += s
        except Exception as e:
            errors += 1
            print(f"  [SKIPPED - ERROR] Could not write Metadata_Summary.json: {e}", file=sys.stderr)

    return (created, updated, skipped, errors)


def extract_datamodelschema_artifacts(
    target_dir: str, schema_bytes: bytes, dry_run: bool = False
) -> Tuple[int, int, int, int]:
    """
    Deconstruct DataModelSchema (TMSL JSON) into formatted model,
    tables metadata, partitions (M queries), and individual DAX measure files.
    """
    created, updated, skipped, errors = 0, 0, 0, 0
    parser = DataModelSchemaParser(schema_bytes)
    artifacts = parser.extract_artifacts()
    if not artifacts:
        return (0, 0, 0, 0)

    # 1. Pretty-printed schema JSON
    if artifacts.get("formatted_schema_json"):
        try:
            rel = "DataModelSchema_Extracted/DataModelSchema_Pretty.json"
            dest = os.path.join(target_dir, rel)
            c, u, s = write_file_delta(dest, artifacts["formatted_schema_json"].encode("utf-8"), rel, dry_run=dry_run)
            created += c; updated += u; skipped += s
        except Exception as e:
            errors += 1
            print(f"  [SKIPPED - ERROR] Could not write DataModelSchema_Pretty.json: {e}", file=sys.stderr)

    # 2. Decomposed DAX Measures
    measures = artifacts.get("measures", {})
    if measures:
        for m_key, m_info in measures.items():
            tbl_name = sanitize_filename(m_info["table"])
            meas_name = sanitize_filename(m_info["name"])
            rel = f"DataModelSchema_Extracted/Measures/{tbl_name}/{meas_name}.dax"
            dest = os.path.join(target_dir, rel)
            m_content = f"// Table: {m_info['table']}\n// Measure: {m_info['name']}\n"
            if m_info.get("description"):
                m_content += f"// Description: {m_info['description']}\n"
            if m_info.get("format_string"):
                m_content += f"// FormatString: {m_info['format_string']}\n"
            m_content += f"\n{m_info['expression']}\n"
            try:
                c, u, s = write_file_delta(dest, m_content.encode("utf-8"), rel, dry_run=dry_run)
                created += c; updated += u; skipped += s
            except Exception as e:
                errors += 1
                print(f"  [SKIPPED - ERROR] Could not write DAX measure '{meas_name}': {e}", file=sys.stderr)

    # 3. Table definitions
    tables = artifacts.get("tables", {})
    if tables:
        for t_name, t_meta in tables.items():
            safe_t = sanitize_filename(t_name)
            rel = f"DataModelSchema_Extracted/Tables/{safe_t}.json"
            dest = os.path.join(target_dir, rel)
            try:
                t_bytes = json.dumps(t_meta, indent=2).encode("utf-8")
                c, u, s = write_file_delta(dest, t_bytes, rel, dry_run=dry_run)
                created += c; updated += u; skipped += s
            except Exception as e:
                errors += 1
                print(f"  [SKIPPED - ERROR] Could not write table metadata '{safe_t}': {e}", file=sys.stderr)

    # 4. M Partitions (Power Query source code inside Model tables)
    m_partitions = artifacts.get("m_partitions", {})
    if m_partitions:
        for p_name, p_code in m_partitions.items():
            safe_p = sanitize_filename(p_name)
            rel = f"DataModelSchema_Extracted/Partitions/{safe_p}.m"
            dest = os.path.join(target_dir, rel)
            try:
                c, u, s = write_file_delta(dest, p_code.encode("utf-8"), rel, dry_run=dry_run)
                created += c; updated += u; skipped += s
            except Exception as e:
                errors += 1
                print(f"  [SKIPPED - ERROR] Could not write M partition '{safe_p}': {e}", file=sys.stderr)

    # 5. Relationships
    relationships = artifacts.get("relationships", [])
    if relationships:
        rel = "DataModelSchema_Extracted/Relationships.json"
        dest = os.path.join(target_dir, rel)
        try:
            r_bytes = json.dumps(relationships, indent=2).encode("utf-8")
            c, u, s = write_file_delta(dest, r_bytes, rel, dry_run=dry_run)
            created += c; updated += u; skipped += s
        except Exception as e:
            errors += 1
            print(f"  [SKIPPED - ERROR] Could not write Relationships.json: {e}", file=sys.stderr)

    return (created, updated, skipped, errors)


def unpack_single_pbix(
    pbix_path: str,
    dry_run: bool = False,
    parse_mashup: bool = True,
    parse_schema: bool = True,
) -> Tuple[int, int, int, int]:
    """
    Unpack a single .pbix or .pbit file into a folder with the same base name.
    Extracts DataMashup (Power Query M) and DataModelSchema (TMSL JSON) if present.
    
    Returns:
        (created_count, updated_count, skipped_count, error_count)
    """
    pbix_dir = os.path.dirname(pbix_path)
    base_name = os.path.splitext(os.path.basename(pbix_path))[0]
    target_dir = os.path.join(pbix_dir, base_name)

    created_count = 0
    updated_count = 0
    skipped_count = 0
    error_count = 0

    print(f"\nProcessing: {pbix_path}")
    print(f"Target directory: {target_dir}")

    if not zipfile.is_zipfile(pbix_path):
        print(f"  [ERROR] Not a valid zip/pbix/pbit archive: {pbix_path}", file=sys.stderr)
        return (0, 0, 0, 1)

    datamashup_bytes: Optional[bytes] = None
    datamodelschema_bytes: Optional[bytes] = None

    try:
        with zipfile.ZipFile(pbix_path, "r") as zf:
            for zip_info in zf.infolist():
                if zip_info.is_dir() or zip_info.filename.endswith("/"):
                    continue

                rel_path = os.path.normpath(zip_info.filename)
                if rel_path.startswith("..") or os.path.isabs(rel_path):
                    print(f"  [WARN] Skipping unsafe zip path: {zip_info.filename}")
                    continue

                dest_file_path = os.path.join(target_dir, rel_path)

                try:
                    with zf.open(zip_info) as src_stream:
                        content_bytes = src_stream.read()

                    # Intercept special files for deep deconstruction
                    if rel_path == "DataMashup":
                        datamashup_bytes = content_bytes
                    elif rel_path in ("DataModelSchema", "DataModelSchema.json"):
                        datamodelschema_bytes = content_bytes

                    c, u, s = write_file_delta(dest_file_path, content_bytes, rel_path, dry_run=dry_run)
                    created_count += c
                    updated_count += u
                    skipped_count += s

                except Exception as file_err:
                    error_count += 1
                    print(f"  [SKIPPED - ERROR] Could not create/write '{rel_path}': {file_err}", file=sys.stderr)

    except Exception as exc:
        print(f"  [ERROR] Failed to unpack {pbix_path}: {exc}", file=sys.stderr)
        return (created_count, updated_count, skipped_count, error_count + 1)

    # 1. Deep DataMashup extraction
    if parse_mashup:
        if datamashup_bytes:
            print("  -> Decompiling DataMashup (Power Query M & Settings)...")
            mc, mu, ms, me = extract_datamashup_artifacts(target_dir, datamashup_bytes, dry_run=dry_run)
            created_count += mc
            updated_count += mu
            skipped_count += ms
            error_count += me
        else:
            print("  [NOTE] No DataMashup entry in this file (e.g. Live Connection or Direct Lake).")

    # 2. Deep DataModelSchema extraction
    if parse_schema:
        if datamodelschema_bytes:
            print("  -> Decompiling DataModelSchema (TMSL JSON, DAX Measures & Tables)...")
            sc, su, ss, se = extract_datamodelschema_artifacts(target_dir, datamodelschema_bytes, dry_run=dry_run)
            created_count += sc
            updated_count += su
            skipped_count += ss
            error_count += se
        else:
            print("  [NOTE] No DataModelSchema entry in this file (standard in .pbit templates; .pbix files store raw VertiPaq in DataModel).")

    action_label = "Dry-run summary" if dry_run else "Unpacked"
    error_note = f", {error_count} failed/skipped" if error_count > 0 else ""
    print(f"  -> {action_label}: {created_count} created, {updated_count} updated, {skipped_count} unchanged{error_note}.")
    return (created_count, updated_count, skipped_count, error_count)


def find_pbix_files(
    root_folder: str,
    recursive: bool,
    file_include: Optional[List[str]],
    file_exclude: Optional[List[str]],
    folder_include: Optional[List[str]],
    folder_exclude: Optional[List[str]],
) -> List[str]:
    """Discover all .pbix and .pbit files adhering to path, recursion, and inclusion/exclusion filters."""
    root_folder = os.path.abspath(root_folder)
    pbix_files: List[str] = []

    if not os.path.isdir(root_folder):
        raise ValueError(f"Root path is not a directory: {root_folder}")

    valid_exts = (".pbix", ".pbit")

    if not recursive:
        for entry in os.listdir(root_folder):
            if entry.lower().endswith(valid_exts) and os.path.isfile(os.path.join(root_folder, entry)):
                if should_include_name(entry, file_include, file_exclude):
                    pbix_files.append(os.path.join(root_folder, entry))
        return sorted(pbix_files)

    for current_dir, subdirs, files in os.walk(root_folder, topdown=True):
        if current_dir != root_folder:
            rel_dir = os.path.relpath(current_dir, root_folder)
            dir_name = os.path.basename(current_dir)
            if not should_include_name(dir_name, folder_include, folder_exclude) and not should_include_name(rel_dir, folder_include, folder_exclude):
                subdirs.clear()
                continue

        subdirs[:] = [
            d for d in subdirs
            if should_include_name(d, folder_include, folder_exclude)
        ]

        for file_name in files:
            if file_name.lower().endswith(valid_exts):
                if should_include_name(file_name, file_include, file_exclude):
                    pbix_files.append(os.path.join(current_dir, file_name))

    return sorted(pbix_files)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Unpack PBIX and PBIT files into folders maintaining internal tree structure with delta sync, DataMashup & DataModelSchema deconstruction."
    )
    parser.add_argument(
        "-p", "--path",
        required=True,
        help="Target folder containing .pbix/.pbit files."
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Recursively scan subfolders for .pbix/.pbit files."
    )
    parser.add_argument(
        "--file-include",
        nargs="+",
        help="Only process files whose names contain any of these strings/patterns."
    )
    parser.add_argument(
        "--file-exclude",
        nargs="+",
        help="Skip files whose names contain any of these strings/patterns."
    )
    parser.add_argument(
        "--folder-include",
        nargs="+",
        help="When searching folders recursively, only traverse folders containing any of these strings/patterns."
    )
    parser.add_argument(
        "--folder-exclude",
        nargs="+",
        help="When searching folders recursively, skip folders containing any of these strings/patterns."
    )
    parser.add_argument(
        "--no-mashup",
        action="store_true",
        help="Disable deep extraction of DataMashup (Power Query M and settings)."
    )
    parser.add_argument(
        "--no-schema",
        action="store_true",
        help="Disable deep extraction of DataModelSchema (TMSL JSON and DAX measures)."
    )
    parser.add_argument(
        "-n", "--dry-run",
        action="store_true",
        help="Simulate the unpacking without modifying files."
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    try:
        matched_pbix = find_pbix_files(
            root_folder=args.path,
            recursive=args.recursive,
            file_include=args.file_include,
            file_exclude=args.file_exclude,
            folder_include=args.folder_include,
            folder_exclude=args.folder_exclude,
        )
    except Exception as err:
        print(f"Error finding files: {err}", file=sys.stderr)
        return 1

    if not matched_pbix:
        print("No matching .pbix/.pbit files found.")
        return 0

    print(f"Found {len(matched_pbix)} matching file(s).")
    total_created = 0
    total_updated = 0
    total_skipped = 0
    total_errors = 0

    for pbix in matched_pbix:
        c, u, s, e = unpack_single_pbix(
            pbix,
            dry_run=args.dry_run,
            parse_mashup=(not args.no_mashup),
            parse_schema=(not args.no_schema),
        )
        total_created += c
        total_updated += u
        total_skipped += s
        total_errors += e

    print("\n" + "=" * 50)
    print("Summary:")
    print(f"  Files processed      : {len(matched_pbix)}")
    print(f"  Files created        : {total_created}")
    print(f"  Files updated        : {total_updated}")
    print(f"  Files unchanged      : {total_skipped}")
    if total_errors > 0:
        print(f"  Files/items errored  : {total_errors} (skipped gracefully)")
    print("=" * 50)
    return 0


if __name__ == "__main__":
    sys.exit(main())
