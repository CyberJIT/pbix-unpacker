#!/usr/bin/env python3
"""
pbix_unpacker.py

A tool to extract .pbix files (treated as zip archives) into subfolders named
after the base name of each .pbix file in the same directory, maintaining internal
folder hierarchies, and only replacing files if they are new or have different contents.

Supports recursive scanning, file name filters (include/exclude), folder name filters (include/exclude),
Windows extended-length paths (\\\\?\\) and short 8.3 path fallbacks, and graceful error recovery.
"""

import argparse
import ctypes
import fnmatch
import hashlib
import os
import sys
import zipfile
from typing import List, Optional, Tuple


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
        # UNC path: \\server\share -> \\?\UNC\server\share
        return "\\\\?\\UNC\\" + abs_path[2:]
    return "\\\\?\\" + abs_path


def get_short_path_name(long_name: str) -> Optional[str]:
    """
    Retrieve Windows 8.3 short path equivalent using GetShortPathNameW API.
    Returns None if not on Windows or if short path generation is disabled/unavailable.
    """
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
    """
    Create directory ensuring support for long paths and short path workarounds.
    """
    try:
        os.makedirs(dir_path, exist_ok=True)
    except (OSError, FileNotFoundError) as err:
        # Check if Windows path length or similar error
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
    """
    Check if a name contains any pattern as a substring (case-insensitive)
    or matches a wildcard pattern.
    """
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
    """
    Evaluate inclusion/exclusion rules for a given file or folder name.
    1. If exclude patterns are given and name matches, exclude.
    2. If include patterns are given, name must match at least one.
    3. Otherwise include.
    """
    if exclude_patterns and matches_patterns(name, exclude_patterns):
        return False
    if include_patterns and not matches_patterns(name, include_patterns):
        return False
    return True


def write_file_safe(dest_file_path: str, content_bytes: bytes) -> None:
    """
    Write bytes to destination file, using extended long paths or short path
    equivalents if standard open fails.
    """
    dest_dir = os.path.dirname(dest_file_path)
    ensure_directory(dest_dir)

    try:
        with open(dest_file_path, "wb") as dst_stream:
            dst_stream.write(content_bytes)
    except OSError as err:
        # Long path or file system error fallback on Windows
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


def unpack_single_pbix(
    pbix_path: str, dry_run: bool = False
) -> Tuple[int, int, int, int]:
    """
    Unpack a single .pbix file into a folder with the same base name.
    
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
        print(f"  [ERROR] Not a valid zip/pbix archive: {pbix_path}", file=sys.stderr)
        return (0, 0, 0, 1)

    try:
        with zipfile.ZipFile(pbix_path, "r") as zf:
            for zip_info in zf.infolist():
                # Skip directory entries
                if zip_info.is_dir() or zip_info.filename.endswith("/"):
                    continue

                # Protect against zip slip / path traversal
                rel_path = os.path.normpath(zip_info.filename)
                if rel_path.startswith("..") or os.path.isabs(rel_path):
                    print(f"  [WARN] Skipping unsafe zip path: {zip_info.filename}")
                    continue

                dest_file_path = os.path.join(target_dir, rel_path)

                try:
                    # Read content & compute hash of archive stream
                    with zf.open(zip_info) as src_stream:
                        content_bytes = src_stream.read()
                    src_hash = hashlib.sha256(content_bytes).hexdigest()

                    if file_exists_safe(dest_file_path):
                        # Fast check on length then SHA256
                        dest_size = get_path_size(dest_file_path)
                        if dest_size == zip_info.file_size:
                            dest_hash = calculate_file_sha256(dest_file_path)
                            if dest_hash == src_hash:
                                skipped_count += 1
                                continue

                        # Content differs -> update
                        if not dry_run:
                            write_file_safe(dest_file_path, content_bytes)
                        updated_count += 1
                        status_prefix = "[DRY-RUN UPDATED]" if dry_run else "[UPDATED]"
                        print(f"  {status_prefix} {rel_path}")
                    else:
                        # New file
                        if not dry_run:
                            write_file_safe(dest_file_path, content_bytes)
                        created_count += 1
                        status_prefix = "[DRY-RUN CREATED]" if dry_run else "[CREATED]"
                        print(f"  {status_prefix} {rel_path}")

                except Exception as file_err:
                    error_count += 1
                    print(f"  [SKIPPED - ERROR] Could not create/write '{rel_path}': {file_err}", file=sys.stderr)

    except Exception as exc:
        print(f"  [ERROR] Failed to unpack {pbix_path}: {exc}", file=sys.stderr)
        return (created_count, updated_count, skipped_count, error_count + 1)

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
    """Discover all .pbix files adhering to path, recursion, and inclusion/exclusion filters."""
    root_folder = os.path.abspath(root_folder)
    pbix_files: List[str] = []

    if not os.path.isdir(root_folder):
        raise ValueError(f"Root path is not a directory: {root_folder}")

    if not recursive:
        for entry in os.listdir(root_folder):
            if entry.lower().endswith(".pbix") and os.path.isfile(os.path.join(root_folder, entry)):
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
            if file_name.lower().endswith(".pbix"):
                if should_include_name(file_name, file_include, file_exclude):
                    pbix_files.append(os.path.join(current_dir, file_name))

    return sorted(pbix_files)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Unpack PBIX files into folders maintaining internal tree structure with change detection and long-path resilience."
    )
    parser.add_argument(
        "-p", "--path",
        required=True,
        help="Target folder containing .pbix files."
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Recursively scan subfolders for .pbix files."
    )
    parser.add_argument(
        "--file-include",
        nargs="+",
        help="Only process .pbix files whose names contain any of these strings/patterns."
    )
    parser.add_argument(
        "--file-exclude",
        nargs="+",
        help="Skip .pbix files whose names contain any of these strings/patterns."
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
        print("No matching .pbix files found.")
        return 0

    print(f"Found {len(matched_pbix)} matching .pbix file(s).")
    total_created = 0
    total_updated = 0
    total_skipped = 0
    total_errors = 0

    for pbix in matched_pbix:
        c, u, s, e = unpack_single_pbix(pbix, dry_run=args.dry_run)
        total_created += c
        total_updated += u
        total_skipped += s
        total_errors += e

    print("\n" + "=" * 50)
    print("Summary:")
    print(f"  PBIX files processed : {len(matched_pbix)}")
    print(f"  Files created        : {total_created}")
    print(f"  Files updated        : {total_updated}")
    print(f"  Files unchanged      : {total_skipped}")
    if total_errors > 0:
        print(f"  Files/items errored  : {total_errors} (skipped gracefully)")
    print("=" * 50)
    return 0


if __name__ == "__main__":
    sys.exit(main())
