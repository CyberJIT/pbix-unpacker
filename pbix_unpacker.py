#!/usr/bin/env python3
"""
pbix_unpacker.py

A tool to extract .pbix files (treated as zip archives) into subfolders named
after the base name of each .pbix file in the same directory, maintaining internal
folder hierarchies, and only replacing files if they are new or have different contents.

Supports recursive scanning, file name filters (include/exclude), and folder name filters (include/exclude).
"""

import argparse
import fnmatch
import hashlib
import os
import sys
import zipfile
from typing import List, Optional, Tuple


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
    """Calculate SHA256 checksum from an on-disk file."""
    with open(filepath, "rb") as f:
        return calculate_stream_sha256(f, chunk_size=chunk_size)


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


def unpack_single_pbix(
    pbix_path: str, dry_run: bool = False
) -> Tuple[int, int, int]:
    """
    Unpack a single .pbix file into a folder with the same base name.
    
    Returns:
        (created_count, updated_count, skipped_count)
    """
    pbix_dir = os.path.dirname(pbix_path)
    base_name = os.path.splitext(os.path.basename(pbix_path))[0]
    target_dir = os.path.join(pbix_dir, base_name)

    created_count = 0
    updated_count = 0
    skipped_count = 0

    print(f"\nProcessing: {pbix_path}")
    print(f"Target directory: {target_dir}")

    if not zipfile.is_zipfile(pbix_path):
        print(f"  [ERROR] Not a valid zip/pbix archive: {pbix_path}", file=sys.stderr)
        return (0, 0, 0)

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
                dest_dir = os.path.dirname(dest_file_path)

                # Compute hash of archive content
                with zf.open(zip_info) as src_stream:
                    src_hash = calculate_stream_sha256(src_stream)

                if os.path.exists(dest_file_path):
                    # Check size first as fast path, then hash
                    dest_size = os.path.getsize(dest_file_path)
                    if dest_size == zip_info.file_size:
                        dest_hash = calculate_file_sha256(dest_file_path)
                        if dest_hash == src_hash:
                            skipped_count += 1
                            continue

                    # Content differs -> update
                    if not dry_run:
                        os.makedirs(dest_dir, exist_ok=True)
                        with zf.open(zip_info) as src_stream, open(dest_file_path, "wb") as dst_stream:
                            dst_stream.write(src_stream.read())
                    updated_count += 1
                    status_prefix = "[DRY-RUN UPDATED]" if dry_run else "[UPDATED]"
                    print(f"  {status_prefix} {rel_path}")
                else:
                    # New file
                    if not dry_run:
                        os.makedirs(dest_dir, exist_ok=True)
                        with zf.open(zip_info) as src_stream, open(dest_file_path, "wb") as dst_stream:
                            dst_stream.write(src_stream.read())
                    created_count += 1
                    status_prefix = "[DRY-RUN CREATED]" if dry_run else "[CREATED]"
                    print(f"  {status_prefix} {rel_path}")

    except Exception as exc:
        print(f"  [ERROR] Failed to unpack {pbix_path}: {exc}", file=sys.stderr)
        return (created_count, updated_count, skipped_count)

    action_label = "Dry-run summary" if dry_run else "Unpacked"
    print(f"  -> {action_label}: {created_count} created, {updated_count} updated, {skipped_count} unchanged.")
    return (created_count, updated_count, skipped_count)


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
        # Check files directly in root_folder
        for entry in os.listdir(root_folder):
            if entry.lower().endswith(".pbix") and os.path.isfile(os.path.join(root_folder, entry)):
                if should_include_name(entry, file_include, file_exclude):
                    pbix_files.append(os.path.join(root_folder, entry))
        return sorted(pbix_files)

    for current_dir, subdirs, files in os.walk(root_folder, topdown=True):
        # Prune subdirectories according to folder filters
        # Note: We do not filter the initial root_folder itself
        if current_dir != root_folder:
            rel_dir = os.path.relpath(current_dir, root_folder)
            dir_name = os.path.basename(current_dir)
            # Evaluate folder filter both on basename and relative directory path
            if not should_include_name(dir_name, folder_include, folder_exclude) and not should_include_name(rel_dir, folder_include, folder_exclude):
                subdirs.clear()
                continue

        # In-place filter subdirs for next recursion level
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
        description="Unpack PBIX files into folders maintaining internal tree structure with change detection."
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

    for pbix in matched_pbix:
        c, u, s = unpack_single_pbix(pbix, dry_run=args.dry_run)
        total_created += c
        total_updated += u
        total_skipped += s

    print("\n" + "=" * 50)
    print("Summary:")
    print(f"  PBIX files processed : {len(matched_pbix)}")
    print(f"  Files created        : {total_created}")
    print(f"  Files updated        : {total_updated}")
    print(f"  Files unchanged      : {total_skipped}")
    print("=" * 50)
    return 0


if __name__ == "__main__":
    sys.exit(main())
