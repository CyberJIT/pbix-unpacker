"""
mashup_parser.py

Module to parse and extract Microsoft Power Query DataMashup binary streams (MS-QDEFF specification)
into human-readable plain text Power Query M code, individual query files, and XML/JSON configuration settings.
"""

import hashlib
import io
import json
import os
import re
import struct
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple


def pretty_format_xml(xml_bytes_or_str) -> str:
    """Format XML string or bytes with readable indentations."""
    try:
        raw_str = xml_bytes_or_str.decode("utf-8") if isinstance(xml_bytes_or_str, bytes) else xml_bytes_or_str
        elem = ET.fromstring(raw_str)
        rough = ET.tostring(elem, encoding="utf-8")
        reparsed = minidom.parseString(rough)
        return reparsed.toprettyxml(indent="  ")
    except Exception:
        if isinstance(xml_bytes_or_str, bytes):
            return xml_bytes_or_str.decode("utf-8", errors="replace")
        return str(xml_bytes_or_str)


def parse_m_queries(section_text: str) -> Dict[str, str]:
    """
    Decompose a Power Query M Section document (such as Formulas/Section1.m)
    into individual shared query expressions.
    
    Returns:
        Dict mapping query_name -> M expression string.
    """
    # Regex to find top-level shared declarations in Section1:
    # e.g., shared #"Query Name" = ... or shared QueryName = ...
    pattern = re.compile(
        r'^[ \t]*shared[ \t]+(#\"[^\"]+\"|[a-zA-Z_][a-zA-Z0-9_]*)[ \t]*=',
        re.MULTILINE
    )
    matches = list(pattern.finditer(section_text))
    queries: Dict[str, str] = {}

    if not matches:
        # Fallback: check if entire file is a single query without 'shared'
        stripped = section_text.strip()
        if stripped and not stripped.startswith("section "):
            queries["Query"] = stripped
        return queries

    for i, m in enumerate(matches):
        raw_name = m.group(1)
        # Strip #"..." quotes if present
        clean_name = raw_name[2:-1] if raw_name.startswith('#"') and raw_name.endswith('"') else raw_name
        start_pos = m.start()
        end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(section_text)
        expr = section_text[start_pos:end_pos].strip()
        if expr.endswith(";"):
            expr = expr[:-1].rstrip()
        queries[clean_name] = expr

    return queries


def parse_metadata_xml(xml_content: str) -> Dict[str, Any]:
    """
    Parse MS-QDEFF Metadata XML into structured dictionary with typed values.
    """
    result: Dict[str, Any] = {"items": {}, "raw_items_count": 0}
    try:
        root = ET.fromstring(xml_content)
        items = root.findall(".//{*}Item")
        result["raw_items_count"] = len(items)

        for item in items:
            loc = item.find(".//{*}ItemPath")
            item_path = loc.text if loc is not None and loc.text else "AllFormulas"
            entries: Dict[str, Any] = {}

            for entry in item.findall(".//{*}Entry"):
                entry_type = entry.attrib.get("Type", "")
                raw_val = entry.attrib.get("Value", "")
                if not entry_type:
                    continue

                # Type prefix decoding as per MS-QDEFF Section 2.5.1
                # l: integer / boolean
                # f: decimal float
                # s: string
                # c: GUID / UUID
                # d: date/time ISO string
                if raw_val.startswith("l"):
                    try:
                        val_num = int(raw_val[1:])
                        # 0/1 for booleans in certain properties
                        if entry_type in ("AddedToDataModel", "IsHidden", "IsParameterQuery", "IsNavigationStep"):
                            entries[entry_type] = bool(val_num)
                        else:
                            entries[entry_type] = val_num
                    except ValueError:
                        entries[entry_type] = raw_val[1:]
                elif raw_val.startswith("f"):
                    try:
                        entries[entry_type] = float(raw_val[1:])
                    except ValueError:
                        entries[entry_type] = raw_val[1:]
                elif raw_val.startswith("s"):
                    entries[entry_type] = raw_val[1:]
                elif raw_val.startswith("c"):
                    entries[entry_type] = raw_val[1:]
                elif raw_val.startswith("d"):
                    entries[entry_type] = raw_val[1:]
                else:
                    entries[entry_type] = raw_val

            result["items"][item_path] = entries
    except Exception as exc:
        result["error"] = str(exc)

    return result


class DataMashupParser:
    """
    Parser for Microsoft MS-QDEFF binary DataMashup file format.
    """

    def __init__(self, data: bytes):
        self.data = data
        self.version = 0
        self.package_bytes = b""
        self.permissions_bytes = b""
        self.metadata_bytes = b""
        self.bindings_bytes = b""
        self._parsed = False

    def parse_structure(self) -> bool:
        """Parse top-level MS-QDEFF binary fields according to Section 2.2."""
        if len(self.data) < 8:
            return False

        try:
            offset = 0
            self.version, pkg_len = struct.unpack_from("<II", self.data, offset)
            offset += 8

            if offset + pkg_len > len(self.data):
                return False
            self.package_bytes = self.data[offset : offset + pkg_len]
            offset += pkg_len

            if offset + 4 <= len(self.data):
                perm_len = struct.unpack_from("<I", self.data, offset)[0]
                offset += 4
                self.permissions_bytes = self.data[offset : offset + perm_len]
                offset += perm_len

            if offset + 4 <= len(self.data):
                meta_len = struct.unpack_from("<I", self.data, offset)[0]
                offset += 4
                self.metadata_bytes = self.data[offset : offset + meta_len]
                offset += meta_len

            if offset + 4 <= len(self.data):
                bindings_len = struct.unpack_from("<I", self.data, offset)[0]
                offset += 4
                self.bindings_bytes = self.data[offset : offset + bindings_len]
                offset += bindings_len

            self._parsed = True
            return True
        except Exception:
            return False

    def extract_artifacts(self) -> Dict[str, Any]:
        """
        Decompile DataMashup into plain text M formulas, individual queries,
        and settings.
        
        Returns dict containing:
          - 'package_files': dict of {rel_path: bytes}
          - 'section1_m': full Section1.m text (if present)
          - 'queries': dict of {query_name: query_m_code}
          - 'package_xml': formatted Package.xml (if present)
          - 'permissions_xml': formatted Permissions.xml (if present)
          - 'metadata_xml': formatted Metadata.xml (if present)
          - 'metadata_summary': parsed JSON dict of query metadata
          - 'permission_bindings': raw bytes of permission bindings
        """
        if not self._parsed:
            if not self.parse_structure():
                return {}

        results: Dict[str, Any] = {
            "version": self.version,
            "package_files": {},
            "queries": {},
            "section1_m": None,
            "package_xml": None,
            "permissions_xml": None,
            "metadata_xml": None,
            "metadata_summary": {},
            "permission_bindings": self.bindings_bytes,
        }

        # 1. Decompress Package Parts ZIP
        if self.package_bytes and zipfile.is_zipfile(io.BytesIO(self.package_bytes)):
            try:
                with zipfile.ZipFile(io.BytesIO(self.package_bytes), "r") as zf:
                    for name in zf.namelist():
                        if name.endswith("/"):
                            continue
                        content = zf.read(name)
                        results["package_files"][name] = content

                        if name == "Formulas/Section1.m":
                            text_m = content.decode("utf-8", errors="replace")
                            results["section1_m"] = text_m
                            results["queries"] = parse_m_queries(text_m)
                        elif name == "Config/Package.xml":
                            results["package_xml"] = pretty_format_xml(content)
            except Exception as err:
                results["package_error"] = str(err)

        # 2. Permissions XML
        if self.permissions_bytes:
            results["permissions_xml"] = pretty_format_xml(self.permissions_bytes)

        # 3. Metadata Stream (MS-QDEFF Section 2.5)
        # Version (4 bytes) | Metadata XML Length (4 bytes) | Metadata XML | Content Length (4 bytes) | Content
        if len(self.metadata_bytes) >= 8:
            try:
                m_ver, xml_len = struct.unpack_from("<II", self.metadata_bytes, 0)
                if 8 + xml_len <= len(self.metadata_bytes):
                    xml_raw = self.metadata_bytes[8 : 8 + xml_len]
                    results["metadata_xml"] = pretty_format_xml(xml_raw)
                    results["metadata_summary"] = parse_metadata_xml(xml_raw.decode("utf-8", errors="replace"))

                    # Content length & optional package
                    content_offset = 8 + xml_len
                    if content_offset + 4 <= len(self.metadata_bytes):
                        cnt_len = struct.unpack_from("<I", self.metadata_bytes, content_offset)[0]
                        if cnt_len > 0 and content_offset + 4 + cnt_len <= len(self.metadata_bytes):
                            results["metadata_content_bytes"] = self.metadata_bytes[
                                content_offset + 4 : content_offset + 4 + cnt_len
                            ]
            except Exception as err:
                results["metadata_error"] = str(err)

        return results
