"""
layout_parser.py

Module to deconstruct and extract Power BI Report/Layout JSON and related report artifacts.
Deconstructs monolithic Layout into per-page definitions, individual visual configurations,
filters, and bookmarks. Also parses DiagramLayout, LinguisticSchema, and report metadata.
"""

import json
import os
import re
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Tuple


def decode_text_auto(data: bytes) -> str:
    """Decode bytes by testing UTF-16LE, UTF-8-SIG, UTF-8, and Latin1."""
    if data.startswith(b"\xff\xfe"):
        return data.decode("utf-16le", errors="replace")
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig", errors="replace")
    if len(data) >= 2 and data[1] == 0:
        try:
            return data.decode("utf-16le")
        except UnicodeDecodeError:
            pass
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("utf-16le", errors="replace")


def sanitize_filename(name: str) -> str:
    """Sanitize string for safe file/folder names."""
    clean = re.sub(r'[\\/*?:"<>|%]', "_", name).strip()
    return clean if clean else "Unnamed"


def pretty_xml(xml_bytes_or_str) -> str:
    """Format XML string with readable indentations."""
    try:
        raw_str = xml_bytes_or_str.decode("utf-8") if isinstance(xml_bytes_or_str, bytes) else xml_bytes_or_str
        elem = ET.fromstring(raw_str)
        rough = ET.tostring(elem, encoding="utf-8")
        reparsed = minidom.parseString(rough)
        return reparsed.toprettyxml(indent="  ")
    except Exception:
        if isinstance(xml_bytes_or_str, bytes):
            return decode_text_auto(xml_bytes_or_str)
        return str(xml_bytes_or_str)


class ReportLayoutParser:
    """
    Parser for Power BI Report/Layout JSON.
    """

    def __init__(self, raw_bytes_or_str):
        if isinstance(raw_bytes_or_str, bytes):
            self.raw_text = decode_text_auto(raw_bytes_or_str)
        else:
            self.raw_text = raw_bytes_or_str
        self.layout_dict: Optional[Dict[str, Any]] = None
        self._parse()

    def _parse(self) -> None:
        try:
            self.layout_dict = json.loads(self.raw_text)
        except Exception:
            self.layout_dict = None

    def is_valid(self) -> bool:
        return isinstance(self.layout_dict, dict) and "sections" in self.layout_dict

    def extract_artifacts(self) -> Dict[str, Any]:
        """
        Deconstruct Layout into pages and visuals.
        Returns:
          - 'layout_pretty': full formatted layout JSON
          - 'pages': dict of {page_name: {'meta': dict, 'visuals': list}}
          - 'config': parsed report-level configuration
        """
        if not self.is_valid():
            return {}

        results: Dict[str, Any] = {
            "layout_pretty": json.dumps(self.layout_dict, indent=2),
            "pages": {},
            "config": None,
        }

        # Parse report-level config
        if "config" in self.layout_dict and isinstance(self.layout_dict["config"], str):
            try:
                results["config"] = json.loads(self.layout_dict["config"])
            except Exception:
                results["config"] = self.layout_dict["config"]

        sections = self.layout_dict.get("sections", [])
        for idx, sec in enumerate(sections):
            display_name = sec.get("displayName") or f"Page_{idx+1}"
            sec_name = sec.get("name") or f"Section_{idx+1}"
            page_key = f"{idx+1:02d}_{sanitize_filename(display_name)}"

            page_meta = {
                "id": sec.get("id"),
                "name": sec_name,
                "displayName": display_name,
                "ordinal": sec.get("ordinal", idx),
                "width": sec.get("width"),
                "height": sec.get("height"),
                "displayOption": sec.get("displayOption"),
                "filters": None,
                "config": None,
            }

            # Parse page filters
            if "filters" in sec and isinstance(sec["filters"], str):
                try:
                    page_meta["filters"] = json.loads(sec["filters"])
                except Exception:
                    page_meta["filters"] = sec["filters"]

            # Parse page config
            if "config" in sec and isinstance(sec["config"], str):
                try:
                    page_meta["config"] = json.loads(sec["config"])
                except Exception:
                    page_meta["config"] = sec["config"]

            visuals_list = []
            for v_idx, v_container in enumerate(sec.get("visualContainers", [])):
                v_id = v_container.get("id", v_idx)
                v_entry = {
                    "id": v_id,
                    "x": v_container.get("x"),
                    "y": v_container.get("y"),
                    "z": v_container.get("z"),
                    "width": v_container.get("width"),
                    "height": v_container.get("height"),
                    "visualType": "Unknown",
                    "title": None,
                    "config": None,
                    "filters": None,
                    "dataTransforms": None,
                }

                # Unpack inner JSON config strings
                if "config" in v_container and isinstance(v_container["config"], str):
                    try:
                        c_obj = json.loads(v_container["config"])
                        v_entry["config"] = c_obj
                        sv = c_obj.get("singleVisual", {})
                        v_entry["visualType"] = sv.get("visualType", "Unknown")
                        # Try to find visual title if present
                        vc_title = sv.get("vcObjects", {}).get("title", [{}])[0].get("properties", {}).get("text", {})
                        if isinstance(vc_title, dict) and "expr" in vc_title:
                            v_entry["title"] = vc_title.get("expr", {}).get("Literal", {}).get("Value")
                    except Exception:
                        v_entry["config"] = v_container["config"]

                if "filters" in v_container and isinstance(v_container["filters"], str):
                    try:
                        v_entry["filters"] = json.loads(v_container["filters"])
                    except Exception:
                        v_entry["filters"] = v_container["filters"]

                if "dataTransforms" in v_container and isinstance(v_container["dataTransforms"], str):
                    try:
                        v_entry["dataTransforms"] = json.loads(v_container["dataTransforms"])
                    except Exception:
                        v_entry["dataTransforms"] = v_container["dataTransforms"]

                visuals_list.append(v_entry)

            results["pages"][page_key] = {
                "metadata": page_meta,
                "visuals": visuals_list,
            }

        return results
