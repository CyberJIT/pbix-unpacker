"""
schema_parser.py

Module to parse and extract Microsoft Power BI DataModelSchema (TMSL JSON in UTF-16LE/UTF-8)
found in .pbit templates, .pbix files, and PBIP models.
Deconstructs tables, columns, partitions (M queries), relationships, and individual DAX measures.
"""

import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple


def decode_schema_bytes(data: bytes) -> str:
    """
    Decode DataModelSchema bytes trying UTF-16LE, UTF-16BE, UTF-8-SIG, and UTF-8.
    """
    # Detect BOM or try standard encodings
    if data.startswith(b"\xff\xfe"):
        return data.decode("utf-16le", errors="replace")
    elif data.startswith(b"\xfe\xff"):
        return data.decode("utf-16be", errors="replace")
    elif data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig", errors="replace")

    # Often DataModelSchema starts without BOM as UTF-16LE (e.g. b'{\x00"\x00')
    if len(data) >= 2 and data[1] == 0:
        try:
            return data.decode("utf-16le")
        except UnicodeDecodeError:
            pass

    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("utf-16le", errors="replace")


def sanitize_name(name: str) -> str:
    """Sanitize identifier for filesystem naming."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


class DataModelSchemaParser:
    """
    Parser for Tabular Model Scripting Language (TMSL) DataModelSchema JSON.
    """

    def __init__(self, raw_bytes_or_str):
        if isinstance(raw_bytes_or_str, bytes):
            self.raw_text = decode_schema_bytes(raw_bytes_or_str)
        else:
            self.raw_text = raw_bytes_or_str
        self.schema_dict: Optional[Dict[str, Any]] = None
        self._parse_json()

    def _parse_json(self) -> None:
        try:
            self.schema_dict = json.loads(self.raw_text)
        except Exception:
            self.schema_dict = None

    def is_valid(self) -> bool:
        """Check if parsed content contains a valid TMSL model."""
        if not isinstance(self.schema_dict, dict):
            return False
        # TMSL typically has 'model' or is the model object itself
        return "model" in self.schema_dict or "tables" in self.schema_dict

    def get_model(self) -> Dict[str, Any]:
        """Retrieve root model dict."""
        if not self.schema_dict:
            return {}
        if "model" in self.schema_dict and isinstance(self.schema_dict["model"], dict):
            return self.schema_dict["model"]
        return self.schema_dict

    def extract_artifacts(self) -> Dict[str, Any]:
        """
        Deconstruct DataModelSchema into structured artifacts:
          - 'formatted_schema_json': pretty-printed full schema JSON
          - 'tables': dict of {table_name: table_metadata_dict}
          - 'measures': dict of {measure_key: {'table': str, 'name': str, 'expression': str, 'description': str}}
          - 'relationships': list of model relationship definitions
          - 'm_partitions': dict of {partition_name: m_query_code}
          - 'roles': list of security roles
        """
        if not self.is_valid():
            return {}

        model = self.get_model()
        results: Dict[str, Any] = {
            "formatted_schema_json": json.dumps(self.schema_dict, indent=2),
            "tables": {},
            "measures": {},
            "relationships": model.get("relationships", []),
            "m_partitions": {},
            "roles": model.get("roles", []),
        }

        tables = model.get("tables", [])
        for tbl in tables:
            t_name = tbl.get("name", "UnnamedTable")
            results["tables"][t_name] = tbl

            # Extract measures
            for m in tbl.get("measures", []):
                m_name = m.get("name", "UnnamedMeasure")
                expr = m.get("expression", "")
                if isinstance(expr, list):
                    expr_text = "\n".join(expr)
                else:
                    expr_text = str(expr)

                m_key = f"{t_name}_{m_name}"
                results["measures"][m_key] = {
                    "table": t_name,
                    "name": m_name,
                    "expression": expr_text,
                    "format_string": m.get("formatString", ""),
                    "description": m.get("description", ""),
                }

            # Extract partitions & M expressions
            for part in tbl.get("partitions", []):
                p_name = part.get("name", f"{t_name}_Partition")
                src = part.get("source", {})
                if isinstance(src, dict) and src.get("type", "").lower() == "m":
                    m_expr = src.get("expression", "")
                    if isinstance(m_expr, list):
                        m_code = "\n".join(m_expr)
                    else:
                        m_code = str(m_expr)
                    results["m_partitions"][f"{t_name}_{p_name}"] = m_code

        return results
