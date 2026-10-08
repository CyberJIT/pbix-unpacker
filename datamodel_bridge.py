"""
DataModel & VertiPaq ABF Bridge Module.
Provides lightweight automated utilities to inspect, extract, and bridge
the proprietary VertiPaq Analysis Services Backup File (ABF) found in .pbix files:
1. Header analysis & metadata extraction (XPress9 / Analysis Services ABF format).
2. Standalone ABF export & XMLA restore script generation (for SSAS developer instances).
3. Local Power BI Desktop SSAS instance discovery (detect active msmdsrv.port.txt).
4. Automated TMDL / DAX Query View script extraction if embedded in modern PBIX files.
"""

import os
import sys
import glob
import re
from typing import Dict, Any, List, Optional, Tuple


def inspect_abf_header(data: bytes) -> Dict[str, Any]:
    """
    Inspect raw binary DataModel bytes to identify the ABF VertiPaq container header.
    """
    res = {
        "is_abf": False,
        "compression": "Unknown",
        "header_text": None,
        "size_bytes": len(data),
    }
    if not data or len(data) < 16:
        return res

    # Check for UTF-16LE preamble commonly found in SSAS ABF files:
    # "This backup was created using XPress9 compression."
    prefix = data[:256]
    try:
        decoded_utf16 = prefix.decode("utf-16le", errors="ignore")
        if "backup" in decoded_utf16.lower() or "xpress" in decoded_utf16.lower():
            res["is_abf"] = True
            clean_str = "".join(ch for ch in decoded_utf16 if ch.isprintable() or ch in " \r\n\t").strip()
            res["header_text"] = clean_str
            if "xpress9" in clean_str.lower():
                res["compression"] = "XPress9"
            elif "xpress" in clean_str.lower():
                res["compression"] = "XPress"
            return res
    except Exception:
        pass

    # Check UTF-8 / ASCII fallback
    decoded_ascii = "".join(chr(b) if 32 <= b < 127 else " " for b in prefix)
    if "backup" in decoded_ascii.lower() or "xpress" in decoded_ascii.lower():
        res["is_abf"] = True
        res["header_text"] = decoded_ascii.strip()

    return res


def generate_xmla_restore_script(
    database_name: str,
    abf_file_path: str,
    allow_overwrite: bool = True
) -> str:
    """
    Generate standard XMLA RESTORE script targeting a local SSAS (Tabular) instance.
    """
    allow_ow_str = "true" if allow_overwrite else "false"
    xmla = f"""<Restore xmlns="http://schemas.microsoft.com/analysisservices/2003/engine">
  <File>{abf_file_path}</File>
  <DatabaseName>{database_name}</DatabaseName>
  <AllowOverwrite>{allow_ow_str}</AllowOverwrite>
</Restore>"""
    return xmla


def find_active_powerbi_ssas_instances() -> List[Dict[str, Any]]:
    """
    Detect active Power BI Desktop Analysis Services instances running on the local machine
    by scanning the user's Microsoft Power BI Desktop AnalysisServicesWorkspace directories.
    Works natively on Windows and wine/VM paths.
    """
    instances = []
    candidates = []

    # Windows standard paths
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(os.path.join(local_app_data, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspace"))
    user_profile = os.environ.get("USERPROFILE")
    if user_profile:
        candidates.append(os.path.join(user_profile, "AppData", "Local", "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspace"))

    for base_dir in candidates:
        if os.path.isdir(base_dir):
            for root, dirs, files in os.walk(base_dir):
                if "msmdsrv.port.txt" in files:
                    port_file = os.path.join(root, "msmdsrv.port.txt")
                    try:
                        with open(port_file, "r", encoding="utf-16le", errors="ignore") as pf:
                            content = pf.read().strip()
                        if not content:
                            with open(port_file, "r", encoding="utf-8", errors="ignore") as pf:
                                content = pf.read().strip()
                        m = re.search(r"(\d+)", content)
                        if m:
                            port = int(m.group(1))
                            instances.append({
                                "port": port,
                                "port_file": port_file,
                                "workspace_dir": root,
                                "connection_string": f"Provider=MSOLAP;Data Source=localhost:{port};"
                            })
                    except Exception:
                        pass
    return instances


def extract_datamodel_bridge(
    target_dir: str,
    datamodel_bytes: bytes,
    base_name: str,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    Extract DataModel bridge artifacts:
    1. DataModel.abf (Standalone SSAS backup file ready for XMLA restore).
    2. DataModel_Restore.xmla (Ready-to-execute XMLA restore script).
    3. DataModel_Info.json (Technical metadata, compression mode, and connection guidance).
    """
    info = inspect_abf_header(datamodel_bytes)
    safe_db_name = re.sub(r"[^\w\-_]", "_", base_name).strip() or "PowerBI_Model"

    abf_rel = "DataModel_Bridge/DataModel.abf"
    xmla_rel = "DataModel_Bridge/Restore_Database.xmla"
    info_rel = "DataModel_Bridge/DataModel_Info.json"

    # XMLA restore script uses relative or placeholder path
    xmla_content = generate_xmla_restore_script(
        database_name=safe_db_name,
        abf_file_path=os.path.abspath(os.path.join(target_dir, abf_rel))
    )

    info_summary = {
        "file_name": "DataModel",
        "file_size_bytes": len(datamodel_bytes),
        "is_analysis_services_backup": info["is_abf"],
        "compression_type": info["compression"],
        "header_descriptor": info["header_text"],
        "recommended_workflows": {
            "workflow_1_powerbi_live_connection": {
                "description": "Open the .pbix file in Power BI Desktop, locate msmdsrv.port.txt in AppData, and connect DAX Studio or Tabular Editor to localhost:<port>.",
                "powerbi_port_detection": "Run pbix_unpacker with --find-active-ports"
            },
            "workflow_2_pbip_export": {
                "description": "In Power BI Desktop, click File -> Save As -> Power BI Project (.pbip). This extracts TMDL / model.bim into git-friendly text files."
            },
            "workflow_3_ssas_restore": {
                "description": "Restore DataModel.abf into a local SSAS Tabular instance using Restore_Database.xmla.",
                "target_database_name": safe_db_name,
                "restore_script": xmla_rel
            }
        }
    }

    return {
        "abf_rel": abf_rel,
        "abf_bytes": datamodel_bytes,
        "xmla_rel": xmla_rel,
        "xmla_text": xmla_content,
        "info_rel": info_rel,
        "info_dict": info_summary
    }
