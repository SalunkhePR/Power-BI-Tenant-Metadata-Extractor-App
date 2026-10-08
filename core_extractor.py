# =============================================================================
#  Power BI Complete Metadata Extractor
#  Silent, client-ready, no console flashing
# =============================================================================

import subprocess
import time
import re
import os
import sys
import tempfile
import json
from pathlib import Path
from typing import Callable, Optional

import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

SCAN_CHUNK_SIZE = 100
SCAN_POLL_SECS = 4
CREATE_NO_WINDOW = 0x08000000  # Hide console windows on Windows


def get_default_output_dir() -> Path:
    home = Path.home()
    for candidate in [
        home / "Desktop" / "PowerBI_Metadata",
        home / "Documents" / "PowerBI_Metadata",
        home / "PowerBI_Metadata",
    ]:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        except Exception:
            continue
    fallback = Path.cwd() / "PowerBI_Metadata_Output"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def _silent_subprocess(cmd: list, timeout: int = 120) -> subprocess.CompletedProcess:
    """Run subprocess with no visible console window on Windows."""
    kwargs = {}
    if os.name == "nt":
        kwargs["creationflags"] = CREATE_NO_WINDOW
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
        kwargs["startupinfo"] = startupinfo
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, **kwargs)


def build_param_map(expressions_list: list) -> dict:
    params = {}
    for expr in expressions_list:
        name = expr.get("name", "")
        m_code = expr.get("expression", "")
        if not name or not m_code:
            continue
        m1 = re.search(r'"([^"]+)"\s*meta\s*\[', m_code)
        if m1:
            params[name] = m1.group(1)
            continue
        m2 = re.search(r'^\s*"([^"]+)"\s*$', m_code)
        if m2:
            params[name] = m2.group(1)
            continue
        m3 = re.search(r'=\s*"([^"]+)"', m_code)
        if m3:
            params[name] = m3.group(1)
    return params


def resolve_params(m_expr: str, param_map: dict) -> str:
    resolved = m_expr
    for p_name, p_val in param_map.items():
        pattern = r'(?<!")\b' + re.escape(p_name) + r'\b(?!")'
        resolved = re.sub(pattern, f'"{p_val}"', resolved)
    return resolved


def extract_connection_from_m(m_expr: str, param_map: dict = None) -> tuple:
    if not m_expr or not m_expr.strip():
        return "Unknown", "N/A"
    m = m_expr.strip()
    if param_map:
        m = resolve_params(m, param_map)

    patterns = [
        (r'GoogleSheets\.Contents\s*\(\s*"([^"]+)"', "Google Sheets", lambda g: g[0]),
        (r'Teradata\.Database\s*\(\s*"([^"]+)"', "Teradata", lambda g: f"Server: {g[0]}"),
        (r'Snowflake\.Databases?\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', "Snowflake",
         lambda g: f"Account: {g[0]} | Warehouse: {g[1]}"),
        (r'Sql\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', "SQL Server / Azure SQL",
         lambda g: f"Server: {g[0]} | Database: {g[1]}"),
        (r'Sql\.Database\s*\(\s*(\w+)\s*,\s*(\w+)', "SQL Server / Azure SQL (via param)",
         lambda g: f"Server param: {g[0]} | DB param: {g[1]}"),
        (r'Oracle\.Database\s*\(\s*"([^"]+)"', "Oracle", lambda g: f"Server: {g[0]}"),
        (r'MySQL\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', "MySQL",
         lambda g: f"Server: {g[0]} | Database: {g[1]}"),
        (r'PostgreSQL\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', "PostgreSQL",
         lambda g: f"Server: {g[0]} | Database: {g[1]}"),
        (r'Odbc\.(?:Query|Tables|DataSource)\s*\(\s*"([^"]+)"', "ODBC", lambda g: g[0]),
        (r'Web\.(?:Contents|Page|BrowserContents)\s*\(\s*"([^"]+)"', "Web / REST API", lambda g: g[0]),
        (r'SharePoint\.(?:Files|Tables|Contents)\s*\(\s*"([^"]+)"', "SharePoint", lambda g: g[0]),
        (r'AzureStorage\.(?:Blobs|Tables|BlobContents|DataLake)\s*\(\s*"([^"]+)"', "Azure Blob / ADLS", lambda g: g[0]),
        (r'Databricks\.Catalogs\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', "Databricks",
         lambda g: f"Host: {g[0]} | HTTP Path: {g[1]}"),
        (r'AzureDataExplorer\.Contents\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', "Azure Data Explorer / Kusto",
         lambda g: f"Cluster: {g[0]} | Database: {g[1]}"),
        (r'File\.Contents\s*\(\s*"([^"]+\.xlsx?)"', "Excel File", lambda g: g[0]),
        (r'File\.Contents\s*\(\s*"([^"]+\.csv)"', "CSV File", lambda g: g[0]),
        (r'File\.Contents\s*\(\s*"([^"]+)"', "Local / Network File", lambda g: g[0]),
        (r'AdminInsights\.GetAzureBlobContents\s*\(\s*"([^"]+)"', "Azure Blob / Admin Insights",
         lambda g: f"Container: {g[0]}"),
        (r'"([^"]*\.datawarehouse\.(?:fabric|pbidedicated)\.(?:microsoft\.com|windows\.net)[^"]*)"',
         "Microsoft Fabric / Synapse Warehouse", lambda g: g[0]),
    ]
    for pattern, name, formatter in patterns:
        match = re.search(pattern, m, re.IGNORECASE)
        if match:
            return name, formatter(match.groups())

    if "GoogleBigQuery.Database" in m:
        proj = re.search(r'\{?\[Name\s*=\s*"([^"]+)"', m)
        return "Google BigQuery", f"Project: {proj.group(1)}" if proj else "Google BigQuery"

    for kw in ["CALENDAR(", "CALENDARAUTO(", "GENERATESERIES(", "ADDCOLUMNS(", "NAMEOF(", "Row("]:
        if kw in m:
            return "DAX Calculated Table", "N/A (computed in-memory)"
    if "Binary.Decompress" in m or "Table.FromRows" in m:
        return "Embedded / Inline Data", "N/A (hardcoded rows in M)"
    if "Table.Combine" in m:
        refs = re.findall(r'\{(\w[^}]*)\}', m)
        return "Combined Tables", f"Union of: {', '.join(refs[:5])}"
    if re.match(r'^\w+\s*$', m):
        return "Fabric / Warehouse Table Ref", f"Table name: {m.strip()}"
    return "Other / Unknown", m[:150].replace("\n", " ")


def extract_nested_joins(m_expr: str) -> list:
    results = []
    p1 = re.compile(
        r'Table\.NestedJoin\s*\(\s*[^,]+,\s*\{([^}]+)\}\s*,\s*(\w+)\s*,\s*\{([^}]+)\}'
        r'(?:\s*,\s*"[^"]*")?(?:\s*,\s*(JoinKind\.\w+))?', re.IGNORECASE)
    for match in p1.findall(m_expr):
        results.append({
            "from_col": match[0].strip().strip('"'),
            "to_table": match[1].strip(),
            "to_col": match[2].strip().strip('"'),
            "join_kind": match[3] if match[3] else "N/A",
            "join_func": "Table.NestedJoin",
        })
    p2 = re.compile(
        r'Table\.NestedJoin\s*\(\s*[^,]+,\s*"([^"]+)"\s*,\s*(\w+)\s*,\s*"([^"]+)"'
        r'(?:\s*,\s*"[^"]*")?(?:\s*,\s*(JoinKind\.\w+))?', re.IGNORECASE)
    for match in p2.findall(m_expr):
        if not any(r["from_col"] == match[0].strip('"') and r["to_table"] == match[1] for r in results):
            results.append({
                "from_col": match[0].strip().strip('"'),
                "to_table": match[1].strip(),
                "to_col": match[2].strip().strip('"'),
                "join_kind": match[3] if match[3] else "N/A",
                "join_func": "Table.NestedJoin",
            })
    p3 = re.compile(
        r'Table\.Join\s*\(\s*(\w+)\s*,\s*"([^"]+)"\s*,\s*(\w+)\s*,\s*"([^"]+)"'
        r'(?:\s*,\s*(JoinKind\.\w+))?', re.IGNORECASE)
    for match in p3.findall(m_expr):
        results.append({
            "from_col": match[1].strip(),
            "to_table": match[2].strip(),
            "to_col": match[3].strip(),
            "join_kind": match[4] if match[4] else "N/A",
            "join_func": "Table.Join",
        })
    return results


def fetch_model_relationships_via_dax(headers, ws_id, ws_name, ds_id, ds_name) -> list:
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{ds_id}/executeQueries"
    payload = {"queries": [{"query": "EVALUATE INFO.VIEW.RELATIONSHIPS()"}],
               "serializerSettings": {"includeNulls": True}}
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=12)
        if resp.status_code == 200:
            tables = resp.json().get("results", [{}])[0].get("tables", [])
            if tables:
                rels = []
                for r in tables[0].get("rows", []):
                    clean = {k.strip("[]"): v for k, v in r.items()}
                    from_card = clean.get("FromCardinality", "")
                    to_card = clean.get("ToCardinality", "")
                    card = f"{from_card}To{to_card}" if (from_card and to_card) else clean.get("Cardinality", "N/A")
                    rels.append({
                        "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                        "Relationship Name": clean.get("Name", "AutoDetected"),
                        "From Table (Fact/Many)": clean.get("FromTable", "N/A"),
                        "From Column": clean.get("FromColumn", "N/A"),
                        "To Table (Dim/One)": clean.get("ToTable", "N/A"),
                        "To Column": clean.get("ToColumn", "N/A"),
                        "Cross Filter Direction": clean.get("CrossFilteringBehavior", "OneDirection"),
                        "Security Filter Behavior": clean.get("SecurityFilteringBehavior", "OneDirection"),
                        "Cardinality": card,
                        "Is Active": clean.get("IsActive", True),
                        "Rely On Ref Integrity": clean.get("RelyOnReferentialIntegrity", False),
                        "Discovery Source": "DAX executeQueries (INFO.VIEW.RELATIONSHIPS)",
                    })
                return rels
    except Exception:
        pass
    return []


def run_extraction(
    output_dir: Optional[str | Path] = None,
    open_excel_when_done: bool = False,
    log_callback: Optional[Callable[[str], None]] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
) -> dict:
    """Run full extraction. Returns dict with success, paths, message, row_counts."""

    def log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    def progress(pct: float, msg: str = ""):
        if progress_callback:
            progress_callback(pct, msg)

    try:
        out_path = Path(output_dir) if output_dir else get_default_output_dir()
        out_path.mkdir(parents=True, exist_ok=True)
        excel_path = out_path / "PowerBI_Complete_Architecture_Map_Output.xlsx"
        json_path = out_path / "PowerBI_Raw_Metadata.json"

        log("Starting Power BI metadata extraction…")
        log(f"Output folder: {out_path}")
        log("")

        # ---- Authentication (completely silent) ----
        progress(0.03, "Signing in to Microsoft…")
        log("Opening Microsoft sign-in window…")

        temp_token_file = os.path.join(tempfile.gettempdir(), "pbi_token_temp.txt")
        ps_token_path = temp_token_file.replace("'", "''")

        ps_script = f"""
$ErrorActionPreference = 'Stop'
try {{
    Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop
    Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue
    Connect-PowerBIServiceAccount | Out-Null
    $token = Get-PowerBIAccessToken -AsString
    [System.IO.File]::WriteAllText('{ps_token_path}', $token, [System.Text.Encoding]::ASCII)
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
"""
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=180,
            )
        except subprocess.TimeoutExpired:
            raise Exception(
                "Sign-in timed out.\n\nPlease try again and complete the Microsoft login within 3 minutes."
            )

        if not os.path.exists(temp_token_file):
            err = (result.stderr or result.stdout or "").strip().lower()
            module_keywords = [
                "microsoftpowerbimgmt", "not recognized", "could not be loaded",
                "module", "not found", "was not loaded", "cannot find",
                "no valid module", "commandnotfoundexception"
            ]
            if any(k in err for k in module_keywords):
                raise Exception(
                    "Required PowerShell module is missing or cannot be loaded.\n\n"
                    "Please open PowerShell and run these commands:\n\n"
                    "Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force\n"
                    "Import-Module MicrosoftPowerBIMgmt\n\n"
                    "Then close and restart this application."
                )
            raise Exception(
                "Microsoft sign-in was cancelled or failed.\n\n"
                "Please try again and complete the login popup.\n\n"
                "If the login window did not appear, the PowerShell module may need to be reinstalled."
            )

        with open(temp_token_file, "r", encoding="utf-8-sig") as f:
            TOKEN = f.read().strip()
        try:
            os.remove(temp_token_file)
        except Exception:
            pass

        if TOKEN and not TOKEN.lower().startswith("bearer"):
            TOKEN = "Bearer " + TOKEN
        if not TOKEN:
            raise Exception("Failed to obtain a valid access token. Please try signing in again.")

        log("Signed in successfully.")
        log("")
        HEADERS = {"Content-Type": "application/json", "Authorization": TOKEN}

        # ---- Data stores ----
        workspaces_data, reports_data, dashboards_data, tiles_data = [], [], [], []
        datasets_data, datasources_data, columns_data, measures_data = [], [], [], []
        expressions_data, roles_data, relationships_data = [], [], []
        pq_joins_data, artifact_lineage_data = [], []
        datasource_instances_map = {}

        # ---- Phase 1: Admin Scan ----
        progress(0.08, "Fetching workspaces…")
        log("Fetching workspace list…")

        try:
            ws_resp = requests.get(
                "https://api.powerbi.com/v1.0/myorg/admin/groups?$top=5000",
                headers=HEADERS, timeout=60
            )
        except requests.RequestException as e:
            raise Exception(f"Network error while contacting Power BI service:\n{e}")

        if ws_resp.status_code in (401, 403):
            raise Exception(
                "Access denied.\n\n"
                "Your account does not have Power BI Administrator rights,\n"
                "or the required tenant setting is disabled.\n\n"
                "Required:\n"
                "• Power BI Administrator role\n"
                "• Tenant setting “Enhance admin API responses with detailed metadata” = Enabled"
            )
        if ws_resp.status_code != 200:
            raise Exception(
                f"Unable to retrieve workspaces (HTTP {ws_resp.status_code}).\n\n"
                "Please verify your Power BI Admin permissions and try again."
            )

        all_workspaces_list = ws_resp.json().get("value", [])
        WORKSPACE_IDS = [ws["id"] for ws in all_workspaces_list if ws.get("type") != "PersonalGroup"]
        log(f"Found {len(WORKSPACE_IDS)} workspaces.")
        log("")

        if not WORKSPACE_IDS:
            raise Exception("No accessible workspaces found.\n\nPlease check that your account has access to at least one workspace.")

        chunks = [WORKSPACE_IDS[i:i + SCAN_CHUNK_SIZE] for i in range(0, len(WORKSPACE_IDS), SCAN_CHUNK_SIZE)]
        SCAN_BASE = "https://api.powerbi.com/v1.0/myorg/admin/workspaces"
        all_raw_metadata = []
        total_chunks = len(chunks)

        for idx, chunk in enumerate(chunks, 1):
            pct = 0.10 + (0.45 * (idx - 1) / max(total_chunks, 1))
            progress(pct, f"Scanning batch {idx} of {total_chunks}…")
            log(f"Scanning batch {idx} of {total_chunks} ({len(chunk)} workspaces)…")

            trigger_url = (
                f"{SCAN_BASE}/getInfo?lineage=True&datasourceDetails=True"
                "&datasetSchema=True&datasetExpressions=True"
            )
            try:
                t_resp = requests.post(trigger_url, headers=HEADERS, json={"workspaces": chunk}, timeout=60)
                t_resp.raise_for_status()
                scan_id = t_resp.json().get("id")
            except Exception as e:
                raise Exception(f"Failed to start scan for batch {idx}:\n{e}")

            while True:
                try:
                    status_resp = requests.get(f"{SCAN_BASE}/scanStatus/{scan_id}", headers=HEADERS, timeout=30)
                    status = status_resp.json().get("status")
                except Exception:
                    time.sleep(SCAN_POLL_SECS)
                    continue
                if status == "Succeeded":
                    break
                if status == "Failed":
                    raise Exception(f"Scan failed for batch {idx}. Please try again later.")
                time.sleep(SCAN_POLL_SECS)

            try:
                r_resp = requests.get(f"{SCAN_BASE}/scanResult/{scan_id}", headers=HEADERS, timeout=120)
                r_resp.raise_for_status()
                scan_result_json = r_resp.json()
            except Exception as e:
                raise Exception(f"Failed to retrieve scan results for batch {idx}:\n{e}")

            batch_data = scan_result_json.get("workspaces", [])
            all_raw_metadata.extend(batch_data)

            for d_inst in scan_result_json.get("datasourceInstances", []):
                inst_id = d_inst.get("datasourceId") or d_inst.get("datasourceInstanceId")
                d_type = d_inst.get("datasourceType", "Unknown")
                cd = d_inst.get("connectionDetails", {})
                c_str = "; ".join(f"{k}={v}" for k, v in cd.items()) if cd else "N/A"
                if inst_id:
                    datasource_instances_map[inst_id] = (d_type, c_str)

            log(f"  → Retrieved {len(batch_data)} workspaces.")

        log("")
        progress(0.58, "Parsing metadata…")
        log("Parsing metadata, queries and relationships…")

        # ---- Parse metadata ----
        for ws in all_raw_metadata:
            ws_id = ws.get("id", "N/A")
            ws_name = ws.get("name", "Unknown Workspace")
            workspaces_data.append({
                "Workspace ID": ws_id, "Workspace Name": ws_name,
                "Workspace Type": ws.get("type", "N/A"), "Workspace State": ws.get("state", "N/A"),
            })

            for rep in ws.get("reports", []):
                reports_data.append({
                    "Workspace Name": ws_name, "Report Name": rep.get("name", "N/A"),
                    "Report ID": rep.get("id", "N/A"), "Report Type": rep.get("reportType", "N/A"),
                    "Associated Dataset ID": rep.get("datasetId", "N/A"),
                    "Created By": rep.get("createdBy", "N/A"), "Modified By": rep.get("modifiedBy", "N/A"),
                    "Created Date": rep.get("createdDateTime", "N/A"),
                    "Modified Date": rep.get("modifiedDateTime", "N/A"),
                    "Endorsement": rep.get("endorsementDetails", {}).get("endorsement", "N/A"),
                })

            for dash in ws.get("dashboards", []):
                dash_id = dash.get("id", "N/A")
                dash_name = dash.get("displayName", "N/A")
                tiles = dash.get("tiles", [])
                dashboards_data.append({
                    "Workspace Name": ws_name, "Dashboard Name": dash_name,
                    "Dashboard ID": dash_id, "Is Read Only": dash.get("isReadOnly", False),
                    "Total Tiles": len(tiles),
                })
                for tile_num, tile in enumerate(tiles, 1):
                    tiles_data.append({
                        "Workspace Name": ws_name, "Dashboard Name": dash_name, "Dashboard ID": dash_id,
                        "Tile #": tile_num, "Tile ID": tile.get("id", "N/A"),
                        "Tile Title": tile.get("title", "(No title)"),
                        "Tile Sub-Title": tile.get("subTitle", ""),
                        "Pinned From Report ID": tile.get("reportId", ""),
                        "Associated Dataset ID": tile.get("datasetId", "N/A"),
                    })

            for ds in ws.get("datasets", []):
                ds_id = ds.get("id", "N/A")
                ds_name = ds.get("name", "Unknown Dataset")
                datasets_data.append({
                    "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                    "Content Type": ds.get("contentProviderType", "N/A"),
                    "Storage Mode": ds.get("targetStorageMode", "N/A"),
                    "Configured By": ds.get("configuredBy", "N/A"),
                    "Created Date": ds.get("createdDate", "N/A"),
                    "Schema Stale?": ds.get("schemaMayNotBeUpToDate", False),
                    "Schema Error": ds.get("schemaRetrievalError", "None") or "None",
                })

                param_map = build_param_map(ds.get("expressions", []))
                seen_connections = set()

                for table in ds.get("tables", []):
                    tb_name = table.get("name", "N/A")
                    for col in table.get("columns", []):
                        col_expr = col.get("expression", "")
                        col_type = col.get("columnType", "Data")
                        if not col_expr:
                            if "calculatedtable" in str(col_type).lower():
                                calc_formula = "N/A (Calculated Table Column)"
                            elif "calculated" in str(col_type).lower():
                                calc_formula = "N/A (Calculated without expression)"
                            else:
                                calc_formula = "N/A (Direct Source / Imported)"
                        else:
                            calc_formula = col_expr
                        columns_data.append({
                            "Workspace Name": ws_name, "Dataset Name": ds_name, "Table Name": tb_name,
                            "Column Name": col.get("name", "N/A"), "Data Type": col.get("dataType", "N/A"),
                            "Column Type": col_type, "Calculated Formula (DAX)": calc_formula,
                            "Is Hidden": col.get("isHidden", False),
                        })

                    for meas in table.get("measures", []):
                        measures_data.append({
                            "Workspace Name": ws_name, "Dataset Name": ds_name, "Table Name": tb_name,
                            "Measure Name": meas.get("name", "N/A"),
                            "DAX Expression": meas.get("expression", "N/A"),
                            "Is Hidden": meas.get("isHidden", False),
                        })

                    for src in table.get("source", []):
                        m_expr = src.get("expression", "")
                        if not m_expr:
                            continue
                        expressions_data.append({
                            "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                            "Query / Object Name": tb_name, "Query Type": "Table Query (Advanced Editor)",
                            "M Expression": m_expr,
                            "Description": table.get("description", "Table query in Advanced Editor"),
                            "Is Parameter?": False, "Parameter Value": "N/A",
                        })
                        src_type, conn_str = extract_connection_from_m(m_expr, param_map)
                        dedup_key = f"{src_type}|{conn_str}"
                        if (src_type not in ("DAX Calculated Table", "Embedded / Inline Data", "Combined Tables")
                                and dedup_key not in seen_connections):
                            seen_connections.add(dedup_key)
                            datasources_data.append({
                                "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                                "Table Name": tb_name, "Source Scope": "Table",
                                "Source Type": src_type, "Connection String": conn_str,
                                "Extraction Method": "Power Query M",
                            })
                        for join in extract_nested_joins(m_expr):
                            pq_joins_data.append({
                                "Workspace Name": ws_name, "Dataset Name": ds_name,
                                "Host Table (From)": tb_name, "Join From Column": join["from_col"],
                                "Joined Table (To)": join["to_table"], "Join To Column": join["to_col"],
                                "Join Kind": join.get("join_kind", "N/A"),
                                "Join Function": join.get("join_func", "Table.NestedJoin"),
                            })

                if not seen_connections:
                    for usage in ds.get("datasourceUsages", []):
                        u_id = usage.get("datasourceInstanceId", "")
                        if u_id in datasource_instances_map:
                            inst_type, inst_conn = datasource_instances_map[u_id]
                            datasources_data.append({
                                "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                                "Table Name": "(Dataset Level / All Tables)", "Source Scope": "Dataset",
                                "Source Type": inst_type, "Connection String": inst_conn,
                                "Extraction Method": "Scanner Datasource Instance",
                            })

                dax_rels = fetch_model_relationships_via_dax(HEADERS, ws_id, ws_name, ds_id, ds_name)
                if dax_rels:
                    relationships_data.extend(dax_rels)
                else:
                    for rel in ds.get("relationships", []):
                        relationships_data.append({
                            "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                            "Relationship Name": rel.get("name", "AutoDetected"),
                            "From Table (Fact/Many)": rel.get("fromTable", "N/A"),
                            "From Column": rel.get("fromColumn", "N/A"),
                            "To Table (Dim/One)": rel.get("toTable", "N/A"),
                            "To Column": rel.get("toColumn", "N/A"),
                            "Cross Filter Direction": rel.get("crossFilteringBehavior", "OneDirection"),
                            "Security Filter Behavior": rel.get("securityFilteringBehavior", "OneDirection"),
                            "Cardinality": rel.get("cardinality", "ManyToOne"),
                            "Is Active": rel.get("isActive", True),
                            "Rely On Ref Integrity": rel.get("relyOnReferentialIntegrity", False),
                            "Discovery Source": "Scanner API Schema",
                        })

                for expr in ds.get("expressions", []):
                    name = expr.get("name", "N/A")
                    m_code = expr.get("expression", "N/A")
                    is_param = bool(re.search(r'IsParameterQuery\s*=\s*true', m_code, re.I))
                    expressions_data.append({
                        "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                        "Query / Object Name": name,
                        "Query Type": "Power Query Parameter" if is_param else "Shared Query / Custom Function",
                        "M Expression": m_code,
                        "Description": expr.get("description", "Shared query or parameter"),
                        "Is Parameter?": is_param,
                        "Parameter Value": param_map.get(name, "N/A") if is_param else "N/A",
                    })

                for role in ds.get("roles", []):
                    role_name = role.get("name", "N/A")
                    permission = role.get("modelPermission", "N/A")
                    table_filters = "; ".join(
                        f"{tp.get('name')}[{tp.get('filterExpression', 'N/A')}]"
                        for tp in role.get("tablePermissions", [])
                    )
                    members = role.get("members", [])
                    if members:
                        for member in members:
                            roles_data.append({
                                "Workspace Name": ws_name, "Dataset Name": ds_name,
                                "Role Name": role_name, "Permission": permission,
                                "Table Filters": table_filters,
                                "Member Name": member.get("memberName", "N/A"),
                                "Member Type": member.get("memberType", "N/A"),
                                "Identity Provider": member.get("identityProvider", "N/A"),
                            })
                    else:
                        roles_data.append({
                            "Workspace Name": ws_name, "Dataset Name": ds_name,
                            "Role Name": role_name, "Permission": permission,
                            "Table Filters": table_filters,
                            "Member Name": "(No members assigned)", "Member Type": "N/A",
                            "Identity Provider": "N/A",
                        })

                for rel in ds.get("relations", []):
                    artifact_lineage_data.append({
                        "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                        "Dependent On Artifact ID": rel.get("dependentOnArtifactId", "N/A"),
                        "Dependent Workspace ID": rel.get("workspaceId", "N/A"),
                        "Relation Type": rel.get("relationType", "N/A"),
                        "Settings": rel.get("settingsList", "N/A"),
                    })

        log("Metadata parsing complete.")
        log("")

        # ---- Phase 2 ----
        progress(0.75, "Checking additional data sources…")
        log("Checking supplemental data sources…")
        try:
            std_resp = requests.get("https://api.powerbi.com/v1.0/myorg/groups?$top=5000", headers=HEADERS, timeout=60)
            if std_resp.status_code == 200:
                for ws in std_resp.json().get("value", []):
                    ws_id, ws_name = ws.get("id", "N/A"), ws.get("name", "Unknown")
                    try:
                        ds_resp = requests.get(f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets",
                                               headers=HEADERS, timeout=30)
                        if ds_resp.status_code != 200:
                            continue
                        for ds in ds_resp.json().get("value", []):
                            ds_id, ds_name = ds.get("id", "N/A"), ds.get("name", "N/A")
                            try:
                                src_resp = requests.get(
                                    f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{ds_id}/datasources",
                                    headers=HEADERS, timeout=30)
                                if src_resp.status_code != 200:
                                    continue
                                for src in src_resp.json().get("value", []):
                                    conn = src.get("connectionString", "")
                                    if not conn:
                                        cd = src.get("connectionDetails", {})
                                        conn = "; ".join(f"{k}={v}" for k, v in cd.items()) if cd else "N/A"
                                    datasources_data.append({
                                        "Workspace Name": ws_name, "Dataset Name": ds_name, "Dataset ID": ds_id,
                                        "Table Name": "(Dataset Level / All Tables)", "Source Scope": "Dataset",
                                        "Source Type": src.get("datasourceType", "N/A"),
                                        "Connection String": conn,
                                        "Extraction Method": "Power BI Service Standard API",
                                    })
                            except Exception:
                                continue
                    except Exception:
                        continue
                log("Supplemental check complete.")
            else:
                log("Supplemental check skipped.")
        except Exception:
            log("Supplemental check skipped.")
        log("")

        # ---- Compile DataFrames ----
        progress(0.85, "Preparing Excel report…")
        log("Compiling results…")

        def safe_df(data_list, cols):
            return pd.DataFrame(data_list) if data_list else pd.DataFrame(columns=cols)

        df_workspaces = safe_df(workspaces_data, ["Workspace ID", "Workspace Name", "Workspace Type", "Workspace State"])
        df_reports = safe_df(reports_data, ["Workspace Name", "Report Name", "Report ID", "Report Type",
                                            "Associated Dataset ID", "Created By", "Modified By",
                                            "Created Date", "Modified Date", "Endorsement"])
        df_dashboards = safe_df(dashboards_data, ["Workspace Name", "Dashboard Name", "Dashboard ID", "Is Read Only", "Total Tiles"])
        df_tiles = safe_df(tiles_data, ["Workspace Name", "Dashboard Name", "Dashboard ID", "Tile #", "Tile ID",
                                        "Tile Title", "Tile Sub-Title", "Pinned From Report ID", "Associated Dataset ID"])
        df_datasets = safe_df(datasets_data, ["Workspace Name", "Dataset Name", "Dataset ID", "Content Type",
                                              "Storage Mode", "Configured By", "Created Date", "Schema Stale?", "Schema Error"])
        df_datasources = safe_df(datasources_data, ["Workspace Name", "Dataset Name", "Dataset ID", "Table Name",
                                                    "Source Scope", "Source Type", "Connection String", "Extraction Method"])

        if not df_datasources.empty:
            def norm_conn(s):
                return re.sub(r'[\s:;=_/|\-]', '', str(s).lower())
            df_datasources["_conn_norm"] = df_datasources["Connection String"].apply(norm_conn)
            table_covered = set(df_datasources[df_datasources["Source Scope"] == "Table"]["_conn_norm"])
            cleaned = [row.to_dict() for _, row in df_datasources.iterrows()
                       if not (row["Source Scope"] == "Dataset" and row["_conn_norm"] in table_covered)]
            df_datasources = pd.DataFrame(cleaned).drop(columns=["_conn_norm"], errors="ignore")
            df_datasources.drop_duplicates(subset=["Workspace Name", "Dataset Name", "Table Name", "Connection String"], inplace=True)
            df_datasources.sort_values(by=["Workspace Name", "Dataset Name", "Table Name"], inplace=True)

        df_columns = safe_df(columns_data, ["Workspace Name", "Dataset Name", "Table Name", "Column Name",
                                            "Data Type", "Column Type", "Calculated Formula (DAX)", "Is Hidden"])
        df_measures = safe_df(measures_data, ["Workspace Name", "Dataset Name", "Table Name", "Measure Name",
                                              "DAX Expression", "Is Hidden"])
        df_expressions = safe_df(expressions_data, ["Workspace Name", "Dataset Name", "Dataset ID", "Query / Object Name",
                                                    "Query Type", "M Expression", "Description", "Is Parameter?", "Parameter Value"])
        if not df_expressions.empty:
            df_expressions.drop_duplicates(inplace=True)
            df_expressions.sort_values(by=["Workspace Name", "Dataset Name", "Query Type", "Query / Object Name"], inplace=True)
        df_roles = safe_df(roles_data, ["Workspace Name", "Dataset Name", "Role Name", "Permission", "Table Filters",
                                        "Member Name", "Member Type", "Identity Provider"])
        df_relationships = safe_df(relationships_data, ["Workspace Name", "Dataset Name", "Dataset ID", "Relationship Name",
                                                        "From Table (Fact/Many)", "From Column", "To Table (Dim/One)", "To Column",
                                                        "Cross Filter Direction", "Security Filter Behavior",
                                                        "Cardinality", "Is Active", "Rely On Ref Integrity", "Discovery Source"])
        if not df_relationships.empty:
            df_relationships.drop_duplicates(inplace=True)
        df_pq_joins = safe_df(pq_joins_data, ["Workspace Name", "Dataset Name", "Host Table (From)", "Join From Column",
                                              "Joined Table (To)", "Join To Column", "Join Kind", "Join Function"])
        if not df_pq_joins.empty:
            df_pq_joins.drop_duplicates(inplace=True)
        df_artifact_lineage = safe_df(artifact_lineage_data, ["Workspace Name", "Dataset Name", "Dataset ID",
                                                              "Dependent On Artifact ID", "Dependent Workspace ID",
                                                              "Relation Type", "Settings"])

        dfs = {
            "Workspaces": df_workspaces, "Reports": df_reports, "Dashboards": df_dashboards,
            "Dashboard Tiles": df_tiles, "Datasets": df_datasets, "Data Sources & Lineage": df_datasources,
            "Tables & Columns": df_columns, "Measures (DAX)": df_measures, "M Expressions": df_expressions,
            "RLS Roles": df_roles, "Model Relationships": df_relationships,
            "PQ Implicit Joins": df_pq_joins, "Artifact Lineage": df_artifact_lineage,
        }

        log("")
        log("Extracted records:")
        row_counts = {}
        for name, df in dfs.items():
            cnt = len(df)
            row_counts[name] = cnt
            log(f"  • {name}: {cnt}")
        log("")

        # ---- Excel export ----
        progress(0.92, "Writing Excel file…")
        log("Writing Excel report…")

        wb = Workbook()
        wb.remove(wb.active)
        HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        BORDER = Border(left=Side(style="thin", color="E0E0E0"), right=Side(style="thin", color="E0E0E0"),
                        top=Side(style="thin", color="E0E0E0"), bottom=Side(style="thin", color="E0E0E0"))
        WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
        STD_ALIGN = Alignment(vertical="top")
        WRAP_MAP = {"Measures (DAX)": {"E"}, "M Expressions": {"F"}, "Data Sources & Lineage": {"G"},
                    "Tables & Columns": {"G"}, "RLS Roles": {"E"}, "Dashboard Tiles": {"F"}}

        for sheet_name, df in dfs.items():
            ws_xl = wb.create_sheet(title=sheet_name)
            for row in dataframe_to_rows(df, index=False, header=True):
                ws_xl.append(row)
            ws_xl.freeze_panes = "A2"
            for cell in ws_xl[1]:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = STD_ALIGN
            wrap_cols = WRAP_MAP.get(sheet_name, set())
            for col in ws_xl.columns:
                letter = col[0].column_letter
                max_len = 0
                for cell in col:
                    cell.border = BORDER
                    if cell.row > 1:
                        cell.alignment = WRAP_ALIGN if letter in wrap_cols else STD_ALIGN
                    try:
                        val_str = str(cell.value or "")
                        max_len = max(max_len, 70 if len(val_str) > 100 else len(val_str))
                    except Exception:
                        pass
                ws_xl.column_dimensions[letter].width = min(max_len + 2, 70)

        wb.save(excel_path)
        log(f"Excel saved: {excel_path.name}")

        progress(0.97, "Writing JSON backup…")
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(all_raw_metadata, jf, indent=4, default=str)
        log(f"JSON saved:  {json_path.name}")

        if open_excel_when_done and os.name == "nt":
            try:
                os.startfile(str(excel_path))
            except Exception:
                pass

        progress(1.0, "Completed successfully")
        log("")
        log("Extraction completed successfully.")
        log(f"Excel : {excel_path}")
        log(f"JSON  : {json_path}")

        return {"success": True, "excel_path": str(excel_path), "json_path": str(json_path),
                "message": "Extraction completed successfully.", "row_counts": row_counts}

    except Exception as e:
        friendly = str(e)
        log(f"Error: {friendly}")
        progress(0.0, "Failed")
        return {"success": False, "excel_path": None, "json_path": None,
                "message": friendly, "row_counts": {}}


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Power BI Metadata Extractor")
    parser.add_argument("-o", "--output", type=str, default=None)
    parser.add_argument("--open", action="store_true")
    args = parser.parse_args()
    result = run_extraction(output_dir=args.output, open_excel_when_done=args.open)
    if not result["success"]:
        sys.exit(1)