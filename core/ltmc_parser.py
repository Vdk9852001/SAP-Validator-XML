"""
SAP LTMC / SpreadsheetML XML Parser  —  V3 (Universal)
=======================================================
Handles SAP S/4HANA Migration Cockpit XML exports for ANY SAP object.

Two variants exist depending on the release and object:

  VARIANT A  (e.g. Product, S/4HANA Cloud 2602)
    Row 4  HIDDEN  : SAP table name   "S_MARA"
    Row 5  HIDDEN  : SAP field names  PRODUCT, MTART, MAKTL ...  ← HEADER
    Row 6  HIDDEN  : type specs       ETE;80;0;C;80;0

  VARIANT B  (e.g. Work Center, S/4HANA Cloud 2508)
    Row 4  VISIBLE : SAP table name   "S_WORK_CNTR_HDR"
    Row 5  VISIBLE : SAP field names  ARBPL, WERKS, VERWE ...    ← HEADER
    Row 6  VISIBLE : type specs       ETE;80;0;C;80;0

Fix from V2: scan ALL rows (hidden AND visible) for the field-name row.
The row with the highest proportion of SAP field-code patterns wins.
"""

from __future__ import annotations
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd
import xml.etree.ElementTree as ET


SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"

def _tag(local: str) -> str:
    return f"{{{SS_NS}}}{local}"

def _cell_value(cell_el: ET.Element) -> str:
    data = cell_el.find(_tag("Data"))
    if data is not None and data.text:
        return data.text.strip()
    return ""

def _row_values(row_el: ET.Element) -> List[str]:
    """Extract cell values respecting ss:Index gaps."""
    vals: List[str] = []
    for cell in row_el.findall(_tag("Cell")):
        idx_attr = cell.get(_tag("Index"))
        if idx_attr:
            target = int(idx_attr) - 1
            while len(vals) < target:
                vals.append("")
        vals.append(_cell_value(cell))
    return vals

def _is_hidden(row_el: ET.Element) -> bool:
    return row_el.get(_tag("Hidden"), "0") == "1"

def _is_blank_row(vals: List[str]) -> bool:
    return not any(v.strip() for v in vals)

def _is_field_spec_row(vals: List[str]) -> bool:
    """Type spec rows: ETE;80;0;C;80;0"""
    non_empty = [v for v in vals if v.strip()]
    if not non_empty:
        return False
    spec_count = sum(1 for v in non_empty
                     if re.match(r"^E[A-Z]{2};\d+;\d+;[A-Z];\d+;\d+$", v))
    return spec_count / len(non_empty) >= 0.7

def _is_sap_fieldname_row(vals: List[str]) -> Tuple[bool, float]:
    """
    Check if a row looks like a SAP field-name header row.
    SAP field codes: uppercase letters + digits + underscore, 1-30 chars,
    must start with a letter. Needs >= 80% of non-empty cells to match.
    Returns (is_match, ratio).
    """
    non_empty = [v for v in vals if v.strip()]
    if len(non_empty) < 2:
        return False, 0.0
    sap_count = sum(1 for v in non_empty
                    if re.match(r"^[A-Z][A-Z0-9_]{0,29}$", v.strip()))
    ratio = sap_count / len(non_empty)
    return ratio >= 0.80, ratio

def _is_table_name_row(vals: List[str]) -> bool:
    """SAP table name rows: S_MARA, S_WORK_CNTR_HDR etc."""
    non_empty = [v.strip() for v in vals if v.strip()]
    if not non_empty:
        return False
    return bool(re.match(r"^S_[A-Z0-9_]+$", non_empty[0]))

def _is_description_row(vals: List[str]) -> bool:
    """Verbose description rows (long text or newlines)."""
    non_empty = [v for v in vals if v.strip()]
    if not non_empty:
        return False
    long_count = sum(1 for v in non_empty if len(v) > 60 or "\n" in v)
    return long_count >= max(1, len(non_empty) // 2)

def _is_group_label_row(vals: List[str]) -> bool:
    """Group label rows like ['Key', 'General Data', 'MRP Data']."""
    non_empty = [v.strip() for v in vals if v.strip()]
    if not non_empty:
        return False
    readable = sum(1 for v in non_empty
                   if len(v) <= 50
                   and not re.match(r"^[A-Z][A-Z0-9_]{2,}$", v)
                   and not re.match(r"^E[A-Z]{2};", v))
    return readable >= max(1, len(non_empty) // 2)


_SKIP_SHEETS = {"introduction", "field list"}


def parse_ltmc_xml(file_path: str) -> Dict[str, pd.DataFrame]:
    """
    Parse any SAP LTMC SpreadsheetML XML file.
    Handles Variant A (hidden rows) and Variant B (visible rows) automatically.
    Returns {worksheet_name: DataFrame}.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Not found: {file_path}")

    content = path.read_bytes().lstrip(b"\xef\xbb\xbf")
    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        raise ValueError(f"Cannot parse XML: {e}")

    worksheets = root.findall(f".//{_tag('Worksheet')}")
    if not worksheets:
        raise ValueError("No worksheets found in XML.")

    sheets: Dict[str, pd.DataFrame] = {}

    for ws in worksheets:
        name = ws.get(_tag("Name"), f"Sheet{len(sheets)+1}").strip()
        if name.lower() in _SKIP_SHEETS:
            continue
        table = ws.find(_tag("Table"))
        if table is None:
            continue
        rows = table.findall(_tag("Row"))
        if not rows:
            continue

        df, meta = _parse_data_sheet(name, rows)
        if df is not None and not df.empty:
            df.attrs.update(meta)
            sheets[name] = df

    if not sheets:
        raise ValueError("No data found in any worksheet.")

    return sheets


def _parse_data_sheet(
    sheet_name: str,
    rows: List[ET.Element],
) -> Tuple[Optional[pd.DataFrame], dict]:
    """
    Parse one worksheet.

    KEY CHANGE from V2: scans ALL rows (hidden AND visible) for the
    field-name row. Picks the row with the highest SAP field-code ratio.
    This handles both Variant A (hidden) and Variant B (visible).
    """
    meta = {"table_name": "", "sheet_name": sheet_name, "field_count": 0}

    # ── Find best field-name row by scanning ALL rows ─────────────────────
    best_idx   = None
    best_ratio = 0.0

    for i, row in enumerate(rows):
        vals = _row_values(row)
        is_match, ratio = _is_sap_fieldname_row(vals)
        if is_match and ratio > best_ratio:
            best_ratio = ratio
            best_idx   = i

    if best_idx is None:
        return None, meta

    # ── Capture table name from row immediately before header ─────────────
    if best_idx > 0:
        prev_vals = _row_values(rows[best_idx - 1])
        if _is_table_name_row(prev_vals):
            meta["table_name"] = prev_vals[0].strip()

    # ── Extract column headers ─────────────────────────────────────────────
    header_vals = _row_values(rows[best_idx])
    while header_vals and not header_vals[-1].strip():
        header_vals.pop()
    headers = [v.strip().upper() for v in header_vals]
    meta["field_count"] = len(headers)

    if not headers:
        return None, meta

    # ── Collect data rows ──────────────────────────────────────────────────
    # After the header: skip spec row + group label row + description row
    # (up to 3 metadata rows). As soon as a row doesn't match any of those
    # patterns, start collecting data.
    data_rows: List[List[str]] = []
    skip_budget = 3

    for i in range(best_idx + 1, len(rows)):
        row  = rows[i]
        vals = _row_values(row)

        if _is_blank_row(vals):
            continue

        if skip_budget > 0:
            if (_is_field_spec_row(vals) or
                    _is_group_label_row(vals) or
                    _is_description_row(vals)):
                skip_budget -= 1
                continue
            skip_budget = 0   # first non-metadata row → start data

        padded = vals + [""] * max(0, len(headers) - len(vals))
        data_rows.append(padded[:len(headers)])

    if not data_rows:
        return pd.DataFrame(columns=headers), meta

    df = pd.DataFrame(data_rows, columns=headers)
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip().replace("nan", "")

    return df, meta


def get_sheet_summary(sheets: Dict[str, pd.DataFrame]) -> List[dict]:
    return [
        {
            "sheet_name": name,
            "table_name": df.attrs.get("table_name", ""),
            "row_count":  len(df),
            "col_count":  len(df.columns),
            "columns":    list(df.columns),
        }
        for name, df in sheets.items()
    ]
