"""
SAP Field Label Resolution — V4
================================
Resolves SAP technical field names (KUNNR, MATNR, NAME1 etc.) to
human-friendly business labels (Customer Number, Material Number).

Priority order:
  1. User-uploaded custom labels (custom_labels.csv)
  2. config/field_labels.json
  3. Built-in SAP_FIELD_LABELS dictionary
  4. The raw technical field name itself (fallback)

Usage:
    from core.field_labels import get_label, get_display, enrich_field_rows

    get_label("KUNNR")               -> "Customer Number"
    get_label("NAME1")               -> "Name 1 / Customer Name"
    get_display("NAME1", "NAMORG1")  -> full display dict
"""

import json
import csv
from pathlib import Path
from functools import lru_cache

_LABELS_JSON = Path(__file__).parent.parent / "config" / "field_labels.json"

# ── Built-in fallback dictionary ──────────────────────────────────────────────
SAP_FIELD_LABELS: dict = {
    "MATNR": "Material Number",       "LIFNR": "Vendor Number",
    "KUNNR": "Customer Number",       "SAKNR": "GL Account Number",
    "BELNR": "Document Number",       "EBELN": "Purchase Order Number",
    "VBELN": "Sales Order Number",    "ANLN1": "Asset Number",
    "KOSTL": "Cost Centre",           "PRCTR": "Profit Centre",
    "BANKL": "Bank Key",              "BUKRS": "Company Code",
    "WERKS": "Plant",                 "LGORT": "Storage Location",
    "KOKRS": "Controlling Area",      "VKORG": "Sales Organisation",
    "VTWEG": "Distribution Channel",  "SPART": "Division",
    "NAME1": "Name 1 / Customer Name","NAME2": "Name 2",
    "NAME3": "Name 3",                "NAME4": "Name 4",
    "NAMORG1": "Organisation Name 1", "NAMORG2": "Organisation Name 2",
    "NAMORG3": "Organisation Name 3", "NAMORG4": "Organisation Name 4",
    "NAMECOMBINED": "Combined Name",  "SORTL": "Sort Field / Search Term",
    "SORT1": "Search Term 1",
    "LAND1": "Country",               "COUNTRY": "Country",
    "REGIO": "Region / State",        "REGION": "Region / State",
    "ORT01": "City",                  "CITY1": "City",
    "ORT02": "District",              "CITY2": "District",
    "PSTLZ": "Postal Code",           "POST_CODE1": "Postal Code",
    "STRAS": "Street Address",        "STREET": "Street Address",
    "PFACH": "PO Box",                "PO_BOX": "PO Box",
    "ADRNR": "Address Number",        "HOUSE_NUM1": "House Number",
    "BUILDING": "Building",           "LOCATION": "Location",
    "HOME_CITY": "Home City",         "FLOOR": "Floor",
    "SPRAS": "Language",              "LANGU": "Language",
    "TELF1": "Telephone 1",           "TELF2": "Telephone 2",
    "TELNR_LONG": "Telephone (Long)", "TELFX": "Fax Number",
    "FAXNR_LONG": "Fax Number (Long)","SMTP_ADDR": "Email Address",
    "MOBILE_LONG": "Mobile Number",   "URL_ADDR": "URL / Website",
    "MAKTX": "Material Description",  "MTART": "Material Type",
    "MATKL": "Material Group",        "MEINS": "Base Unit of Measure",
    "BRGEW": "Gross Weight",          "NTGEW": "Net Weight",
    "GEWEI": "Weight Unit",           "VOLUM": "Volume",
    "STPRS": "Standard Price",        "VPRSV": "Price Control",
    "PEINH": "Price Unit",            "WAERS": "Currency",
    "BKLAS": "Valuation Class",       "VERPR": "Moving Average Price",
    "PRDHA": "Product Hierarchy",
    "BLDAT": "Document Date",         "BUDAT": "Posting Date",
    "WRBTR": "Amount (Doc Currency)", "DMBTR": "Amount (Local Currency)",
    "SGTXT": "Item Text",             "ZUONR": "Assignment",
    "BLART": "Document Type",         "GJAHR": "Fiscal Year",
    "ZTERM": "Payment Terms",         "ZFBDT": "Baseline Payment Date",
    "ZBD1T": "Cash Discount Days 1",  "ZBD2T": "Cash Discount Days 2",
    "AKONT": "Reconciliation Account","XBILK": "Balance Sheet Account",
    "FSTAG": "Field Status Group",    "TXT20": "Short Description",
    "TXT50": "Long Description",      "KTOKS": "GL Account Group",
    "KTOKK": "Vendor Account Group",  "KTOKD": "Customer Account Group",
    "ERDAT": "Created On",            "ERNAM": "Created By",
    "LAEDA": "Last Changed Date",
    "LABST": "Unrestricted Stock",    "INSME": "Inspection Stock",
    "DISPO": "MRP Controller",        "DISMM": "MRP Type",
    "MINBE": "Reorder Point",         "EISBE": "Safety Stock",
    "PLIFZ": "Planned Delivery Time", "BESKZ": "Procurement Type",
    "VRKME": "Sales Unit",            "KWMENG": "Order Quantity",
    "NETWR": "Net Value",             "NETPR": "Net Price",
    "MENGE": "Quantity",              "EBELP": "PO Item",
    "POSNR": "Sales Order Item",
    "ANLKL": "Asset Class",           "AKTIV": "Capitalisation Date",
    "ANLN2": "Asset Sub-Number",
    "KTEXT": "Cost Centre Description","KOSAR": "Cost Centre Category",
    "ABTEI": "Department",            "VERAK": "Person Responsible",
    "DATAB": "Valid From",            "DATBI": "Valid To",
    "BANKA": "Bank Name",             "SWIFT": "SWIFT / BIC Code",
    "BANKN": "Bank Account Number",   "BANKS": "Bank Country",
    "BPKIND": "BP Category",         "BU_GROUP": "BP Group",
    "CO_NAME": "Company Name",        "DEFLT_COMM": "Default Communication",
    "LANGU_CORR": "Correspondence Language",
    "NOTE_SMTP": "Email Note",        "NOTE_TELNR": "Telephone Note",
    "NOTE_MOBILE": "Mobile Note",     "TIME_ZONE": "Time Zone",
    "LOEVM": "Deletion Flag",         "BRSCH": "Industry Sector",
    "XCPDK": "One-Time Account",      "XZEMP": "Payment Block",
    "DUEFL": "Dunning Indicator",     "AUFSD": "Order Block",
    "LIFSD": "Delivery Block",        "FAKSD": "Billing Block",
    "STCD1": "Tax Number 1",          "STCD2": "Tax Number 2",
    "XBLNR": "Reference Document",    "EKGRP": "Purchasing Group",
    "BUZEI": "Line Item",             "MANDT": "Client",
    "NIELS": "Nielsen ID",            "COUNC": "County Code",
    "RPMKR": "Regional Market",       "LOCCO": "County",
}


@lru_cache(maxsize=1)
def _load_json_labels() -> dict:
    """Load config/field_labels.json once and cache."""
    if _LABELS_JSON.exists():
        try:
            return json.loads(_LABELS_JSON.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"Warning: could not load field_labels.json: {e}")
    return {}


def load_custom_labels(csv_path: str) -> dict:
    """
    Load user-uploaded label overrides from CSV.
    Format: FIELD_NAME,FRIENDLY_LABEL
    """
    custom = {}
    p = Path(csv_path)
    if not p.exists():
        return custom
    try:
        with open(p, encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) < 2:
                    continue
                key, label = row[0].strip().upper(), row[1].strip()
                if key in ("FIELD", "FIELD_NAME", "SAP_FIELD", "TECHNICAL", "FIELDNAME"):
                    continue
                if key and label:
                    custom[key] = label
    except Exception as e:
        print(f"Warning: could not load custom labels: {e}")
    return custom


def get_label(field_name: str, custom_map: dict = None) -> str:
    """
    Resolve a SAP technical field name to its friendly English label.

    Priority: custom_map > field_labels.json > SAP_FIELD_LABELS > field_name

    Examples:
        get_label("KUNNR")         -> "Customer Number"
        get_label("MATNR")         -> "Material Number"
        get_label("NAME1")         -> "Name 1 / Customer Name"
        get_label("NAMORG1")       -> "Organisation Name 1"
        get_label("LAND1")         -> "Country"
        get_label("PSTLZ")         -> "Postal Code"
        get_label("UNKNOWN_XYZ")   -> "UNKNOWN_XYZ"
    """
    if not field_name:
        return field_name
    key = field_name.strip().upper()
    if custom_map and key in custom_map:
        return custom_map[key]
    json_lbl = _load_json_labels().get(key)
    if json_lbl:
        return json_lbl
    builtin = SAP_FIELD_LABELS.get(key)
    if builtin:
        return builtin
    return field_name


def get_display(
    source_field: str,
    target_field: str = None,
    custom_map:   dict = None,
) -> dict:
    """
    Build the complete display object for a field pair.

    Returns:
        {
          "label":           "Customer Name",
          "source_field":    "NAME1",
          "source_label":    "Name 1 / Customer Name",
          "target_field":    "NAMORG1",
          "target_label":    "Organisation Name 1",
          "display_name":    "Name 1 / Customer Name",
          "display_mapping": "Name 1 / Customer Name  (NAME1 → NAMORG1)",
          "is_cross_mapped": True,
        }
    """
    src_label = get_label(source_field, custom_map)
    tgt_label = get_label(target_field, custom_map) if target_field else src_label
    best_label = src_label if src_label != source_field else tgt_label
    is_cross = (
        target_field is not None and
        target_field.upper() != source_field.upper()
    )
    mapping_str = (
        f"{src_label}  ({source_field} \u2192 {target_field})"
        if is_cross else src_label
    )
    return {
        "label":           best_label,
        "source_field":    source_field.upper(),
        "source_label":    src_label,
        "target_field":    target_field.upper() if target_field else source_field.upper(),
        "target_label":    tgt_label,
        "display_name":    best_label,
        "display_mapping": mapping_str,
        "is_cross_mapped": is_cross,
    }


def enrich_field_rows(field_rows: list, custom_map: dict = None) -> list:
    """
    Add display fields to every entry in a field_results list.
    Call once in run_validation before storing result_dict.
    """
    enriched = []
    for fr in field_rows:
        src = fr.get("field", "")
        tgt = fr.get("field_target", src)
        disp = get_display(src, tgt, custom_map)
        enriched.append({
            **fr,
            "field_label":        disp["source_label"],
            "field_target_label": disp["target_label"],
            "display_name":       disp["display_name"],
            "display_mapping":    disp["display_mapping"],
            "is_cross_mapped":    disp["is_cross_mapped"],
        })
    return enriched
