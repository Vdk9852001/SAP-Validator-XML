"""
Label Resolver
==============
Resolves post-load file columns (which may use friendly/business names)
back to SAP technical field codes so they can be matched against
LTMC template columns (which always use SAP technical names).

Examples:
  "Customer Number"  → KUNNR
  "Name"             → NAME1  or  NAMORG1
  "Country"          → LAND1  or  COUNTRY
  "Vendor Number"    → LIFNR
  "Material"         → MATNR

Resolution strategy (in priority order):
  1. Exact SAP technical name match (column already is KUNNR etc.)
  2. Reverse label lookup: field_labels dict {KUNNR: "Customer Number"}
     → build reverse {lower("Customer Number"): "KUNNR"}
  3. Fuzzy match on label (similarity >= 0.80)
  4. Return original name unchanged
"""

from __future__ import annotations
import re
from typing import Dict, List, Optional, Tuple


def _normalize(s: str) -> str:
    """Lowercase, strip, collapse whitespace and punctuation."""
    return re.sub(r"[\s_\-/]+", " ", str(s).lower().strip())


def _similarity(a: str, b: str) -> float:
    """Simple LCS-based similarity."""
    a, b = _normalize(a), _normalize(b)
    if a == b:
        return 1.0
    if not a or not b:
        return 0.0
    la, lb = len(a), len(b)
    dp = [[0] * (lb + 1) for _ in range(la + 1)]
    for i in range(1, la + 1):
        for j in range(1, lb + 1):
            dp[i][j] = dp[i-1][j-1]+1 if a[i-1] == b[j-1] else max(dp[i-1][j], dp[i][j-1])
    return (2.0 * dp[la][lb]) / (la + lb)


def build_reverse_label_map(
    field_labels: Dict[str, str],
    custom_labels: Dict[str, str] = None,
) -> Dict[str, str]:
    """
    Build a reverse map: normalized_label → SAP_FIELD_CODE.
    Merges built-in field_labels with any custom overrides.
    Custom labels take priority.
    """
    merged = dict(field_labels)
    if custom_labels:
        merged.update(custom_labels)

    reverse: Dict[str, str] = {}
    for code, label in merged.items():
        key = _normalize(label)
        # Don't overwrite with a worse match
        if key not in reverse:
            reverse[key] = code.upper()
    return reverse


def resolve_column(
    col_name: str,
    known_sap_fields: set,
    reverse_label_map: Dict[str, str],
    fuzzy_threshold: float = 0.80,
) -> Tuple[str, str]:
    """
    Resolve a single column name to a SAP field code.

    Returns (resolved_code, method) where method is one of:
      "exact"   — already a SAP code
      "label"   — matched via reverse label lookup
      "fuzzy"   — matched via fuzzy similarity
      "original"— could not resolve, returned as-is
    """
    col_upper = col_name.strip().upper()

    # 1. Already a SAP technical name
    if col_upper in known_sap_fields:
        return col_upper, "exact"

    # 2. Reverse label lookup
    col_norm = _normalize(col_name)
    if col_norm in reverse_label_map:
        return reverse_label_map[col_norm], "label"

    # 3. Fuzzy match against all known labels
    best_code, best_score = None, 0.0
    for label_norm, code in reverse_label_map.items():
        score = _similarity(col_norm, label_norm)
        if score > best_score:
            best_score, best_code = score, code

    if best_code and best_score >= fuzzy_threshold:
        return best_code, f"fuzzy({best_score:.0%})"

    # 4. Try matching directly against SAP field codes
    for sap_col in known_sap_fields:
        if _similarity(col_upper, sap_col) >= 0.90:
            return sap_col, "fuzzy_code"

    return col_upper, "original"


def resolve_postload_columns(
    postload_columns: List[str],
    ltmc_columns: List[str],
    field_labels: Dict[str, str],
    custom_labels: Dict[str, str] = None,
    fuzzy_threshold: float = 0.80,
) -> Dict[str, dict]:
    """
    Resolve all post-load column names against a set of known LTMC column names.

    Parameters
    ----------
    postload_columns : columns from the post-load file (may be friendly names)
    ltmc_columns     : columns from the LTMC template (SAP technical names)
    field_labels     : {SAP_CODE: "Friendly Label"} — full label dictionary
    custom_labels    : user-uploaded label overrides
    fuzzy_threshold  : minimum similarity for fuzzy matching

    Returns
    -------
    dict mapping each postload column to:
      {
        "resolved":   "KUNNR",        # resolved SAP technical name
        "method":     "label",         # how it was resolved
        "matched":    True,            # whether it matched an LTMC column
        "ltmc_col":   "KUNNR",         # the LTMC column it matched (or None)
      }
    """
    known_sap = set(c.upper() for c in ltmc_columns)
    rev_map   = build_reverse_label_map(field_labels, custom_labels)

    result = {}
    for col in postload_columns:
        resolved, method = resolve_column(col, known_sap, rev_map, fuzzy_threshold)
        matched   = resolved in known_sap
        ltmc_col  = resolved if matched else None
        result[col] = {
            "original": col,
            "resolved": resolved,
            "method":   method,
            "matched":  matched,
            "ltmc_col": ltmc_col,
        }
    return result


def build_field_map_from_resolution(
    resolution: Dict[str, dict],
) -> Dict[str, str]:
    """
    From a resolution dict, build a field_map {ltmc_col: postload_col}
    i.e. {SAP_technical_name: post_load_column_name}
    for use with MaterialValidator.
    """
    field_map = {}
    for postload_col, info in resolution.items():
        if info["matched"] and info["ltmc_col"]:
            ltmc_col = info["ltmc_col"]
            # If multiple postload cols resolve to same LTMC col, first wins
            if ltmc_col not in field_map:
                field_map[ltmc_col] = postload_col
    return field_map
