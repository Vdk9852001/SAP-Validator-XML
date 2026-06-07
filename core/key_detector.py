"""
SAP Composite Key Detector — V4
================================
Dynamically detects the best composite join key for any SAP object
by analysing the actual data in the uploaded files.

Design principles:
  1. Never assume a single column is enough.
  2. Start from known SAP business key candidates.
  3. Test uniqueness — a good key makes every row unique.
  4. Score candidates and pick the minimal set that gives highest uniqueness.
  5. Fall back to all common columns if nothing works.
  6. Always expose the result so the user can override it from the UI.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional
import pandas as pd


# ── SAP composite key catalogue ───────────────────────────────────────────────
# Priority order: most-specific objects first.
# Each entry lists the CANDIDATE key columns in priority order.
# The detector will try combinations starting from the first column.

SAP_KEY_CATALOGUE = {
    # Condition / Pricing records
    # Sales conditions (SD): keyed by KSCHL + MATNR + sales org/channel
    # Purchasing conditions (MM): keyed by KSCHL + MATNR + EKORG (purchasing org)
    # Both variants included so EKORG is always tried alongside VKORG
    "CONDITION":    ["KSCHL","MATNR","EKORG","LIFNR","KUNNR","VKORG","VTWEG",
                     "SPART","WERKS","DATAB","DATBI","PLTYP"],
    "COND_MM":      ["KSCHL","MATNR","EKORG","LIFNR","WERKS","DATAB","DATBI"],
    "COND_SD":      ["KSCHL","MATNR","KUNNR","VKORG","VTWEG","SPART","DATAB","DATBI"],
    "KONP":         ["KNUMH","KOPOS"],
    "KONV":         ["KNUMV","KPOSN","STUNR","ZAEHK"],
    # Purchasing price conditions (A-tables for MM)
    "A016":         ["KSCHL","MATNR","LIFNR","EKORG","DATAB"],
    "A017":         ["KSCHL","MATNR","EKORG","DATAB"],
    "A018":         ["KSCHL","LIFNR","EKORG","DATAB"],
    "A304":         ["KSCHL","MATNR","VKORG","VTWEG","DATAB"],
    "A305":         ["KSCHL","KUNNR","MATNR","VKORG","VTWEG","DATAB"],

    # Material
    "MATERIAL":     ["MATNR","WERKS","LGORT","SPRAS"],
    "MARA":         ["MATNR"],
    "MARC":         ["MATNR","WERKS"],
    "MARD":         ["MATNR","WERKS","LGORT"],
    "MAKT":         ["MATNR","SPRAS"],
    "MBEW":         ["MATNR","BWKEY","BWTAR"],
    "MVKE":         ["MATNR","VKORG","VTWEG"],

    # Customer / Business Partner
    "CUSTOMER":     ["KUNNR","VKORG","VTWEG","SPART","SPRAS"],
    "KNA1":         ["KUNNR"],
    "KNVV":         ["KUNNR","VKORG","VTWEG","SPART"],
    "KNVP":         ["KUNNR","PARVW","PARID"],
    "KNVK":         ["KUNNR","PARNR"],
    "KNB1":         ["KUNNR","BUKRS"],

    # Vendor / Supplier
    "VENDOR":       ["LIFNR","EKORG","BUKRS","WERKS"],
    "LFA1":         ["LIFNR"],
    "LFB1":         ["LIFNR","BUKRS"],
    "LFM1":         ["LIFNR","EKORG"],
    "LFM2":         ["LIFNR","EKORG","WERKS"],

    # Finance
    "GL_ACCOUNT":   ["SAKNR","BUKRS"],
    "SKA1":         ["SAKNR","KTOPL"],
    "SKB1":         ["SAKNR","BUKRS"],
    "OPEN_ITEMS_AR":["BELNR","GJAHR","BUZEI","BUKRS"],
    "OPEN_ITEMS_AP":["BELNR","GJAHR","BUZEI","BUKRS"],
    "BSEG":         ["BUKRS","BELNR","GJAHR","BUZEI"],
    "BKPF":         ["BUKRS","BELNR","GJAHR"],

    # Purchasing
    "PURCHASE_ORDER":["EBELN","EBELP"],
    "EKKO":         ["EBELN"],
    "EKPO":         ["EBELN","EBELP"],
    "EINA":         ["INFNR","MATNR"],
    "EINE":         ["INFNR","EKORG","WERKS"],

    # Sales
    "SALES_ORDER":  ["VBELN","POSNR"],
    "VBAK":         ["VBELN"],
    "VBAP":         ["VBELN","POSNR"],

    # Asset
    "ASSET":        ["ANLN1","ANLN2","BUKRS"],
    "ANLA":         ["BUKRS","ANLN1","ANLN2"],

    # Controlling
    "COST_CENTRE":  ["KOSTL","KOKRS","DATAB"],
    "PROFIT_CENTRE":["PRCTR","KOKRS","DATAB"],
    "CSKS":         ["KOKRS","KOSTL","DATAB"],
    "CEPC":         ["KOKRS","PRCTR","DATAB"],

    # Plant / Org
    "INVENTORY":    ["MATNR","WERKS","LGORT"],
    "PLANT":        ["WERKS","BUKRS"],
    "BANK":         ["BANKL","BANKS"],

    # Info records
    "INFO_RECORD":  ["INFNR","MATNR","LIFNR","EKORG","WERKS"],
    "EBAN":         ["BANFN","BNFPO"],
}

# Single-column fallback priority (used when no composite key found)
SINGLE_KEY_PRIORITY = [
    "BELNR","EBELN","VBELN","INFNR","ANLN1",
    "MATNR","KUNNR","LIFNR","SAKNR","KOSTL","PRCTR","BANKL",
    "WERKS","BUKRS","KOKRS","VKORG","EKORG",
]


@dataclass
class KeyDetectionResult:
    join_keys:          List[str]      # final composite key columns
    uniqueness_src:     float          # % unique rows in source with this key
    uniqueness_tgt:     float          # % unique rows in target
    detection_method:   str            # "catalogue", "auto", "single", "fallback"
    duplicate_src:      int            # rows with duplicate keys in source
    duplicate_tgt:      int            # rows with duplicate keys in target
    candidate_tested:   List[str]      # all columns that were considered
    confidence:         str            # "high", "medium", "low"
    object_hint:        str = ""       # which catalogue entry was used


def detect_composite_key(
    src_df:        pd.DataFrame,
    tgt_df:        pd.DataFrame,
    object_name:   str   = "",
    manual_keys:   List[str] = None,
    max_key_cols:  int   = 5,
    min_uniqueness: float = 0.98,
) -> KeyDetectionResult:
    """
    Detect the best composite join key for a given source/target pair.

    Parameters
    ----------
    src_df         : source DataFrame (headers already upper-cased)
    tgt_df         : target DataFrame
    object_name    : hint e.g. "CONDITION", "CUSTOMER", "A305"
    manual_keys    : if the user has specified keys manually, use these directly
    max_key_cols   : maximum columns to include in composite key
    min_uniqueness : target uniqueness ratio (0.98 = 98% unique rows)

    Returns
    -------
    KeyDetectionResult
    """
    common_cols = sorted(set(src_df.columns) & set(tgt_df.columns))

    # ── Manual override: trust the user completely ────────────────────────────
    if manual_keys:
        keys = [k.upper() for k in manual_keys if k.upper() in set(src_df.columns) & set(tgt_df.columns)]
        if keys:
            u_src, d_src = _uniqueness(src_df, keys)
            u_tgt, d_tgt = _uniqueness(tgt_df, keys)
            return KeyDetectionResult(
                join_keys=keys,
                uniqueness_src=u_src, uniqueness_tgt=u_tgt,
                detection_method="manual",
                duplicate_src=d_src, duplicate_tgt=d_tgt,
                candidate_tested=keys,
                confidence="high",
                object_hint="user-defined",
            )

    # ── Step 1: Try catalogue-based detection ─────────────────────────────────
    catalogue_keys = _catalogue_candidates(object_name, common_cols)
    if catalogue_keys:
        result = _test_key_combination(
            src_df, tgt_df, catalogue_keys, common_cols,
            max_key_cols, min_uniqueness, "catalogue"
        )
        if result and result.confidence in ("high", "medium"):
            return result

    # ── Step 2: Auto-detect by uniqueness scoring ─────────────────────────────
    auto_result = _auto_detect(
        src_df, tgt_df, common_cols, max_key_cols, min_uniqueness
    )
    if auto_result and auto_result.confidence in ("high", "medium"):
        return auto_result

    # ── Step 3: Single-column fallback ────────────────────────────────────────
    for col in SINGLE_KEY_PRIORITY:
        if col in common_cols:
            u_src, d_src = _uniqueness(src_df, [col])
            u_tgt, d_tgt = _uniqueness(tgt_df, [col])
            conf = "medium" if u_src >= 0.90 else "low"
            return KeyDetectionResult(
                join_keys=[col],
                uniqueness_src=u_src, uniqueness_tgt=u_tgt,
                detection_method="single",
                duplicate_src=d_src, duplicate_tgt=d_tgt,
                candidate_tested=common_cols[:20],
                confidence=conf,
                object_hint="",
            )

    # ── Step 4: Use first common column ───────────────────────────────────────
    fallback = [common_cols[0]] if common_cols else ["__row__"]
    u_src, d_src = _uniqueness(src_df, fallback) if common_cols else (0.0, 0)
    u_tgt, d_tgt = _uniqueness(tgt_df, fallback) if common_cols else (0.0, 0)
    return KeyDetectionResult(
        join_keys=fallback,
        uniqueness_src=u_src, uniqueness_tgt=u_tgt,
        detection_method="fallback",
        duplicate_src=d_src, duplicate_tgt=d_tgt,
        candidate_tested=common_cols[:10],
        confidence="low",
    )


# ── Internal helpers ───────────────────────────────────────────────────────────

def _uniqueness(df: pd.DataFrame, cols: List[str]) -> tuple:
    """Return (uniqueness_ratio, duplicate_count) for a set of columns."""
    valid_cols = [c for c in cols if c in df.columns]
    if not valid_cols or len(df) == 0:
        return 0.0, 0
    n_unique = df[valid_cols].drop_duplicates().shape[0]
    n_total  = len(df)
    dupes    = n_total - n_unique
    return round(n_unique / n_total, 4), dupes


def _catalogue_candidates(object_name: str, common_cols: List[str]) -> List[str]:
    """
    Return catalogue key columns that exist in common_cols, for the best matching object.

    Matching priority:
      1. Exact object name match  (CONDITION → CONDITION)
      2. Name contains match      (MY_CONDITION_DATA → CONDITION)
      3. Best overlap score       — picks the entry whose candidate columns
                                    overlap most with what is actually in the file
                                    This ensures COND_MM wins over COND_SD when
                                    the file has EKORG but not VKORG, and vice versa.
    """
    if not object_name:
        return []
    name       = object_name.upper().replace("-", "_").replace(" ", "_")
    common_set = set(common_cols)

    # Step 1: direct or partial name match — collect all matching entries
    matched_entries = []
    for key, cols in SAP_KEY_CATALOGUE.items():
        if key == name or key in name or name in key:
            overlap = [c for c in cols if c in common_set]
            if overlap:
                matched_entries.append((len(overlap), key, overlap, cols))

    if matched_entries:
        # Among matching entries, pick the one with the most columns
        # actually present in the file — this selects COND_MM (has EKORG)
        # over COND_SD (has VKORG) when the file contains EKORG.
        matched_entries.sort(key=lambda x: -x[0])
        best_overlap, best_key, best_cols, full_cols = matched_entries[0]

        # If multiple entries tied on overlap count, prefer the one whose
        # first candidate column (highest priority key) is in the file
        top = [e for e in matched_entries if e[0] == best_overlap]
        for entry in top:
            if entry[2] and entry[2][0] in common_set:
                return entry[2]
        return best_cols

    # Step 2: no name match — score ALL catalogue entries by column overlap
    scores = []
    for key, cols in SAP_KEY_CATALOGUE.items():
        overlap = [c for c in cols if c in common_set]
        if len(overlap) >= 2:   # need at least 2 matching cols to be meaningful
            scores.append((len(overlap), key, overlap))

    if scores:
        scores.sort(reverse=True)
        return scores[0][2]

    return []


def _test_key_combination(
    src_df: pd.DataFrame,
    tgt_df: pd.DataFrame,
    candidates: List[str],
    common_cols: List[str],
    max_cols: int,
    min_uniqueness: float,
    method: str,
) -> Optional[KeyDetectionResult]:
    """
    Try incrementally adding candidate columns until uniqueness target is met.
    """
    best_keys  = []
    best_u_src = 0.0
    best_u_tgt = 0.0

    for i in range(1, min(len(candidates), max_cols) + 1):
        keys   = candidates[:i]
        u_src, d_src = _uniqueness(src_df, keys)
        u_tgt, d_tgt = _uniqueness(tgt_df, keys)
        u_min  = min(u_src, u_tgt)

        if u_min > best_u_src or (u_min == best_u_src and len(keys) > len(best_keys)):
            best_keys  = keys
            best_u_src = u_src
            best_u_tgt = u_tgt
            best_d_src = d_src
            best_d_tgt = d_tgt

        if u_min >= min_uniqueness:
            break

    if not best_keys:
        return None

    u_min = min(best_u_src, best_u_tgt)
    conf  = "high" if u_min >= 0.99 else "medium" if u_min >= 0.95 else "low"

    return KeyDetectionResult(
        join_keys=best_keys,
        uniqueness_src=best_u_src, uniqueness_tgt=best_u_tgt,
        detection_method=method,
        duplicate_src=best_d_src, duplicate_tgt=best_d_tgt,
        candidate_tested=candidates,
        confidence=conf,
    )


def _auto_detect(
    src_df: pd.DataFrame,
    tgt_df: pd.DataFrame,
    common_cols: List[str],
    max_cols: int,
    min_uniqueness: float,
) -> Optional[KeyDetectionResult]:
    """
    Automatically score all common columns and build composite key.
    Strategy:
      1. Score each column by uniqueness in source + target combined.
      2. Start with the highest-scoring column.
      3. Keep adding columns (in score order) until uniqueness target is met.
    """
    if not common_cols:
        return None

    # Score each column: (uniqueness_src + uniqueness_tgt) / 2
    scores = []
    for col in common_cols:
        u_src, _ = _uniqueness(src_df, [col])
        u_tgt, _ = _uniqueness(tgt_df, [col])
        scores.append((col, (u_src + u_tgt) / 2))
    scores.sort(key=lambda x: -x[1])

    # Build composite key greedily
    selected = []
    for col, score in scores:
        if len(selected) >= max_cols:
            break
        # Skip columns that add no uniqueness improvement
        trial = selected + [col]
        u_src, d_src = _uniqueness(src_df, trial)
        u_tgt, d_tgt = _uniqueness(tgt_df, trial)
        u_min_trial  = min(u_src, u_tgt)
        u_min_prev   = min(_uniqueness(src_df, selected)[0], _uniqueness(tgt_df, selected)[0]) if selected else 0.0

        if u_min_trial > u_min_prev or not selected:
            selected = trial
            if u_min_trial >= min_uniqueness:
                break

    if not selected:
        return None

    u_src, d_src = _uniqueness(src_df, selected)
    u_tgt, d_tgt = _uniqueness(tgt_df, selected)
    u_min = min(u_src, u_tgt)
    conf  = "high" if u_min >= 0.99 else "medium" if u_min >= 0.95 else "low"

    return KeyDetectionResult(
        join_keys=selected,
        uniqueness_src=u_src, uniqueness_tgt=u_tgt,
        detection_method="auto",
        duplicate_src=d_src, duplicate_tgt=d_tgt,
        candidate_tested=[c for c, _ in scores[:20]],
        confidence=conf,
    )
