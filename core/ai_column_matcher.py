"""
AI Column Matcher — powered by Claude
======================================
Uses Claude to intelligently map post-load file columns to SAP technical
field names from the LTMC template, for any SAP object.

Why LLM over pure dictionary lookup:
  - Handles columns never seen before ("Net Book Value" → CNSNG_PRICE)
  - Understands SAP context ("Plant" in a WC sheet = WERKS, not PRVBE)
  - Learns patterns across objects (numbered fields, Ref. suffixes, etc.)
  - Resolves ambiguous labels that collide in the static dictionary
  - Works for custom Z-fields that have no entry in field_labels.json

Caching:
  - All AI resolutions saved to ai_column_cache.json
  - Key = (sap_object, sheet_name, postload_col, ltmc_cols_hash)
  - Cache survives server restarts — no re-billing for same columns
"""

from __future__ import annotations
import json
import hashlib
import re
import time
from pathlib import Path
from typing import Dict, List, Optional

_CACHE_FILE = Path(__file__).parent.parent / "ai_column_cache.json"
_cache: dict = {}
_cache_dirty = False


def _load_cache():
    global _cache
    if _CACHE_FILE.exists():
        try:
            _cache = json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            _cache = {}


def _save_cache():
    global _cache_dirty
    if _cache_dirty:
        _CACHE_FILE.write_text(
            json.dumps(_cache, indent=2), encoding="utf-8"
        )
        _cache_dirty = False


def _cache_key(
    sap_object: str,
    sheet_name: str,
    postload_col: str,
    ltmc_cols: List[str],
) -> str:
    ltmc_hash = hashlib.md5("|".join(sorted(ltmc_cols)).encode()).hexdigest()[:8]
    return f"{sap_object}|{sheet_name}|{postload_col}|{ltmc_hash}"


# ── Main public function ──────────────────────────────────────────────────────

def ai_resolve_columns(
    postload_columns: List[str],
    ltmc_columns: List[str],
    sap_object: str,
    sheet_name: str,
    already_resolved: Dict[str, dict] = None,
    field_labels: dict = None,
) -> Dict[str, dict]:
    """
    Use Claude to resolve post-load column names to SAP technical field codes.

    Only calls the API for columns that:
    1. Are NOT already matched by the rule-based resolver
    2. Are NOT already in the cache

    Parameters
    ----------
    postload_columns   : all post-load file column names
    ltmc_columns       : SAP technical field names from the LTMC sheet
    sap_object         : e.g. "PRODUCT", "WORK_CENTER", "CUSTOMER"
    sheet_name         : e.g. "Basic Data", "Plant Data"
    already_resolved   : result from rule-based label_resolver (skip matched ones)
    field_labels       : {SAP_CODE: friendly_label} for context

    Returns
    -------
    dict  {postload_col: {resolved, method, matched, ltmc_col, confidence}}
    """
    _load_cache()

    already_resolved = already_resolved or {}
    field_labels     = field_labels or {}

    # Columns still needing AI resolution
    unresolved = [
        col for col in postload_columns
        if not already_resolved.get(col, {}).get("matched", False)
    ]

    if not unresolved:
        return {}

    # Split into cached and truly new
    cached_results  = {}
    need_api        = []

    for col in unresolved:
        key = _cache_key(sap_object, sheet_name, col, ltmc_columns)
        if key in _cache:
            cached_results[col] = _cache[key]
        else:
            need_api.append(col)

    if not need_api:
        return cached_results

    # Call Claude API for uncached columns
    api_results = _call_claude(
        unresolved_cols=need_api,
        ltmc_columns=ltmc_columns,
        sap_object=sap_object,
        sheet_name=sheet_name,
        field_labels=field_labels,
    )

    # Cache the results
    global _cache_dirty
    for col, result in api_results.items():
        key = _cache_key(sap_object, sheet_name, col, ltmc_columns)
        _cache[key] = result
        _cache_dirty = True

    _save_cache()

    return {**cached_results, **api_results}


# ── Claude API call ───────────────────────────────────────────────────────────

def _call_claude(
    unresolved_cols: List[str],
    ltmc_columns: List[str],
    sap_object: str,
    sheet_name: str,
    field_labels: dict,
) -> Dict[str, dict]:
    """
    Call Claude claude-sonnet-4-20250514 to resolve column names.
    Returns {postload_col: {resolved, method, matched, ltmc_col, confidence}}
    """
    import urllib.request

    # Build context: include labels for LTMC columns if available
    ltmc_with_labels = []
    for col in ltmc_columns:
        label = field_labels.get(col.upper(), "")
        if label and label != col:
            ltmc_with_labels.append(f"{col} ({label})")
        else:
            ltmc_with_labels.append(col)

    prompt = f"""You are an SAP data migration expert. Your task is to map post-load file column names to their corresponding SAP technical field codes from a Migration Cockpit (LTMC) template.

SAP Object: {sap_object}
Sheet: {sheet_name}

LTMC technical field codes available (with friendly labels where known):
{json.dumps(ltmc_with_labels, indent=2)}

Post-load file columns that need to be mapped to the above LTMC codes:
{json.dumps(unresolved_cols, indent=2)}

For each post-load column, find the best matching LTMC technical field code.
Rules:
- Map to the EXACT code from the LTMC list above — do not invent codes
- If a column clearly corresponds to an LTMC field (even with different naming), map it
- If a column has NO reasonable match in the LTMC list, set "matched": false
- Use your knowledge of SAP field naming conventions (e.g. "Plant" = WERKS, "Work Center" = ARBPL)
- Consider the object type and sheet context when resolving ambiguous names
- Numbered fields: "Activity Type 1" likely maps to LSTAR1 or BDE1 etc.
- "Ref." suffix: "Setup Type Ref." maps to the _REF variant of the base field

Respond with ONLY a valid JSON object in this exact format, nothing else:
{{
  "mappings": [
    {{
      "postload_col": "<original column name>",
      "ltmc_col": "<matched LTMC code or null>",
      "matched": true/false,
      "confidence": "high/medium/low",
      "reasoning": "<one sentence why>"
    }}
  ]
}}"""

    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 2000,
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type":      "application/json",
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        # API error — return original columns as unmatched
        return {
            col: {
                "original": col,
                "resolved": col.upper(),
                "method": "ai_error",
                "matched": False,
                "ltmc_col": None,
                "confidence": "low",
                "reasoning": str(e),
            }
            for col in unresolved_cols
        }

    # Parse Claude response
    raw_text = ""
    for block in data.get("content", []):
        if block.get("type") == "text":
            raw_text += block.get("text", "")

    results = {}
    try:
        # Extract JSON from response (Claude may add markdown fences)
        json_match = re.search(r'\{.*\}', raw_text, re.DOTALL)
        if json_match:
            parsed = json.loads(json_match.group(0))
            ltmc_set = set(c.upper() for c in ltmc_columns)
            for mapping in parsed.get("mappings", []):
                col       = mapping.get("postload_col", "")
                ltmc_col  = mapping.get("ltmc_col")
                matched   = mapping.get("matched", False)
                conf      = mapping.get("confidence", "medium")
                reasoning = mapping.get("reasoning", "")

                # Validate: ltmc_col must exist in the LTMC sheet
                if ltmc_col and ltmc_col.upper() in ltmc_set:
                    ltmc_col = ltmc_col.upper()
                    matched  = True
                else:
                    ltmc_col = None
                    matched  = False

                results[col] = {
                    "original":   col,
                    "resolved":   ltmc_col or col.upper(),
                    "method":     f"ai_{conf}",
                    "matched":    matched,
                    "ltmc_col":   ltmc_col,
                    "confidence": conf,
                    "reasoning":  reasoning,
                }
    except Exception as e:
        pass  # Fall through to default

    # Any column not in the parsed results gets a default unmatched entry
    for col in unresolved_cols:
        if col not in results:
            results[col] = {
                "original":   col,
                "resolved":   col.upper(),
                "method":     "ai_unmatched",
                "matched":    False,
                "ltmc_col":   None,
                "confidence": "low",
                "reasoning":  "Not resolved by AI",
            }

    return results


# ── Cache management helpers ──────────────────────────────────────────────────

def get_cache_stats() -> dict:
    _load_cache()
    return {
        "total_entries": len(_cache),
        "cache_file": str(_CACHE_FILE),
        "exists": _CACHE_FILE.exists(),
    }


def clear_cache(sap_object: str = None, sheet_name: str = None):
    global _cache, _cache_dirty
    _load_cache()
    if sap_object is None:
        _cache = {}
    else:
        prefix = f"{sap_object}|{sheet_name or ''}".rstrip("|")
        _cache = {k: v for k, v in _cache.items() if not k.startswith(prefix)}
    _cache_dirty = True
    _save_cache()


def get_cached_mappings(sap_object: str = None) -> list:
    """Return human-readable cache entries, optionally filtered by object."""
    _load_cache()
    entries = []
    for key, val in _cache.items():
        parts = key.split("|")
        if len(parts) >= 3:
            obj, sheet, col = parts[0], parts[1], parts[2]
            if sap_object and obj.upper() != sap_object.upper():
                continue
            entries.append({
                "sap_object":  obj,
                "sheet_name":  sheet,
                "postload_col": col,
                "ltmc_col":    val.get("ltmc_col"),
                "matched":     val.get("matched"),
                "method":      val.get("method"),
                "confidence":  val.get("confidence"),
                "reasoning":   val.get("reasoning", ""),
            })
    return sorted(entries, key=lambda x: (x["sap_object"], x["sheet_name"], x["postload_col"]))
