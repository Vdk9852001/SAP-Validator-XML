"""
SAP Migration Post-Load Validator — Core Engine V4
===================================================
Join key philosophy:
  - User-selected keys are ALWAYS used when provided (highest priority)
  - Auto-detection from key_detector is used ONLY as a suggestion, never forced
  - If no keys are selected and auto-detect finds nothing, validation stops
    and tells the user to select keys from the UI
  - Composite key = "VAL1||VAL2||VAL3" column built from selected fields
  - Every field merge uses this composite key — no single-column joins
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass, field
from typing import Optional, Dict, List
import logging

logger = logging.getLogger(__name__)

DEFAULT_PASS_THRESHOLD = 100.0
_CK = "__CK__"   # internal composite key column


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class FieldResult:
    field_source:      str
    field_target:      str
    field_label:       str
    total_records:     int
    matched:           int
    mismatched:        int
    missing_in_target: int
    missing_in_source: int
    is_numeric:        bool  = False
    is_key_field:      bool  = False   # True when this field is a join key
    tolerance_used:    float = None
    pass_threshold:    float = DEFAULT_PASS_THRESHOLD
    mismatch_details:  list  = field(default_factory=list)

    @property
    def match_pct(self):
        if self.total_records == 0:
            return 0.0
        return round(self.matched / self.total_records * 100, 2)

    @property
    def status(self):
        return "PASS" if self.match_pct >= self.pass_threshold else "FAIL"


@dataclass
class MappingReport:
    join_keys:          List[str]
    join_key_labels:    Dict[str, str]
    matched_fields:     list
    source_only_fields: list
    target_only_fields: list
    numeric_fields:     list
    tolerance_map:      dict
    total_source_cols:  int
    total_target_cols:  int
    selected_fields:    list  = field(default_factory=list)
    pass_threshold:     float = DEFAULT_PASS_THRESHOLD

    @property
    def join_key(self):
        return "+".join(self.join_keys) if self.join_keys else ""

    @property
    def join_key_label(self):
        return " + ".join(self.join_key_labels.get(k, k) for k in self.join_keys)


@dataclass
class ValidationResult:
    source_file:             str
    target_file:             str
    total_source_records:    int
    total_target_records:    int
    records_matched:         int
    records_only_in_source:  int
    records_only_in_target:  int
    mapping:                 MappingReport = None
    field_results:           list = field(default_factory=list)
    errors:                  list = field(default_factory=list)
    join_keys:               List[str] = field(default_factory=list)
    key_detection_method:    str = ""
    key_confidence:          str = ""
    duplicate_src:           int = 0
    duplicate_tgt:           int = 0
    duplicate_key_samples:   list = field(default_factory=list)
    common_columns:          list = field(default_factory=list)  # for UI

    @property
    def overall_status(self):
        if self.errors:
            return "ERROR"
        return "FAIL" if any(f.status == "FAIL" for f in self.field_results) else "PASS"

    @property
    def summary_stats(self):
        total  = len(self.field_results)
        passed = sum(1 for f in self.field_results if f.status == "PASS")
        return {
            "total_fields_validated": total,
            "fields_passed":          passed,
            "fields_failed":          total - passed,
            "pass_rate_pct":          round(passed / total * 100, 1) if total else 0,
        }


# ── Validator ─────────────────────────────────────────────────────────────────

class MaterialValidator:

    def __init__(
        self,
        field_map:           Dict[str, str] = None,
        tolerance_map:       dict  = None,
        join_key:            str   = None,     # single key (legacy)
        join_keys:           List[str] = None, # composite key (new)
        manual_join_keys:    List[str] = None, # user-selected from UI (highest priority)
        pass_threshold:      float = DEFAULT_PASS_THRESHOLD,
        selected_fields:     list  = None,
        numeric_sample_rows: int   = 200,
        numeric_threshold:   float = 0.80,
        custom_labels              = None,
    ):
        self.field_map           = field_map
        self.tolerance_overrides = tolerance_map or {}

        # Join key priority: manual_join_keys > join_keys > join_key (single)
        if manual_join_keys:
            self.manual_join_keys = [k.strip().upper() for k in manual_join_keys if k.strip()]
        elif join_keys:
            self.manual_join_keys = [k.strip().upper() for k in join_keys if k.strip()]
        elif join_key:
            self.manual_join_keys = [join_key.strip().upper()]
        else:
            self.manual_join_keys = []

        self.pass_threshold      = pass_threshold
        self.selected_fields     = [f.strip().upper() for f in (selected_fields or [])]
        self.numeric_sample_rows = numeric_sample_rows
        self.numeric_threshold   = numeric_threshold

        from core.field_labels import load_custom_labels
        if isinstance(custom_labels, str):
            self.custom_labels = load_custom_labels(custom_labels)
        elif isinstance(custom_labels, dict):
            self.custom_labels = {k.upper(): v for k, v in custom_labels.items()}
        else:
            self.custom_labels = {}

    def _label(self, field_name):
        from core.field_labels import get_label
        return get_label(field_name, self.custom_labels)

    # ── Main entry point ──────────────────────────────────────────────────────

    def validate(
        self,
        source_path:       str,
        target_path:       str,
        source_delimiter:  str = ",",
        target_delimiter:  str = ",",
        max_mismatch_rows: int = 200,
        object_name:       str = "",
    ) -> ValidationResult:

        # ── Step 1: Read headers ──────────────────────────────────────────────
        try:
            src_hdrs = self._read_headers(source_path, source_delimiter)
            tgt_hdrs = self._read_headers(target_path, target_delimiter)
        except Exception as e:
            return ValidationResult(
                source_file=source_path, target_file=target_path,
                total_source_records=0, total_target_records=0,
                records_matched=0, records_only_in_source=0,
                records_only_in_target=0,
                errors=[f"Cannot read file headers: {e}"]
            )

        common_cols = sorted(set(src_hdrs) & set(tgt_hdrs))

        # ── Step 2: Resolve join keys ─────────────────────────────────────────
        # Priority 1: user manually selected keys from UI
        if self.manual_join_keys:
            valid_keys = [k for k in self.manual_join_keys
                         if k in set(src_hdrs) and k in set(tgt_hdrs)]
            if valid_keys:
                join_keys         = valid_keys
                detection_method  = "manual"
                confidence        = "high"
                logger.info(f"Using user-selected join keys: {join_keys}")
            else:
                # Keys specified but not found in files
                missing = [k for k in self.manual_join_keys
                          if k not in set(src_hdrs) or k not in set(tgt_hdrs)]
                return ValidationResult(
                    source_file=source_path, target_file=target_path,
                    total_source_records=0, total_target_records=0,
                    records_matched=0, records_only_in_source=0,
                    records_only_in_target=0,
                    common_columns=common_cols,
                    errors=[
                        f"Selected join key(s) not found in both files: {missing}. "
                        f"Common columns available: {common_cols[:20]}"
                    ]
                )

        # Priority 2: auto-detect using key_detector (suggestion only)
        else:
            try:
                # Load a small sample just for key detection (fast)
                src_sample = self._load_file_cols(source_path, source_delimiter,
                                                   common_cols, nrows=2000)
                tgt_sample = self._load_file_cols(target_path, target_delimiter,
                                                   common_cols, nrows=2000)
                from core.key_detector import detect_composite_key
                kd = detect_composite_key(
                    src_df=src_sample,
                    tgt_df=tgt_sample,
                    object_name=object_name or self._infer_object(source_path),
                )
                join_keys        = kd.join_keys
                detection_method = kd.detection_method
                confidence       = kd.confidence
                logger.info(
                    f"Auto-detected join keys: {join_keys} "
                    f"(method={detection_method}, confidence={confidence})"
                )
            except Exception as e:
                join_keys        = []
                detection_method = "failed"
                confidence       = "low"
                logger.warning(f"Key detection failed: {e}")

            if not join_keys:
                return ValidationResult(
                    source_file=source_path, target_file=target_path,
                    total_source_records=0, total_target_records=0,
                    records_matched=0, records_only_in_source=0,
                    records_only_in_target=0,
                    common_columns=common_cols,
                    key_detection_method="none",
                    key_confidence="low",
                    errors=[
                        "No join key could be detected automatically. "
                        "Please select join keys manually from the dashboard before validating. "
                        f"Common columns available: {common_cols}"
                    ]
                )

        # ── Step 3: Load only needed columns ──────────────────────────────────
        if self.field_map:
            field_map  = {s.upper(): t.upper() for s, t in self.field_map.items()}
            needed_src = set(field_map.keys()) | set(join_keys)
            needed_tgt = set(field_map.values()) | set(join_keys)
        else:
            field_map  = {c: c for c in common_cols}
            needed_src = set(src_hdrs)
            needed_tgt = set(tgt_hdrs)

        try:
            src_df = self._load_file_cols(source_path, source_delimiter,
                                          list(needed_src & set(src_hdrs)))
        except Exception as e:
            return ValidationResult(
                source_file=source_path, target_file=target_path,
                total_source_records=0, total_target_records=0,
                records_matched=0, records_only_in_source=0, records_only_in_target=0,
                common_columns=common_cols,
                errors=[f"Cannot load source file: {e}"]
            )
        try:
            tgt_df = self._load_file_cols(target_path, target_delimiter,
                                          list(needed_tgt & set(tgt_hdrs)))
        except Exception as e:
            return ValidationResult(
                source_file=source_path, target_file=target_path,
                total_source_records=len(src_df), total_target_records=0,
                records_matched=0, records_only_in_source=len(src_df),
                records_only_in_target=0,
                common_columns=common_cols,
                errors=[f"Cannot load target file: {e}"]
            )

        # ── Step 4: Normalise join key values ─────────────────────────────────
        for k in join_keys:
            if k in src_df.columns:
                src_df[k] = src_df[k].astype(str).str.strip().str.upper()
            if k in tgt_df.columns:
                tgt_df[k] = tgt_df[k].astype(str).str.strip().str.upper()

        # ── Step 5: Build composite key column ────────────────────────────────
        # e.g. MATNR=1234, KSCHL=PB00, EKORG=CNG1 → "1234||PB00||CNG1"
        SEP = "||"
        src_df[_CK] = src_df[join_keys].astype(str).agg(SEP.join, axis=1)
        tgt_df[_CK] = tgt_df[join_keys].astype(str).agg(SEP.join, axis=1)

        # ── Step 6: Key set counts ─────────────────────────────────────────────
        src_keys = set(src_df[_CK].unique())
        tgt_keys = set(tgt_df[_CK].unique())

        # Duplicate detection
        src_dup_mask  = src_df[_CK].duplicated(keep=False)
        tgt_dup_mask  = tgt_df[_CK].duplicated(keep=False)
        dup_src_count = int(src_dup_mask.sum())
        dup_tgt_count = int(tgt_dup_mask.sum())
        dup_samples   = []
        if dup_src_count > 0:
            dup_samples = src_df[src_dup_mask].head(10)[join_keys].to_dict("records")

        # ── Step 7: Separate join keys from regular fields ────────────────────
        # Join key fields ARE validated (value comparison across matched records)
        # but displayed with a special "KEY" badge in the dashboard.
        jk_set         = set(join_keys)
        key_field_map  = {k: k for k in join_keys
                          if k in src_df.columns and k in tgt_df.columns}
        data_field_map = {s: t for s, t in field_map.items()
                          if s not in jk_set and t not in jk_set
                          and s in src_df.columns and t in tgt_df.columns}

        # ── Step 8: Numeric detection (data fields only) ──────────────────────
        pairs   = list(data_field_map.items())
        tol_map = self._detect_numeric_mapped(src_df, tgt_df, pairs)
        tol_map.update(self.tolerance_overrides)

        # ── Step 9: Mapping report ─────────────────────────────────────────────
        from core.field_labels import get_label
        mapping_report = MappingReport(
            join_keys=join_keys,
            join_key_labels={k: get_label(k, self.custom_labels) for k in join_keys},
            matched_fields=list(data_field_map.keys()),
            source_only_fields=sorted(set(src_hdrs) - set(tgt_hdrs) - jk_set),
            target_only_fields=sorted(set(tgt_hdrs) - set(src_hdrs) - jk_set),
            numeric_fields=sorted(tol_map.keys()),
            tolerance_map=tol_map,
            total_source_cols=len(src_hdrs),
            total_target_cols=len(tgt_hdrs),
            selected_fields=self.selected_fields,
            pass_threshold=self.pass_threshold,
        )

        # ── Step 10: Validate each field ──────────────────────────────────────
        field_results = []

        # First: validate join key fields (marked as key fields)
        for src_col in join_keys:
            tgt_col = key_field_map.get(src_col, src_col)
            fr = self._validate_field_fast(
                src_df, tgt_df, src_col, tgt_col,
                join_keys, None, max_mismatch_rows,
                is_key_field=True,
            )
            if fr:
                field_results.append(fr)

        # Then: validate data fields
        for src_col, tgt_col in data_field_map.items():
            tolerance = tol_map.get(src_col)
            fr = self._validate_field_fast(
                src_df, tgt_df, src_col, tgt_col,
                join_keys, tolerance, max_mismatch_rows,
            )
            if fr:
                field_results.append(fr)

        return ValidationResult(
            source_file=source_path,
            target_file=target_path,
            total_source_records=len(src_df),
            total_target_records=len(tgt_df),
            records_matched=len(src_keys & tgt_keys),
            records_only_in_source=len(src_keys - tgt_keys),
            records_only_in_target=len(tgt_keys - src_keys),
            mapping=mapping_report,
            field_results=field_results,
            join_keys=join_keys,
            key_detection_method=detection_method,
            key_confidence=confidence,
            duplicate_src=dup_src_count,
            duplicate_tgt=dup_tgt_count,
            duplicate_key_samples=dup_samples,
            common_columns=common_cols,
        )

    # ── Per-field validation (2-column merge on composite key) ────────────────

    def _validate_field_fast(
        self,
        src_df, tgt_df, src_col, tgt_col, join_keys,
        tolerance, max_rows, is_key_field=False,
    ):
        if src_col not in src_df.columns or tgt_col not in tgt_df.columns:
            return None

        # Merge only CK + this field — tiny, fast even for lakh-scale
        if src_col == tgt_col:
            m = src_df[[_CK, src_col]].merge(
                tgt_df[[_CK, tgt_col]], on=_CK, how="inner",
                suffixes=("_src", "_tgt"))
            sv_col = src_col + "_src"
            tv_col = tgt_col + "_tgt"
        else:
            m = src_df[[_CK, src_col]].merge(
                tgt_df[[_CK, tgt_col]], on=_CK, how="inner")
            sv_col = src_col
            tv_col = tgt_col

        total = len(m)
        if total == 0:
            return FieldResult(
                field_source=src_col, field_target=tgt_col,
                field_label=self._label(src_col), total_records=0,
                matched=0, mismatched=0, missing_in_target=0, missing_in_source=0,
                is_numeric=(tolerance is not None), is_key_field=is_key_field,
                tolerance_used=tolerance, pass_threshold=self.pass_threshold,
            )

        sv = m[sv_col]
        tv = m[tv_col]

        # Null detection
        NULL_VALS  = {"", "nan", "none", "null"}
        sv_null    = sv.isna() | sv.astype(str).str.strip().str.lower().isin(NULL_VALS)
        tv_null    = tv.isna() | tv.astype(str).str.strip().str.lower().isin(NULL_VALS)
        both_null  = sv_null & tv_null
        miss_src_m = sv_null & ~tv_null
        miss_tgt_m = ~sv_null & tv_null
        valid      = ~sv_null & ~tv_null

        sv_v = sv[valid]
        tv_v = tv[valid]

        # Comparison
        if tolerance is not None:
            sv_f = pd.to_numeric(sv_v.astype(str).str.replace(",", ".", regex=False), errors="coerce")
            tv_f = pd.to_numeric(tv_v.astype(str).str.replace(",", ".", regex=False), errors="coerce")
            ok            = sv_f.notna() & tv_f.notna()
            diff          = (sv_f - tv_f).abs()
            match_mask    = ok & (diff <= tolerance)
            mismatch_mask = ok & (diff > tolerance)
        else:
            sv_s = sv_v.astype(str).str.strip().str.upper()
            tv_s = tv_v.astype(str).str.strip().str.upper()
            match_mask    = sv_s == tv_s
            mismatch_mask = ~match_mask

        matched        = int(both_null.sum()) + int(match_mask.sum())
        miss_src_count = int(miss_src_m.sum())
        miss_tgt_count = int(miss_tgt_m.sum())
        mismatched     = int(mismatch_mask.sum())

        # Decode composite key to readable label
        def _ck_label(ck_val):
            parts = str(ck_val).split("||")
            if len(join_keys) == 1:
                return str(parts[0])
            return " / ".join(
                f"{k}={v}" for k, v in zip(join_keys, parts)
            )

        # Build mismatch list with to_dict (fast)
        mismatches = []

        if miss_src_count > 0:
            for r in m[miss_src_m].head(max_rows).to_dict("records"):
                mismatches.append({
                    "material":     _ck_label(r[_CK]),
                    "source_value": "(blank)",
                    "target_value": str(r.get(tv_col, "")),
                    "issue":        "Missing in source",
                })

        if miss_tgt_count > 0 and len(mismatches) < max_rows:
            for r in m[miss_tgt_m].head(max_rows - len(mismatches)).to_dict("records"):
                mismatches.append({
                    "material":     _ck_label(r[_CK]),
                    "source_value": str(r.get(sv_col, "")),
                    "target_value": "(blank)",
                    "issue":        "Missing in target",
                })

        if mismatched > 0 and len(mismatches) < max_rows:
            mdf = m[valid][mismatch_mask].head(max_rows - len(mismatches))
            if tolerance is not None:
                sv_mf   = sv_f[mismatch_mask].reindex(mdf.index)
                tv_mf   = tv_f[mismatch_mask].reindex(mdf.index)
                diff_mf = diff[mismatch_mask].reindex(mdf.index)
                for idx, r in zip(mdf.index, mdf.to_dict("records")):
                    mismatches.append({
                        "material":     _ck_label(r[_CK]),
                        "source_value": round(float(sv_mf.get(idx, 0)), 4),
                        "target_value": round(float(tv_mf.get(idx, 0)), 4),
                        "issue":        f"Numeric delta (tolerance ±{tolerance})",
                    })
            else:
                for r in mdf.to_dict("records"):
                    mismatches.append({
                        "material":     _ck_label(r[_CK]),
                        "source_value": str(r.get(sv_col, "")),
                        "target_value": str(r.get(tv_col, "")),
                        "issue":        "Value mismatch",
                    })

        return FieldResult(
            field_source=src_col, field_target=tgt_col,
            field_label=self._label(src_col),
            total_records=total, matched=matched, mismatched=mismatched,
            missing_in_target=miss_tgt_count, missing_in_source=miss_src_count,
            is_numeric=(tolerance is not None), is_key_field=is_key_field,
            tolerance_used=tolerance, pass_threshold=self.pass_threshold,
            mismatch_details=mismatches,
        )

    # ── File loading ──────────────────────────────────────────────────────────

    def _read_headers(self, path, delimiter=","):
        import csv
        p = str(path)
        if p.lower().endswith((".xlsx", ".xls")):
            import openpyxl
            wb   = openpyxl.load_workbook(p, read_only=True, data_only=True)
            ws   = wb.active
            cols = [str(c.value).strip().upper()
                    for c in next(ws.iter_rows(max_row=1)) if c.value]
            wb.close()
            return cols
        with open(p, encoding="utf-8-sig") as f:
            return [c.strip().upper()
                    for c in next(csv.reader(f, delimiter=delimiter))]

    def _load_file_cols(self, path, delimiter, needed_cols, nrows=None):
        p           = str(path)
        all_headers = self._read_headers(p, delimiter)
        needed_set  = set(c.upper() for c in (needed_cols or []))
        load_cols   = [c for c in all_headers if not needed_set or c in needed_set]

        kwargs = {"dtype": str}
        if nrows:
            kwargs["nrows"] = nrows

        if p.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(p, usecols=load_cols or None, **kwargs)
        else:
            df = pd.read_csv(
                p, delimiter=delimiter, encoding="utf-8-sig",
                low_memory=False, usecols=load_cols or None,
                na_filter=False, **kwargs,
            )
        df.columns = df.columns.str.strip().str.upper()
        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].str.strip()
        return df

    def _load_file(self, path, delimiter):
        return self._load_file_cols(path, delimiter, [])

    @staticmethod
    def _infer_object(file_path):
        from pathlib import Path
        return Path(file_path).stem.upper()

    # ── Numeric detection ─────────────────────────────────────────────────────

    def _detect_numeric_mapped(self, src_df, tgt_df, pairs):
        numeric_cols = {}
        for src_col, tgt_col in pairs:
            if src_col not in src_df.columns or tgt_col not in tgt_df.columns:
                continue
            sv = src_df[src_col].dropna().head(self.numeric_sample_rows)
            tv = tgt_df[tgt_col].dropna().head(self.numeric_sample_rows)
            if not len(sv) or not len(tv):
                continue

            def rate(s):
                return (pd.to_numeric(
                    s.astype(str).str.replace(",", ".", regex=False),
                    errors="coerce").notna().sum() / len(s))

            if rate(sv) >= self.numeric_threshold and rate(tv) >= self.numeric_threshold:
                tol = self.tolerance_overrides.get(src_col)
                if tol is None:
                    all_v = pd.to_numeric(
                        pd.concat([sv, tv]).astype(str)
                        .str.replace(",", ".", regex=False), errors="coerce"
                    ).dropna().abs()
                    tol = self._scale_tolerance(float(all_v.median()) if len(all_v) else 0.0)
                numeric_cols[src_col] = tol
        return numeric_cols

    @staticmethod
    def _scale_tolerance(v):
        if v == 0:     return 0.0
        elif v < 1:    return 0.0001
        elif v < 10:   return 0.001
        elif v < 1000: return 0.01
        else:          return 0.1

    # ── Legacy methods for backwards compatibility ────────────────────────────

    def _detect_join_key(self, src_df, tgt_df):
        if self.manual_join_keys:
            return self.manual_join_keys[0]
        common = set(src_df.columns) & set(tgt_df.columns)
        for pk in ["MATNR","KUNNR","LIFNR","SAKNR","BELNR","EBELN","VBELN","ANLN1"]:
            if pk in common:
                return pk
        return sorted(common)[0] if common else None

    def _normalise_key(self, df, col):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()
        return df

    def _build_field_map(self, src_df, tgt_df, join_key):
        from core.field_labels import get_label
        src_cols = set(src_df.columns) - {join_key}
        tgt_cols = set(tgt_df.columns) - {join_key}
        if self.field_map:
            vm  = {s.upper(): t.upper() for s, t in self.field_map.items()
                   if s.upper() in src_df.columns and t.upper() in tgt_df.columns}
            tol = self._detect_numeric_mapped(src_df, tgt_df, list(vm.items()))
        else:
            cols = sorted(src_cols & tgt_cols)
            vm   = {c: c for c in cols}
            tol  = self._detect_numeric_mapped(src_df, tgt_df, list(vm.items()))
        tol.update(self.tolerance_overrides)
        jks = self.manual_join_keys or [join_key]
        return vm, MappingReport(
            join_keys=jks,
            join_key_labels={k: get_label(k, self.custom_labels) for k in jks},
            matched_fields=list(vm.keys()),
            source_only_fields=sorted(src_cols - tgt_cols),
            target_only_fields=sorted(tgt_cols - src_cols),
            numeric_fields=sorted(tol.keys()), tolerance_map=tol,
            total_source_cols=len(src_df.columns),
            total_target_cols=len(tgt_df.columns),
            selected_fields=self.selected_fields,
            pass_threshold=self.pass_threshold,
        )
