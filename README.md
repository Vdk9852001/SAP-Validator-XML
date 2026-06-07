# Genpact SAP Migration Validator — V5

Post-load validation tool for SAP 4.7 → S/4HANA migrations.
Validates source extracts against target exports — field by field, record by record.

## Quick Start

```bash
# Install
pip install -r requirements.txt

# Run (always from the project root)
python dashboard/app.py

# Open browser
http://localhost:5000
```

## What's New in V5

- **LTMC XML validation** — upload SAP Migration Cockpit SpreadsheetML exports directly
- **Dynamic XML parsing** — works for any SAP object (Product, Customer, Vendor, etc.)
- **Label resolver** — post-load files with friendly names ("Customer Number") auto-mapped to SAP codes (KUNNR)
- **Composite join keys** — user-selectable from the dashboard, no hardcoding
- **Key fields in results** — join key fields shown in the field-level results table with KEY badge
- **Target-authoritative validation** — all target file columns are validated
- **Performance** — usecols loading, 2-column merge per field, to_dict() mismatch collection

## Folder Structure

```
sap_validator_v5/
  core/
    key_detector.py      # Dynamic composite key detection (49 SAP objects)
    validator.py         # Validation engine — composite key, vectorised
    field_mapper.py      # Alias + fuzzy field mapping (SAP 4.7 ↔ S/4HANA)
    field_labels.py      # SAP field code → friendly label (180+ fields)
    ltmc_parser.py       # SAP LTMC SpreadsheetML XML parser
    label_resolver.py    # Friendly column name → SAP code resolver
    object_config.py     # 14 SAP object definitions
    reporter.py          # Excel report generator
  config/
    field_labels.json    # 180+ SAP field labels
    field_aliases.json   # Cross-name mappings (NAME1→NAMORG1, LAND1→COUNTRY)
  dashboard/
    app.py               # Flask server — all routes
    templates/
      dashboard.html     # Single-page dashboard UI
  data/
    source/              # Drop source files here
    target/              # Drop target files here
  templates/             # Field selection templates (CSV/XLSX/TXT)
  reports/               # Auto-generated Excel reports
  requirements.txt
  run.bat                # Windows launcher
  run.sh                 # Mac/Linux launcher
  README.md
```

## How to Validate

### Standard CSV/Excel Files
1. Upload source file (SAP 4.7 extract)
2. Upload target file (S/4HANA export)
3. If filenames differ, use **Pairs** to link them
4. Click **Scan Now**
5. Select join keys from **Edit join keys** button

### SAP LTMC XML Templates
1. Click **LTMC Validate** in the header
2. Upload your SAP Migration Cockpit `.xml` export
3. All worksheets are parsed automatically (any SAP object)
4. Select a worksheet + post-load extract file
5. Preview column matching (friendly names resolved automatically)
6. Set join keys and click **Run Validation**

## Settings

| Setting | Description |
|---|---|
| Pass Threshold | Minimum match % for PASS (default 100%) |
| Field Templates | Upload CSV with field names to restrict which fields are validated |
| Custom Labels | Upload CSV to override field display names |
| Custom Mapping | Upload CSV of SOURCE,TARGET pairs to override auto-mapping |
| Join Keys | Select per-table composite key from the dashboard |

## Troubleshooting

**Buttons not working** — JavaScript error. Open browser console (F12) and check for `SyntaxError`. Replace `dashboard.html`.

**Only N fields validated** — wrong object name or missing aliases. Check Activity Log.

**Source only: 134,936 records** — join key issue. Click "Edit join keys" and check uniqueness scores.

**LTMC XML not parsing** — ensure file is a SAP Migration Cockpit SpreadsheetML export (not a regular Excel file saved as XML).

## Architecture

```
File Upload
    ↓
ltmc_parser.py          (XML only) → DataFrame per sheet
    ↓
field_mapper.py         alias + fuzzy mapping → field_map {src: tgt}
label_resolver.py       friendly name → SAP code (LTMC only)
    ↓
key_detector.py         composite key detection → join_keys
    ↓
validator.py            __CK__ composite column → 2-col merge per field
    ↓
reporter.py             Excel report
    ↓
dashboard.html          live results every 4 seconds
```
