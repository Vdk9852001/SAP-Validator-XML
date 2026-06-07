+# Genpact SAP Migration Validator

Validates SAP post-load Excel extracts against XML migration templates.
Supports all 40+ SAP objects. Works on Windows, Mac, Linux.

## Quick Start

### Python Desktop App
```bash
# 1. Install dependencies (one time)
pip install openpyxl lxml anthropic

# 2. Run
python sap_validator.py
```

### React Web App (runs in browser)
```bash
# 1. Install Node.js from https://nodejs.org
# 2. Create project
npm create vite@latest sap-react -- --template react
cd sap-react
npm install xlsx

# 3. Replace src/App.jsx with sap_validator.jsx
# 4. Start
npm run dev
# Open http://localhost:5173
```

## Features
- Auto-maps Excel headers → SAP field codes (Work Center, Plant, etc.)
- Composite join key detection (ARBPL+WERKS, MATNR+WERKS, etc.)
- Field-level diff with PASS/FAIL/WARN per record
- Smart Search & Learn — find unmatched records across both files,
  confirm the correct join field, tool learns it for future validation
- AI column mapping via Claude API (optional, needs API key)
- Excel report export: Summary, All Records, Mismatches, XML Only sheets

## API Key (optional, for AI mapping)
Set environment variable OR click "API Key" button in the app:
```bash
# Windows
set ANTHROPIC_API_KEY=sk-ant-...

# Mac/Linux
export ANTHROPIC_API_KEY=sk-ant-...
```
Get a key at https://console.anthropic.com

## Supported SAP Objects
Work Center / Resource, Material Master, Customer, Vendor, BOM,
Routing, Cost Center, G/L Account, Profit Center, Asset, and more.
Any XML + Excel pair works — no hardcoded object configs.
