#!/bin/bash
echo ""
echo " Genpact SAP Migration Validator V5"
echo " ==================================="
echo ""
cd "$(dirname "$0")"
python3 dashboard/app.py
