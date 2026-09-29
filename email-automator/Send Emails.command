#!/bin/bash
cd "$(dirname "$0")"
echo
echo "============================================"
echo "  UBC BioMod Email Sender"
echo "============================================"
echo
if ! command -v python3 >/dev/null 2>&1; then
  echo "Could not find Python on this computer."
  echo "Open README.md in this folder and follow the Install Python steps."
  echo
  read -r -p "Press Enter to close this window."
  exit 1
fi
python3 email.py
echo
read -r -p "Press Enter to close this window."
