#!/bin/bash
echo "=== Goodwood Screening System - Installer ==="
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
echo ""
echo "Instalasi selesai. Jalankan: bash start_mac_linux.sh"
