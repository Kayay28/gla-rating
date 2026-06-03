#!/bin/bash
echo "================================================"
echo "  GLA Teacher Performance Rating System"
echo "  Starting Word Generation Server..."
echo "================================================"
echo ""
cd "$(dirname "$0")"
pip install flask flask-cors python-docx --quiet --break-system-packages 2>/dev/null || pip install flask flask-cors python-docx --quiet 2>/dev/null
python3 server.py
