#!/bin/bash
# Nykaa Scraper - Mac/Linux Setup & Run Script

echo "================================================"
echo "  Nykaa Fragrance Scraper - Mac/Linux Setup"
echo "================================================"
echo

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Mac: Install from https://www.python.org/downloads/"
    echo "Linux: sudo apt install python3 python3-pip"
    exit 1
fi

echo "[1/3] Installing Python packages..."
pip3 install playwright pandas openpyxl --quiet

echo "[2/3] Downloading browser for scraping..."
python3 -m playwright install chromium

echo "[3/3] Starting scraper... (this will take 30-60 minutes for 135 pages)"
echo
python3 scrape_nykaa.py

echo
echo "All done! Check nykaa_products.xlsx in this folder."
