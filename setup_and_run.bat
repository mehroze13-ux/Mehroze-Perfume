@echo off
:: Nykaa Scraper - Windows Setup & Run Script
:: Double-click this file to install everything and start scraping.

echo ================================================
echo   Nykaa Fragrance Scraper - Windows Setup
echo ================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed.
    echo Please download and install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/3] Installing Python packages...
pip install playwright pandas openpyxl --quiet

echo [2/3] Downloading browser for scraping...
python -m playwright install chromium

echo [3/3] Starting scraper... (this will take 30-60 minutes for 135 pages)
echo.
python scrape_nykaa.py

echo.
echo All done! Check nykaa_products.xlsx in this folder.
pause
