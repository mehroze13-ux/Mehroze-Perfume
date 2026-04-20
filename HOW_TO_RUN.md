# How to Run the Nykaa Scraper

## What it does
Scrapes all 135 pages of fragrances from Nykaa and saves the data to **nykaa_products.xlsx**  
(opens in Microsoft Excel or Google Sheets).

## Data collected per product
| Column  | Example |
|---------|---------|
| brand   | Fogg |
| name    | Fogg Scent Xpressio Eau De Parfum |
| price   | ₹299 |
| mrp     | ₹450 |
| rating  | 4.3 |
| link    | https://www.nykaa.com/... |
| image   | https://... |
| page_no | 1 |

---

## Step 1 — Install Python (one-time only)

### Windows
1. Go to https://www.python.org/downloads/
2. Click **Download Python** (the big yellow button)
3. Run the installer — **tick "Add Python to PATH"** before clicking Install

### Mac
- Python is usually already installed.  
  If not: https://www.python.org/downloads/

---

## Step 2 — Run the scraper

### Windows (easiest)
1. Download this whole folder to your computer
2. Double-click **`setup_and_run.bat`**
3. Wait — it will install everything automatically and start scraping

### Mac / Linux
1. Open Terminal
2. `cd` into this folder
3. Run: `bash setup_and_run.sh`

---

## How long will it take?
About **30–60 minutes** for all 135 pages.  
You'll see progress printed in the window like:
```
  Page   1/135  →   24 products  (total 24)
  Page   2/135  →   24 products  (total 48)
  ...
```

---

## Output file
When finished, **`nykaa_products.xlsx`** will appear in the same folder.  
Open it in Excel or upload it to Google Sheets.
