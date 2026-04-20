# How to Run the Nykaa Fragrance Scraper

## What it does
Scrapes all 135 pages of fragrances from Nykaa and saves the data to **nykaa_products.xlsx**  
(opens in Microsoft Excel or Google Sheets).

## Data collected per product
| Column           | Example                                        |
|------------------|------------------------------------------------|
| page_no          | 1                                              |
| brand            | Fogg                                           |
| name             | Fogg Scent Xpressio Eau De Parfum              |
| price            | ₹299                                           |
| mrp              | ₹450                                           |
| discount         | 34% off                                        |
| rating_value     | 4.3                                            |
| rating_count     | 1234                                           |
| olfactory_notes  | Woody, Floral (shown if available on listing)  |
| volume           | 100 ml                                         |
| link             | https://www.nykaa.com/...                      |
| image            | https://...                                    |

> **Note:** `olfactory_notes` and `volume` are only populated when Nykaa displays them on the listing page.  
> For complete fragrance notes, the product detail pages would need separate scraping.

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
3. Wait — it installs everything automatically and starts scraping

### Mac / Linux
1. Open Terminal
2. `cd` into this folder
3. Run: `bash setup_and_run.sh`

---

## How long will it take?
About **45–70 minutes** for all 135 pages (randomised delay to avoid getting blocked).  
Progress is printed live:
```
  Page   1/135  →   24 products  (total 24)
  Page   2/135  →   24 products  (total 48)
  ...
  [checkpoint saved at page 10]
  ...
```
A **checkpoint file** (`nykaa_checkpoint.json`) is saved every 10 pages so you can resume if the script is interrupted — just run it again and it will pick up where it left off.

---

## Output file
When finished, **`nykaa_products.xlsx`** will appear in the same folder.  
Open it in Excel or upload it to Google Sheets.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "No products found" | Nykaa may have detected the scraper. Open `scrape_nykaa.py` and set `headless=False` to watch the browser. Try again after a few minutes. |
| Script stops mid-way | Run it again — it will resume from the last checkpoint automatically. |
| Missing data in some columns | Some fields (olfactory notes, volume) are only visible on individual product pages, not the listing. |
