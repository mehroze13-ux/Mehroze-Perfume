"""
Nykaa Fragrance Scraper  –  uses undetected-chromedriver (your real Chrome)
URL: https://www.nykaa.com/fragrance/c/53?page_no=1&sort=popularity&formulation_filter=228729
Output: nykaa_products.xlsx

Install:  pip install undetected-chromedriver selenium pandas openpyxl
Run:      python scrape_nykaa.py
"""

import time
import sys
import json
import random
import os
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# ── CONFIG ────────────────────────────────────────────────────────────────────
TOTAL_PAGES     = 135
OUTPUT_FILE     = "nykaa_products.xlsx"
CHECKPOINT_FILE = "nykaa_checkpoint.json"
DELAY_MIN       = 2.5
DELAY_MAX       = 4.0

BASE_URL = (
    "https://www.nykaa.com/fragrance/c/53"
    "?page_no={page}&sort=popularity"
    "&search_redirection=True&formulation_filter=228729"
)
# ─────────────────────────────────────────────────────────────────────────────

EXTRACT_JS = """
(function() {
    var results = [];

    function txt(el, sels) {
        for (var i = 0; i < sels.length; i++) {
            var node = el.querySelector(sels[i]);
            if (node && node.innerText.trim()) return node.innerText.trim();
        }
        return "";
    }

    // Find product cards
    var CARD_SELS = [
        '[data-at="sku-card"]',
        '[class*="product-card"]',
        '[class*="productCard"]',
        '[class*="ProductCard"]',
        '.css-lrlqb8',
        '.css-d5z3ro',
        '.css-1xchqfd',
    ];

    var cards = [];
    for (var s = 0; s < CARD_SELS.length; s++) {
        cards = Array.from(document.querySelectorAll(CARD_SELS[s]));
        if (cards.length > 0) break;
    }

    if (cards.length === 0) {
        var seen = new Set();
        document.querySelectorAll('a[href*="/p/"]').forEach(function(a) {
            var card = a.closest('li, article, div[class]') || a;
            if (!seen.has(card)) { seen.add(card); cards.push(card); }
        });
    }

    for (var i = 0; i < cards.length; i++) {
        var card = cards[i];

        var name = txt(card, [
            '[data-at="sku-name"]', '[class*="product-title"]',
            '[class*="productTitle"]', '[class*="ProductTitle"]',
            'h3', 'h4', 'p[class*="name"]',
        ]);
        var brand = txt(card, [
            '[data-at="sku-brand"]', '[class*="brand-name"]',
            '[class*="brandName"]', '[class*="BrandName"]',
            'p[class*="brand"]',
        ]);
        var price = txt(card, [
            '[data-at="sku-selling-price"]', '[class*="selling-price"]',
            '[class*="sellingPrice"]', '[class*="discounted-price"]',
        ]);
        var mrp = txt(card, [
            '[data-at="sku-mrp"]', '[class*="mrp"]',
            '[class*="strike"]', '[class*="original-price"]',
            's', 'del',
        ]);
        var discount = txt(card, [
            '[data-at="sku-discount"]', '[class*="discount"]', 'span[class*="off"]',
        ]);

        // Rating: try to extract value + count from wrapper
        var ratingValue = "", ratingCount = "";
        var ratingWrap = card.querySelector(
            '[data-at="sku-rating"], [class*="rating-wrapper"], [class*="ratingWrapper"], [class*="star-rating"]'
        );
        if (ratingWrap) {
            var rf = ratingWrap.innerText.trim();
            var m = rf.match(/([\d.]+)\s*[(\xb7]\s*([\d,]+)/);
            if (m) {
                ratingValue = m[1];
                ratingCount = m[2].replace(/,/g, "");
            } else {
                ratingValue = rf.replace(/[^0-9.]/g, "");
            }
        }
        if (!ratingValue) ratingValue = txt(card, ['[class*="rating-value"]','[class*="ratingValue"]','[class*="avg-rating"]']);
        if (!ratingCount) ratingCount = txt(card, ['[class*="rating-count"]','[class*="ratingCount"]','[class*="review-count"]']).replace(/[^0-9]/g,"");

        var olfactory = txt(card, [
            '[class*="fragrance-notes"]', '[class*="fragranceNotes"]',
            '[class*="olfactory"]', '[class*="notes"]',
        ]);
        var volume = txt(card, [
            '[class*="volume"]', '[class*="Volume"]',
            '[class*="size"]', '[class*="variant"]',
        ]);

        var linkEl = card.querySelector('a[href*="/p/"]') || card.querySelector('a');
        var link = linkEl ? linkEl.href : "";
        var imgEl = card.querySelector('img');
        var image = imgEl ? (imgEl.dataset.src || imgEl.src || "") : "";

        if (name || link) {
            results.push({
                brand: brand, name: name, price: price, mrp: mrp,
                discount: discount, rating_value: ratingValue,
                rating_count: ratingCount, olfactory_notes: olfactory,
                volume: volume, link: link, image: image,
            });
        }
    }
    return results;
})();
"""


def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE) as f:
            return json.load(f)
    return {"last_page": 0, "products": []}


def save_checkpoint(last_page, products):
    with open(CHECKPOINT_FILE, "w") as f:
        json.dump({"last_page": last_page, "products": products}, f)


def make_driver():
    options = uc.ChromeOptions()
    options.add_argument("--window-size=1280,800")
    options.add_argument("--lang=en-IN")
    options.add_argument("--disable-popup-blocking")
    # Visible browser — Cloudflare cannot block a real visible Chrome session
    driver = uc.Chrome(options=options, use_subprocess=True)
    return driver


def scrape_all():
    checkpoint = load_checkpoint()
    start_page   = checkpoint["last_page"] + 1
    all_products = checkpoint["products"]

    if start_page > 1:
        print(f"  Resuming from page {start_page} ({len(all_products)} products already)")

    print("=" * 64)
    print("  Nykaa Fragrance Scraper  –  135 pages")
    print(f"  Output → {OUTPUT_FILE}")
    print("=" * 64)

    driver = make_driver()

    try:
        # Visit homepage first so Cloudflare sets its cookies
        print("  Opening Nykaa homepage to warm up session...")
        driver.get("https://www.nykaa.com/")
        time.sleep(5)

        # Dismiss any popup/overlay (cookie consent, location, etc.)
        for dismiss_sel in ['button[id*="close"]', 'button[class*="close"]',
                            'button[class*="dismiss"]', '[class*="modal"] button']:
            try:
                btn = driver.find_element(By.CSS_SELECTOR, dismiss_sel)
                btn.click()
                time.sleep(1)
            except Exception:
                pass

        for page_num in range(start_page, TOTAL_PAGES + 1):
            url = BASE_URL.format(page=page_num)

            success = False
            for attempt in range(3):
                try:
                    driver.get(url)
                    time.sleep(3)  # let JS render

                    # Scroll slowly to trigger lazy-load
                    for scroll_pct in [0.25, 0.5, 0.75, 1.0]:
                        driver.execute_script(
                            f"window.scrollTo(0, document.body.scrollHeight * {scroll_pct})"
                        )
                        time.sleep(0.5)

                    # On first page, save HTML for debugging
                    if page_num == start_page:
                        with open("debug_page1.html", "w", encoding="utf-8") as f:
                            f.write(driver.page_source)
                        print("  [debug] Saved page 1 HTML to debug_page1.html")

                    success = True
                    break

                except WebDriverException as e:
                    print(f"  Page {page_num:>3}  retry {attempt+1}/3: {str(e)[:80]}")
                    if attempt < 2:
                        time.sleep(6 * (attempt + 1))
                    else:
                        raise

            if not success:
                continue

            products = driver.execute_script(EXTRACT_JS)
            if products is None:
                products = []

            for item in products:
                item["page_no"] = page_num
            all_products.extend(products)

            print(
                f"  Page {page_num:>3}/{TOTAL_PAGES}"
                f"  →  {len(products):>3} products"
                f"  (total {len(all_products)})"
            )

            if page_num % 10 == 0:
                save_checkpoint(page_num, all_products)
                print(f"  [checkpoint saved at page {page_num}]")

            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    except KeyboardInterrupt:
        print("\n  Interrupted — saving checkpoint...")
        save_checkpoint(page_num - 1, all_products)
    finally:
        driver.quit()

    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)

    if not all_products:
        print("\nNo products found. Try removing --headless=new from the script to debug.")
        sys.exit(1)

    cols = [
        "page_no", "brand", "name", "price", "mrp", "discount",
        "rating_value", "rating_count", "olfactory_notes",
        "volume", "link", "image",
    ]
    df = pd.DataFrame(all_products)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]
    df.drop_duplicates(subset=["link"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone!  Saved {len(df)} unique products -> '{OUTPUT_FILE}'")


if __name__ == "__main__":
    try:
        scrape_all()
    except Exception as e:
        import traceback
        print("\n" + "=" * 64)
        print("  FATAL ERROR:")
        print("=" * 64)
        traceback.print_exc()
    input("\nPress Enter to close...")
