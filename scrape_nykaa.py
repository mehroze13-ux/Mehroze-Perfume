"""
Nykaa Fragrance Scraper
URL: https://www.nykaa.com/fragrance/c/53?page_no=1&sort=popularity&formulation_filter=228729
Output: nykaa_products.xlsx  (opens in Excel / Google Sheets)

Run:  python scrape_nykaa.py
"""

import time
import sys
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── CONFIG ────────────────────────────────────────────────────────────────────
TOTAL_PAGES = 135
OUTPUT_FILE  = "nykaa_products.xlsx"
DELAY        = 1.5   # seconds between pages (be polite to the server)

BASE_URL = (
    "https://www.nykaa.com/fragrance/c/53"
    "?page_no={page}&sort=popularity"
    "&search_redirection=True&formulation_filter=228729"
)
# ─────────────────────────────────────────────────────────────────────────────


def extract_products(page):
    """Pull product data from the current page DOM."""
    return page.evaluate("""
    () => {
        const results = [];

        // Nykaa renders cards with several possible class patterns
        const CARD_SELECTORS = [
            '[data-at="sku-card"]',
            '[class*="product-card"]',
            '[class*="productCard"]',
            '[class*="ProductCard"]',
            '.css-lrlqb8',
            '.css-d5z3ro',
        ];

        let cards = [];
        for (const sel of CARD_SELECTORS) {
            cards = document.querySelectorAll(sel);
            if (cards.length > 0) break;
        }

        // Fallback: collect every anchor that links to a /p/ product page
        if (cards.length === 0) {
            const seen = new Set();
            document.querySelectorAll('a[href*="/p/"]').forEach(a => {
                const card = a.closest('li, article, div[class]') || a;
                if (!seen.has(card)) {
                    seen.add(card);
                    cards = [...cards, card];
                }
            });
        }

        for (const card of cards) {
            const txt = (sel) => {
                const el = card.querySelector(sel);
                return el ? el.innerText.trim() : "";
            };

            const name = txt('[data-at="sku-name"], [class*="product-title"], [class*="productTitle"], h3, h4');
            const brand = txt('[data-at="sku-brand"], [class*="brand"], [class*="Brand"]');
            const price = txt('[data-at="sku-selling-price"], [class*="selling-price"], [class*="sellingPrice"], [class*="discounted-price"]');
            const mrp   = txt('[data-at="sku-mrp"], [class*="strike"], [class*="mrp"], [class*="original-price"]');
            const rating = txt('[data-at="sku-rating"], [class*="rating"], [class*="Rating"]');

            const linkEl  = card.querySelector('a[href*="/p/"]') || card.querySelector('a');
            const link    = linkEl ? linkEl.href : "";

            const imgEl   = card.querySelector('img');
            const image   = imgEl ? (imgEl.dataset.src || imgEl.getAttribute('src') || "") : "";

            // Only keep cards that have at least a name or a link
            if (name || link) {
                results.push({ brand, name, price, mrp, rating, link, image });
            }
        }
        return results;
    }
    """)


def scrape_all():
    all_products = []

    print("=" * 62)
    print("  Nykaa Fragrance Scraper  –  135 pages")
    print("  Output →", OUTPUT_FILE)
    print("=" * 62)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,   # change to False to watch the browser
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1440, "height": 900},
        )
        pg = context.new_page()

        # Speed-up: skip images, fonts, and analytics
        pg.route(
            "**/*.{png,jpg,jpeg,gif,webp,svg,ico,woff,woff2,ttf,eot}",
            lambda route: route.abort(),
        )
        pg.route("**/analytics*", lambda route: route.abort())
        pg.route("**/gtm*",       lambda route: route.abort())

        for page_num in range(1, TOTAL_PAGES + 1):
            url = BASE_URL.format(page=page_num)
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=60_000)

                # Wait for product cards
                try:
                    pg.wait_for_selector(
                        '[data-at="sku-card"], [class*="product-card"], a[href*="/p/"]',
                        timeout=20_000,
                    )
                except PWTimeout:
                    pass  # page may still have products even if selector timed out

                # Scroll down so lazy-loaded items appear
                pg.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(1.0)

                products = extract_products(pg)
                for item in products:
                    item["page_no"] = page_num
                all_products.extend(products)

                print(
                    f"  Page {page_num:>3}/{TOTAL_PAGES}"
                    f"  →  {len(products):>3} products"
                    f"  (total {len(all_products)})"
                )
                time.sleep(DELAY)

            except Exception as exc:
                print(f"  Page {page_num:>3}  !! ERROR: {exc}")
                time.sleep(4)

        browser.close()

    # ── Save ──────────────────────────────────────────────────────────────
    if not all_products:
        print("\nNo products found. The site may have blocked the scraper.")
        print("Try setting headless=False in the script to debug.")
        sys.exit(1)

    df = pd.DataFrame(all_products, columns=["brand","name","price","mrp","rating","link","image","page_no"])
    df.drop_duplicates(subset=["link"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✓ Done!  Saved {len(df)} unique products → '{OUTPUT_FILE}'")
    print("  Open this file in Microsoft Excel or Google Sheets.")


if __name__ == "__main__":
    scrape_all()
