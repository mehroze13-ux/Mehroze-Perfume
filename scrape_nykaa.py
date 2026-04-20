"""
Nykaa Fragrance Scraper
URL: https://www.nykaa.com/fragrance/c/53?page_no=1&sort=popularity&formulation_filter=228729
Output: nykaa_products.xlsx  (opens in Excel / Google Sheets)

Run:  python scrape_nykaa.py
"""

import time
import sys
import json
import random
import os
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── CONFIG ────────────────────────────────────────────────────────────────────
TOTAL_PAGES    = 135
OUTPUT_FILE    = "nykaa_products.xlsx"
CHECKPOINT_FILE = "nykaa_checkpoint.json"
DELAY_MIN      = 2.0   # min seconds between pages
DELAY_MAX      = 3.5   # max seconds between pages (randomised)

BASE_URL = (
    "https://www.nykaa.com/fragrance/c/53"
    "?page_no={page}&sort=popularity"
    "&search_redirection=True&formulation_filter=228729"
)
# ─────────────────────────────────────────────────────────────────────────────

EXTRACT_JS = """
() => {
    const results = [];

    // ── helpers ──────────────────────────────────────────────────────────────
    const txt = (el, sels) => {
        for (const s of sels) {
            const node = el.querySelector(s);
            if (node && node.innerText.trim()) return node.innerText.trim();
        }
        return "";
    };
    const attr = (el, sel, at) => {
        const node = el.querySelector(sel);
        return node ? (node.getAttribute(at) || "").trim() : "";
    };

    // ── find product cards ────────────────────────────────────────────────────
    const CARD_SELS = [
        '[data-at="sku-card"]',
        '[class*="product-card"]',
        '[class*="productCard"]',
        '[class*="ProductCard"]',
        '.css-lrlqb8',
        '.css-d5z3ro',
        '.css-1xchqfd',
        '[class*="NykaaProduct"]',
    ];

    let cards = [];
    for (const s of CARD_SELS) {
        cards = Array.from(document.querySelectorAll(s));
        if (cards.length > 0) break;
    }

    // fallback – anchor-based dedup
    if (cards.length === 0) {
        const seen = new Set();
        document.querySelectorAll('a[href*="/p/"]').forEach(a => {
            const card = a.closest('li, article, div[class]') || a;
            if (!seen.has(card)) { seen.add(card); cards.push(card); }
        });
    }

    for (const card of cards) {

        // ── core fields ──────────────────────────────────────────────────────
        const name  = txt(card, [
            '[data-at="sku-name"]', '[class*="product-title"]',
            '[class*="productTitle"]', '[class*="ProductTitle"]',
            'h3', 'h4', 'p[class*="name"]',
        ]);
        const brand = txt(card, [
            '[data-at="sku-brand"]', '[class*="brand-name"]',
            '[class*="brandName"]', '[class*="BrandName"]',
            'p[class*="brand"]',
        ]);
        const price = txt(card, [
            '[data-at="sku-selling-price"]', '[class*="selling-price"]',
            '[class*="sellingPrice"]', '[class*="discounted-price"]',
            '[class*="DiscountedPrice"]', 'span[class*="price"]:not([class*="strike"])',
        ]);
        const mrp = txt(card, [
            '[data-at="sku-mrp"]', '[class*="mrp"]',
            '[class*="strike"]', '[class*="original-price"]',
            's', 'del',
        ]);

        // ── rating: value + count ─────────────────────────────────────────────
        // Nykaa usually renders "4.2 (1,234)" in adjacent spans or a wrapper
        let ratingValue = "";
        let ratingCount = "";

        const ratingWrap = card.querySelector(
            '[data-at="sku-rating"], [class*="rating-wrapper"], ' +
            '[class*="ratingWrapper"], [class*="RatingWrapper"], ' +
            '[class*="star-rating"], [class*="starRating"]'
        );
        if (ratingWrap) {
            const ratingFull = ratingWrap.innerText.trim();
            // try to split "4.2 (1,234)" or "4.2 · 1,234"
            const m = ratingFull.match(/([\d.]+)\s*[\(·]\s*([\d,]+)/);
            if (m) {
                ratingValue = m[1];
                ratingCount = m[2].replace(/,/g, "");
            } else {
                // just a plain number
                ratingValue = ratingFull.replace(/[^0-9.]/g, "");
            }
        }

        // fallback selectors for value and count separately
        if (!ratingValue) {
            ratingValue = txt(card, [
                '[class*="rating-value"]', '[class*="ratingValue"]',
                '[class*="avg-rating"]',   '[class*="avgRating"]',
            ]);
        }
        if (!ratingCount) {
            ratingCount = txt(card, [
                '[class*="rating-count"]', '[class*="ratingCount"]',
                '[class*="review-count"]', '[class*="reviewCount"]',
                '[class*="num-rating"]',   '[class*="numRating"]',
            ]).replace(/[^0-9]/g, "");
        }

        // ── discount ──────────────────────────────────────────────────────────
        const discount = txt(card, [
            '[data-at="sku-discount"]', '[class*="discount"]',
            '[class*="Discount"]', 'span[class*="off"]',
        ]);

        // ── olfactory / fragrance notes ───────────────────────────────────────
        // Nykaa sometimes shows tags like "Woody | Floral" or "Top notes: …"
        const olfactory = txt(card, [
            '[class*="fragrance-notes"]', '[class*="fragranceNotes"]',
            '[class*="olfactory"]',        '[class*="notes"]',
            '[class*="tag"]',              '[class*="Tag"]',
        ]);

        // ── volume / size ─────────────────────────────────────────────────────
        const volume = txt(card, [
            '[class*="volume"]', '[class*="Volume"]',
            '[class*="size"]',   '[class*="Size"]',
            '[class*="variant"]','[class*="Variant"]',
        ]);

        // ── link & image ──────────────────────────────────────────────────────
        const linkEl = card.querySelector('a[href*="/p/"]') || card.querySelector('a');
        const link   = linkEl ? linkEl.href : "";
        const imgEl  = card.querySelector('img');
        const image  = imgEl ? (imgEl.dataset.src || imgEl.src || "") : "";

        // only keep cards that have meaningful data
        if (name || link) {
            results.push({
                brand, name, price, mrp, discount,
                rating_value: ratingValue,
                rating_count: ratingCount,
                olfactory_notes: olfactory,
                volume, link, image,
            });
        }
    }
    return results;
}
"""


def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE) as f:
            return json.load(f)
    return {"last_page": 0, "products": []}


def save_checkpoint(last_page, products):
    with open(CHECKPOINT_FILE, "w") as f:
        json.dump({"last_page": last_page, "products": products}, f)


def scrape_all():
    checkpoint = load_checkpoint()
    start_page  = checkpoint["last_page"] + 1
    all_products = checkpoint["products"]

    if start_page > 1:
        print(f"  Resuming from page {start_page} ({len(all_products)} products already collected)")

    print("=" * 64)
    print("  Nykaa Fragrance Scraper  –  135 pages")
    print(f"  Output → {OUTPUT_FILE}")
    print("=" * 64)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1440, "height": 900},
            locale="en-IN",
            timezone_id="Asia/Kolkata",
        )
        # hide automation fingerprint
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        pg = context.new_page()

        # Skip heavy assets to speed things up
        pg.route(
            "**/*.{png,jpg,jpeg,gif,webp,svg,ico,woff,woff2,ttf,eot}",
            lambda route: route.abort(),
        )
        for pattern in ["**/analytics*", "**/gtm*", "**/clevertap*",
                         "**/hotjar*", "**/facebook*", "**/doubleclick*"]:
            pg.route(pattern, lambda route: route.abort())

        for page_num in range(start_page, TOTAL_PAGES + 1):
            url = BASE_URL.format(page=page_num)
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=60_000)

                # Wait for product cards to appear
                try:
                    pg.wait_for_selector(
                        '[data-at="sku-card"], [class*="product-card"], a[href*="/p/"]',
                        timeout=25_000,
                    )
                except PWTimeout:
                    pass  # still try extracting

                # Scroll to trigger lazy-load
                pg.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
                time.sleep(0.5)
                pg.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(0.8)

                products = pg.evaluate(EXTRACT_JS)
                for item in products:
                    item["page_no"] = page_num
                all_products.extend(products)

                print(
                    f"  Page {page_num:>3}/{TOTAL_PAGES}"
                    f"  →  {len(products):>3} products"
                    f"  (total {len(all_products)})"
                )

                # Checkpoint every 10 pages
                if page_num % 10 == 0:
                    save_checkpoint(page_num, all_products)
                    print(f"  [checkpoint saved at page {page_num}]")

                time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

            except Exception as exc:
                print(f"  Page {page_num:>3}  !! ERROR: {exc}")
                save_checkpoint(page_num - 1, all_products)
                time.sleep(6)

        browser.close()

    # ── Clean up checkpoint ────────────────────────────────────────────────────
    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)

    # ── Save final output ──────────────────────────────────────────────────────
    if not all_products:
        print("\nNo products found. The site may have blocked the scraper.")
        print("Try setting headless=False in the script to debug.")
        sys.exit(1)

    cols = [
        "page_no", "brand", "name", "price", "mrp", "discount",
        "rating_value", "rating_count", "olfactory_notes",
        "volume", "link", "image",
    ]
    df = pd.DataFrame(all_products)
    # Ensure all expected columns exist
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]
    df.drop_duplicates(subset=["link"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone!  Saved {len(df)} unique products -> '{OUTPUT_FILE}'")
    print("  Open this file in Microsoft Excel or Google Sheets.")


if __name__ == "__main__":
    scrape_all()
