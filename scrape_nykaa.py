"""
Nykaa Fragrance Scraper – Full Product Details (Two-Phase)
URL: https://www.nykaa.com/fragrance/c/53?page_no=1&sort=popularity&formulation_filter=228729

Phase 1 – Visit all 135 listing pages and collect every product URL.
Phase 2 – Visit each product page individually to get the FULL untruncated
           title, brand, price, MRP, discount, rating, review count,
           description, ingredients, how-to-use, size, and images.

A checkpoint file (nykaa_links.json) is saved after Phase 1 so you can
resume Phase 2 without repeating Phase 1 if the script is interrupted.
Progress is also saved every 50 products to nykaa_products_full.xlsx.

Run:  python scrape_nykaa.py
"""

import json
import os
import sys
import time

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

try:
    from playwright_stealth import stealth_sync
    _STEALTH = True
except ImportError:
    _STEALTH = False

# ── CONFIG ────────────────────────────────────────────────────────────────────
TOTAL_PAGES    = 135
OUTPUT_FILE    = "nykaa_products_full.xlsx"
LINKS_FILE     = "nykaa_links.json"   # checkpoint after Phase 1
SAVE_EVERY     = 50                   # save partial results every N products
DELAY_LISTING  = 1.5                  # seconds between listing pages
DELAY_DETAIL   = 2.0                  # seconds between product detail pages

BASE_URL = (
    "https://www.nykaa.com/fragrance/c/53"
    "?page_no={page}&sort=popularity"
    "&search_redirection=True&formulation_filter=228729"
)
# ─────────────────────────────────────────────────────────────────────────────

BLOCKED_EXTENSIONS = (
    "**/*.{png,jpg,jpeg,gif,webp,svg,ico,woff,woff2,ttf,eot,otf,mp4,mp3,avi}"
)
BLOCKED_PATTERNS = [
    "**/analytics**",
    "**/gtm**",
    "**/hotjar**",
    "**/facebook**",
    "**/doubleclick**",
    "**/googlesyndication**",
    "**/adservice**",
]


# ── PHASE 1 HELPERS ───────────────────────────────────────────────────────────

LISTING_JS = """
() => {
    const seen  = new Set();
    const links = [];

    // Primary: known card selectors
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
        const found = document.querySelectorAll(sel);
        if (found.length > 0) { cards = Array.from(found); break; }
    }

    // Fallback: any anchor pointing to a product (/p/) page
    if (cards.length === 0) {
        document.querySelectorAll('a[href*="/p/"]').forEach(a => {
            const card = a.closest('li, article, div[class]') || a;
            if (!seen.has(card)) { seen.add(card); cards.push(card); }
        });
    }

    for (const card of cards) {
        const a = card.querySelector('a[href*="/p/"]') || card.querySelector('a');
        if (!a) continue;
        const href = a.href;
        if (href && !seen.has(href)) {
            seen.add(href);
            links.push(href);
        }
    }
    return links;
}
"""


# ── PHASE 2 HELPERS ───────────────────────────────────────────────────────────

DETAIL_JS = """
() => {
    // ── Helper ──────────────────────────────────────────────────────────────
    const first = (...selectors) => {
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el) {
                const t = el.innerText ? el.innerText.trim() : (el.textContent || "").trim();
                if (t) return t;
            }
        }
        return "";
    };

    // ── 1. JSON-LD structured data (most reliable, used for SEO) ──────────
    let name = "", brand = "", price = "", mrp = "", rating = "",
        reviews = "", description = "";

    document.querySelectorAll('script[type="application/ld+json"]').forEach(s => {
        try {
            const raw  = JSON.parse(s.textContent);
            const items = Array.isArray(raw) ? raw : [raw];
            for (const d of items) {
                if (d["@type"] === "Product") {
                    if (!name)        name        = d.name || "";
                    if (!brand)       brand       = (d.brand && (d.brand.name || d.brand)) || "";
                    if (!description) description = d.description || "";
                    const offers = d.offers || {};
                    if (!price) price = String(offers.price || offers.lowPrice || "");
                    if (!mrp)   mrp   = String(offers.highPrice || "");
                    const agg = d.aggregateRating || {};
                    if (!rating)  rating  = String(agg.ratingValue || "");
                    if (!reviews) reviews = String(agg.reviewCount || agg.ratingCount || "");
                }
            }
        } catch(e) {}
    });

    // ── 2. DOM fallbacks for anything still missing ───────────────────────
    if (!name)
        name = first('h1', '[data-at="pdp-product-name"]',
                     '[class*="product-name"]', '[class*="productName"]',
                     '[class*="product-title"]', '[class*="productTitle"]');

    if (!brand)
        brand = first('[data-at="pdp-brand-name"]', '[class*="brand-name"]',
                      '[class*="brandName"]', 'a[href*="/brand/"]',
                      '[class*="brand"]');

    if (!price)
        price = first('[data-at="sku-selling-price"]', '[class*="selling-price"]',
                      '[class*="sellingPrice"]', '[class*="discounted-price"]',
                      '[class*="final-price"]');

    if (!mrp)
        mrp = first('[data-at="sku-mrp"]', '[class*="strike"]',
                    '[class*="mrp"]', '[class*="MRP"]',
                    '[class*="original-price"]', '[class*="base-price"]');

    if (!rating)
        rating = first('[data-at="sku-rating"]', '[class*="average-rating"]',
                       '[class*="averageRating"]', '[class*="rating-count"]');

    if (!reviews)
        reviews = first('[class*="review-count"]', '[class*="reviewCount"]',
                        '[class*="total-review"]', '[class*="ratings-count"]');

    // Discount
    const discount = first('[class*="discount"]', '[class*="Discount"]',
                           '[class*="offer-tag"]', '[class*="save"]');

    // Size / volume – multiple strategies
    let size = "";

    // Strategy 1: dedicated size/variant selectors
    size = first(
        '[class*="size-label"]', '[class*="sizeLabel"]',
        '[class*="selected-size"]', '[class*="variant-label"]',
        '[class*="pack-size"]',   '[class*="packSize"]',
        '[class*="volume"]',      '[class*="capacity"]',
        '[class*="net-quantity"]','[class*="netQuantity"]'
    );

    // Strategy 2: selected / active variant button (e.g. "100 ml" pill)
    if (!size) {
        const activeBtn = document.querySelector(
            '[class*="variant"] [class*="selected"], [class*="variant"] [class*="active"],' +
            '[class*="size"] button[class*="selected"], [class*="size"] button[class*="active"],' +
            'button[aria-selected="true"], li[aria-selected="true"]'
        );
        if (activeBtn) size = activeBtn.innerText.trim();
    }

    // Strategy 3: all variant buttons – collect them all as "50ml | 100ml | 200ml"
    if (!size) {
        const btns = document.querySelectorAll(
            '[class*="variant"] button, [class*="size-option"], [class*="sizeOption"],' +
            '[class*="pack-option"], [class*="packOption"]'
        );
        const opts = Array.from(btns)
            .map(b => b.innerText.trim())
            .filter(t => /\d/.test(t));
        if (opts.length) size = opts.join(" | ");
    }

    // Strategy 4: regex extract from product name / page text (e.g. "100 ml", "50g")
    if (!size) {
        const h1Text = (document.querySelector('h1') || {}).innerText || "";
        const m = h1Text.match(/\d+(\.\d+)?\s*(ml|ML|g|gm|GM|kg|KG|oz|OZ|L|litre|liter)/i);
        if (m) size = m[0].trim();
    }

    // ── 3. Long-text fields ───────────────────────────────────────────────
    const longText = (selectors) => {
        for (const sel of selectors) {
            const els = document.querySelectorAll(sel);
            for (const el of els) {
                const t = el.innerText ? el.innerText.trim() : "";
                if (t.length > 30) return t;
            }
        }
        return "";
    };

    if (!description)
        description = longText([
            '[class*="description"]', '[class*="about-product"]',
            '[id*="description"]',   '[class*="product-detail"]',
            '[class*="productDescription"]', '[class*="prod-desc"]',
        ]);

    const ingredients = longText([
        '[class*="ingredient"]', '[id*="ingredient"]',
        '[class*="Ingredient"]', '[class*="key-ingredient"]',
    ]);

    const how_to_use = longText([
        '[class*="how-to"]',  '[class*="howTo"]',
        '[id*="how-to"]',     '[class*="usage"]',
        '[class*="direction"]', '[class*="application"]',
    ]);

    // ── 4. Images ─────────────────────────────────────────────────────────
    const imgEls = document.querySelectorAll(
        '[class*="image-container"] img, [class*="product-image"] img, ' +
        '[class*="gallery"] img, [class*="pdp"] img, ' +
        '[class*="carousel"] img, [class*="slider"] img'
    );
    const imgSrcs = Array.from(imgEls)
        .map(img => img.dataset.src || img.getAttribute("src") || "")
        .filter(src => src && !src.startsWith("data:") && src.startsWith("http"));
    const images = [...new Set(imgSrcs)].slice(0, 8).join(" | ");

    return { name, brand, price, mrp, discount, rating, reviews,
             description, ingredients, how_to_use, size, images };
}
"""


# ── SCRAPER FUNCTIONS ─────────────────────────────────────────────────────────

def setup_page(context):
    pg = context.new_page()
    if _STEALTH:
        stealth_sync(pg)
    pg.route(BLOCKED_EXTENSIONS, lambda r: r.abort())
    for pat in BLOCKED_PATTERNS:
        pg.route(pat, lambda r: r.abort())
    return pg


def goto_with_retry(pg, url, retries=3):
    """Navigate with retries and increasing back-off on failure."""
    for attempt in range(1, retries + 1):
        try:
            pg.goto(url, wait_until="domcontentloaded", timeout=60_000)
            return True
        except Exception as exc:
            if attempt == retries:
                raise
            wait = attempt * 5
            print(f"    retry {attempt}/{retries-1} after {wait}s  ({exc})")
            time.sleep(wait)
    return False


def phase1_collect_links(pg):
    print("\n── Phase 1: Collecting product links from 135 listing pages ──")
    all_links = []
    seen = set()

    for page_num in range(1, TOTAL_PAGES + 1):
        url = BASE_URL.format(page=page_num)
        try:
            goto_with_retry(pg, url)
            try:
                pg.wait_for_selector('a[href*="/p/"]', timeout=20_000)
            except PWTimeout:
                pass

            # Scroll in steps to trigger lazy-loaded cards
            pg.evaluate("window.scrollTo(0, document.body.scrollHeight / 3)")
            time.sleep(0.4)
            pg.evaluate("window.scrollTo(0, document.body.scrollHeight * 2 / 3)")
            time.sleep(0.4)
            pg.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(0.8)

            links = pg.evaluate(LISTING_JS)
            new_links = [l for l in links if l not in seen]
            seen.update(new_links)
            all_links.extend(new_links)

            print(
                f"  Page {page_num:>3}/{TOTAL_PAGES}"
                f"  →  {len(new_links):>3} new links"
                f"  (total {len(all_links)})"
            )
            time.sleep(DELAY_LISTING)

        except Exception as exc:
            print(f"  Page {page_num:>3}  !! ERROR: {exc}")
            time.sleep(4)

    with open(LINKS_FILE, "w") as f:
        json.dump(all_links, f)
    print(f"\n  Checkpoint saved: {len(all_links)} links → '{LINKS_FILE}'")
    return all_links


def phase2_scrape_details(pg, links):
    print(f"\n── Phase 2: Scraping full details for {len(links)} products ──")
    products = []

    for i, link in enumerate(links, 1):
        try:
            goto_with_retry(pg, link)
            try:
                pg.wait_for_selector("h1", timeout=20_000)
            except PWTimeout:
                pass

            # Scroll to load lazy sections (description, ingredients, etc.)
            pg.evaluate("window.scrollTo(0, document.body.scrollHeight / 3)")
            time.sleep(0.5)
            pg.evaluate("window.scrollTo(0, document.body.scrollHeight * 2 / 3)")
            time.sleep(0.5)
            pg.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(0.5)
            pg.evaluate("window.scrollTo(0, 0)")

            details = pg.evaluate(DETAIL_JS)
            details["link"] = link

            products.append(details)
            print(
                f"  [{i:>4}/{len(links)}]"
                f"  {(details.get('brand') or '?')[:18]:<18}"
                f"  {(details.get('name') or '?')[:55]}"
            )

            # Incremental save every SAVE_EVERY products
            if i % SAVE_EVERY == 0:
                _save(products, partial=True)

            time.sleep(DELAY_DETAIL)

        except Exception as exc:
            print(f"  [{i:>4}/{len(links)}]  !! ERROR: {link[:60]}")
            print(f"              {exc}")
            products.append({"link": link, "error": str(exc)})
            time.sleep(4)

    return products


def _save(products, partial=False):
    cols = [
        "brand", "name", "price", "mrp", "discount",
        "rating", "reviews", "description", "ingredients",
        "how_to_use", "size", "images", "link",
    ]
    df = pd.DataFrame(products)
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    df = df[cols]
    df.drop_duplicates(subset=["link"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)
    if partial:
        print(f"  [checkpoint] {len(df)} products saved → '{OUTPUT_FILE}'")


# ── MAIN ─────────────────────────────────────────────────────────────────────

def scrape_all():
    print("=" * 70)
    print("  Nykaa Fragrance Scraper  –  Full Product Details (Two-Phase)")
    print()
    print("  Phase 1  →  135 listing pages  →  collect product URLs")
    print("  Phase 2  →  each product page  →  full title, price, description,")
    print("              ingredients, how-to-use, rating, images, and more")
    print()
    print(f"  Output   →  {OUTPUT_FILE}")
    print("=" * 70)

    if _STEALTH:
        print("  [stealth] playwright-stealth is active.")
    else:
        print("  [stealth] playwright-stealth not found – using built-in patches.")
        print("            For better results: pip install playwright-stealth")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--disable-http2",
                "--disable-web-security",
                "--lang=en-US,en",
                "--window-size=1440,900",
            ],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1440, "height": 900},
            locale="en-US",
            java_script_enabled=True,
            extra_http_headers={
                "Accept-Language": "en-US,en;q=0.9",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
                "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
                "Upgrade-Insecure-Requests": "1",
            },
        )
        # Comprehensive stealth patches injected before every page load
        context.add_init_script("""
            // 1. Hide webdriver flag
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            try { delete navigator.__proto__.webdriver; } catch(e){}

            // 2. Mock chrome runtime (absent in headless = bot signal)
            if (!window.chrome) {
                window.chrome = {runtime: {}, loadTimes: ()=>{}, csi: ()=>{}, app: {}};
            }

            // 3. Non-zero plugins list (0 plugins = headless bot)
            Object.defineProperty(navigator, 'plugins', {
                get: () => { const p = [1,2,3,4,5]; p.__proto__ = PluginArray.prototype; return p; }
            });

            // 4. Supported MIME types
            Object.defineProperty(navigator, 'mimeTypes', {
                get: () => { const m = [1]; m.__proto__ = MimeTypeArray.prototype; return m; }
            });

            // 5. Real language list
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});

            // 6. Platform
            Object.defineProperty(navigator, 'platform', {get: () => 'Win32'});

            // 7. Hardware concurrency (0 = bot signal)
            Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});

            // 8. Device memory
            Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});

            // 9. Notification permission bypass
            try {
                const orig = navigator.permissions.query.bind(navigator.permissions);
                navigator.permissions.query = p =>
                    p.name === 'notifications'
                        ? Promise.resolve({state: Notification.permission})
                        : orig(p);
            } catch(e) {}

            // 10. Remove Playwright-specific globals
            ['__playwright', '__pw_manual', '__pw_task_queue'].forEach(k => {
                try { delete window[k]; } catch(e) {}
            });
        """)
        pg = setup_page(context)

        # ── Phase 1 ──────────────────────────────────────────────────────────
        if os.path.exists(LINKS_FILE):
            with open(LINKS_FILE) as f:
                all_links = json.load(f)
            print(f"\n  Found '{LINKS_FILE}' with {len(all_links)} links.")
            print("  Skipping Phase 1.  (Delete the file to re-run Phase 1.)\n")
        else:
            all_links = phase1_collect_links(pg)

        if not all_links:
            print("\nNo product links found. The site may have blocked the scraper.")
            print("Try setting headless=False in browser.launch() to debug.")
            sys.exit(1)

        # ── Phase 2 ──────────────────────────────────────────────────────────
        products = phase2_scrape_details(pg, all_links)
        browser.close()

    # ── Final save ────────────────────────────────────────────────────────────
    if not products:
        print("\nNo product details collected.")
        sys.exit(1)

    _save(products)

    cols = [
        "brand", "name", "price", "mrp", "discount",
        "rating", "reviews", "description", "ingredients",
        "how_to_use", "size", "images", "link",
    ]
    df = pd.DataFrame(products)
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    df = df[cols]
    df.drop_duplicates(subset=["link"], keep="first", inplace=True)

    print(f"\n✓  Done!  {len(df)} unique products saved → '{OUTPUT_FILE}'")
    print("   Open in Microsoft Excel or Google Sheets.")
    print()
    print("   Columns in the output file:")
    for col in cols:
        print(f"     • {col}")


if __name__ == "__main__":
    scrape_all()
