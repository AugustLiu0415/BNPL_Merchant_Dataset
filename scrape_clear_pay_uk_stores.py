import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


URL = "https://www.clearpay.co.uk/en-GB/stores/retailer-pages"
BNPL_PROVIDER = "Clearpay"
COUNTRY = "United Kingdom"

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "Data_Raw"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_DIR / "clearpay_uk_retailer_pages.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "clearpay_uk_retailer_pages.csv"
DEBUG_HTML = OUTPUT_DIR / "clearpay_retailer_pages_debug.html"


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_name(name: str) -> str:
    name = clean_text(name).lower()
    name = re.sub(r"[^a-z0-9]+", "", name)
    return name


def extract_slug_from_url(url: str) -> str:
    parsed = urlparse(url)
    path = parsed.path.rstrip("/")
    parts = [p for p in path.split("/") if p]

    if "stores" in parts:
        idx = parts.index("stores")
        if idx + 1 < len(parts):
            return parts[idx + 1]

    return ""


def is_non_merchant(name: str, url: str = "") -> bool:
    if not name:
        return True

    name_clean = clean_text(name)
    name_lower = name_clean.lower()
    url_lower = (url or "").lower()

    blocked_names = {
        "get the app",
        "how it works",
        "get the facts",
        "invite a friend",
        "help",
        "login",
        "sign up",
        "all categories",
        "most popular",
        "on sale",
        "new",
        "support small",
        "shop in-store",
        "support",
        "account login",
        "security",
        "contact us",
        "hardship",
        "single use payments",
        "gift cards",
        "clearpay card",
        "clearpay app",
        "google play",
        "app store",
        "resources",
        "clearpay access",
        "retailer resources",
        "clearpay api",
        "clearpay for business",
        "partner program",
        "location",
        "united states",
        "australia",
        "new zealand",
        "united kingdom",
        "canada - english",
        "canada - français",
        "about",
        "careers",
        "media",
        "investors",
        "general terms",
        "website terms",
        "cookies notice",
        "privacy",
        "cookie preferences",
        "modern slavery statement",
        "retailer pages",
        "categories",
        "all stores",
    }

    if name_lower in blocked_names:
        return True

    if len(name_clean) > 80:
        return True

    if re.fullmatch(r"\d+", name_clean):
        return True

    if name_lower.startswith("http") or "@" in name_lower:
        return True

    blocked_url_parts = [
        "/help",
        "/security",
        "/contact",
        "/careers",
        "/media",
        "/investors",
        "/terms",
        "/privacy",
        "/cookies",
        "/partner-program",
        "/for-retailers",
        "/clearpay-api",
        "/retailer-resources",
        "/account",
        "portal.clearpay.co.uk",
        "developers.clearpay.co.uk",
        "play.google.com",
        "apps.apple.com",
        "instagram.com",
        "linkedin.com",
        "facebook.com",
        "x.com",
    ]

    if any(part in url_lower for part in blocked_url_parts):
        return True

    return False


def is_likely_clearpay_store_link(name: str, full_url: str) -> bool:
    if is_non_merchant(name, full_url):
        return False

    parsed = urlparse(full_url)
    path = parsed.path.rstrip("/")

    if parsed.netloc not in {"www.clearpay.co.uk", "clearpay.co.uk"}:
        return False

    if not path.startswith("/en-GB/stores/"):
        return False

    if path == "/en-GB/stores/retailer-pages":
        return False

    return True


def extract_links_from_page(page) -> list[dict]:
    rows = []
    links = page.locator("a").all()

    print(f"Total links found on page: {len(links)}")

    for link in links:
        try:
            raw_text = link.inner_text(timeout=2000)
            company_name = clean_text(raw_text)

            href = link.get_attribute("href") or ""
            full_url = urljoin(URL, href)

            if not is_likely_clearpay_store_link(company_name, full_url):
                continue

            rows.append(
                {
                    "company_name": company_name,
                    "merchant_slug": extract_slug_from_url(full_url),
                    "merchant_url": full_url,
                    "bnpl_provider": BNPL_PROVIDER,
                    "country": COUNTRY,
                    "source_page": URL,
                    "availability_type": "Clearpay UK retailer pages",
                    "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

        except Exception:
            continue

    return rows


def extract_list_items_from_page(page) -> list[dict]:
    rows = []
    items = page.locator("li").all()

    print(f"Total list items found on page: {len(items)}")

    for item in items:
        try:
            company_name = clean_text(item.inner_text(timeout=2000))

            if is_non_merchant(company_name):
                continue

            rows.append(
                {
                    "company_name": company_name,
                    "merchant_slug": "",
                    "merchant_url": "",
                    "bnpl_provider": BNPL_PROVIDER,
                    "country": COUNTRY,
                    "source_page": URL,
                    "availability_type": "Clearpay UK retailer pages list item",
                    "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

        except Exception:
            continue

    return rows


def scrape_clearpay_retailer_pages() -> pd.DataFrame:
    all_rows = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        page = browser.new_page(
            viewport={"width": 1440, "height": 1600},
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )

        try:
            print(f"Opening: {URL}")
            page.goto(URL, wait_until="domcontentloaded", timeout=120000)
            page.wait_for_timeout(5000)

            DEBUG_HTML.write_text(page.content(), encoding="utf-8")

            link_rows = extract_links_from_page(page)
            list_rows = extract_list_items_from_page(page)

            print(f"Extracted link rows: {len(link_rows)}")
            print(f"Extracted list-item rows: {len(list_rows)}")

            all_rows.extend(link_rows)
            all_rows.extend(list_rows)

        except PlaywrightTimeoutError:
            print("Timeout while opening Clearpay retailer pages.")
        except Exception as e:
            print(f"Error while scraping Clearpay retailer pages: {e}")

        browser.close()

    seen = set()
    deduped_rows = []

    for row in all_rows:
        norm = normalize_name(row["company_name"])

        if not norm:
            continue

        if norm in seen:
            continue

        seen.add(norm)
        deduped_rows.append(row)

    df = pd.DataFrame(deduped_rows)

    if not df.empty:
        df = df.sort_values("company_name").reset_index(drop=True)

    return df


def save_outputs(df: pd.DataFrame) -> None:
    if df.empty:
        print("No Clearpay retailers extracted.")
        print(f"Check debug HTML: {DEBUG_HTML}")
        return

    df.to_excel(OUTPUT_EXCEL, index=False)
    df.to_csv(OUTPUT_CSV, index=False)

    print(f"\nSaved Excel file to: {OUTPUT_EXCEL}")
    print(f"Saved CSV file to: {OUTPUT_CSV}")
    print(f"Total unique Clearpay retailer rows extracted: {len(df)}")

    print("\nPreview:")
    print(df.head(40).to_string(index=False))


if __name__ == "__main__":
    clearpay_df = scrape_clearpay_retailer_pages()
    save_outputs(clearpay_df)