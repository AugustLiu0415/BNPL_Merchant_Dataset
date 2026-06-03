import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


BASE_URL = "https://zip.co"
DIRECTORY_BASE_URL = "https://zip.co/us/store/directory"
BNPL_PROVIDER = "Zip"
COUNTRY = "United States"

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "Data_Raw"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_DIR / "zip_us_stores.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "zip_us_stores.csv"
DEBUG_DIR = OUTPUT_DIR / "zip_debug_pages"
DEBUG_DIR.mkdir(exist_ok=True)

DIRECTORY_SECTIONS = list("abcdefghijklmnopqrstuvwxyz") + ["0-9"]


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

    if "store" in parts:
        idx = parts.index("store")
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
        "customers",
        "business",
        "for business",
        "how it works",
        "shop stores",
        "shop categories",
        "insights",
        "help",
        "products",
        "affiliate",
        "resource",
        "pricing",
        "change region from united states",
        "sign in",
        "sign up",
        "open mobile navigation",
        "customer sign in",
        "become a zip merchant",
        "merchant sign in",
        "merchant resources",
        "documentation",
        "api reference",
        "pci dss compliance",
        "about",
        "careers",
        "contact us",
        "investors",
        "press",
        "support",
        "privacy policy",
        "system status",
        "terms of service",
        "licenses",
        "vulnerability disclosure program",
        "home",
        "featured",
        "see all",
        "facebook",
        "twitter",
        "instagram",
        "linkedin",
        "0-9",
    }

    if name_lower in blocked_names:
        return True

    # alphabet navigation links
    if re.fullmatch(r"[a-z]", name_lower):
        return True

    # empty/pure number links
    if re.fullmatch(r"\d+", name_clean):
        return True

    # remove long legal/footer text chunks
    if len(name_clean) > 100:
        return True

    blocked_url_parts = [
        "/us/business",
        "/us/for-business",
        "/us/how-it-works",
        "/us/categories",
        "/us/insights",
        "/us/app",
        "/us/chrome",
        "/us/contact",
        "/us/about",
        "/us/careers",
        "/us/press",
        "/us/privacy",
        "/us/terms",
        "/us/licenses",
        "help.us.zip.co",
        "customer.us.zip.co",
        "merchant.us.zip.co",
        "docs.us.zip.co",
        "status.us.zip.co",
        "facebook.com",
        "twitter.com",
        "instagram.com",
        "linkedin.com",
    ]

    if any(part in url_lower for part in blocked_url_parts):
        return True

    return False


def is_likely_zip_store_link(name: str, full_url: str) -> bool:
    if is_non_merchant(name, full_url):
        return False

    parsed = urlparse(full_url)
    path = parsed.path.rstrip("/")

    if parsed.netloc not in {"zip.co", "www.zip.co"}:
        return False

    # Keep store links like /us/store/nike
    if not path.startswith("/us/store/"):
        return False

    # Exclude directory pages like /us/store/directory/a
    if path.startswith("/us/store/directory"):
        return False

    if path == "/us/store":
        return False

    return True


def scrape_one_directory_page(page, section: str) -> list[dict]:
    url = f"{DIRECTORY_BASE_URL}/{section}"
    print(f"\nOpening Zip directory section '{section}': {url}")

    page.goto(url, wait_until="domcontentloaded", timeout=120000)
    page.wait_for_timeout(3000)

    debug_path = DEBUG_DIR / f"zip_directory_{section}.html"
    debug_path.write_text(page.content(), encoding="utf-8")

    rows = []
    links = page.locator("a").all()
    print(f"Section {section}: total links found = {len(links)}")

    for link in links:
        try:
            raw_text = link.inner_text(timeout=2000)
            company_name = clean_text(raw_text)

            href = link.get_attribute("href") or ""
            full_url = urljoin(BASE_URL, href)

            if not is_likely_zip_store_link(company_name, full_url):
                continue

            rows.append(
                {
                    "company_name": company_name,
                    "merchant_slug": extract_slug_from_url(full_url),
                    "merchant_url": full_url,
                    "bnpl_provider": BNPL_PROVIDER,
                    "country": COUNTRY,
                    "source_page": url,
                    "directory_section": section,
                    "availability_type": "Zip US store directory",
                    "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

        except Exception:
            continue

    print(f"Section {section}: extracted store rows = {len(rows)}")
    return rows


def scrape_zip_us_stores() -> pd.DataFrame:
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

        for section in DIRECTORY_SECTIONS:
            try:
                rows = scrape_one_directory_page(page, section)
                all_rows.extend(rows)

            except PlaywrightTimeoutError:
                print(f"Timeout on section {section}. Skipping.")
            except Exception as e:
                print(f"Error on section {section}: {e}")

        browser.close()

    seen = set()
    deduped_rows = []

    for row in all_rows:
        norm = normalize_name(row["company_name"])

        if not norm:
            continue

        key = (norm, row["merchant_url"].lower())
        if key in seen:
            continue

        seen.add(key)
        deduped_rows.append(row)

    df = pd.DataFrame(deduped_rows)

    if not df.empty:
        df = df.sort_values(["directory_section", "company_name"]).reset_index(drop=True)

    return df


def save_outputs(df: pd.DataFrame) -> None:
    if df.empty:
        print("No Zip stores extracted.")
        print(f"Check debug pages in: {DEBUG_DIR}")
        return

    df.to_excel(OUTPUT_EXCEL, index=False)
    df.to_csv(OUTPUT_CSV, index=False)

    print(f"\nSaved Excel file to: {OUTPUT_EXCEL}")
    print(f"Saved CSV file to: {OUTPUT_CSV}")
    print(f"Total unique Zip store rows extracted: {len(df)}")

    print("\nPreview:")
    print(df.head(40).to_string(index=False))


if __name__ == "__main__":
    zip_df = scrape_zip_us_stores()
    save_outputs(zip_df)