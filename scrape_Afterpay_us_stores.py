import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


URL = "https://www.afterpay.com/en-us/stores"
BNPL_PROVIDER = "Afterpay"

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "Data_Raw"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_DIR / "afterpay_us_stores.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "afterpay_us_stores.csv"
DEBUG_HTML = OUTPUT_DIR / "afterpay_page_debug.html"


def clean_text(text: str) -> str:
    """Clean extracted text."""
    if not text:
        return ""

    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_slug_from_url(url: str) -> str:
    """Extract merchant slug from an Afterpay store URL."""
    path = urlparse(url).path.rstrip("/")
    parts = path.split("/")
    if len(parts) >= 4 and parts[-2] == "stores":
        return parts[-1]
    return ""


def is_likely_afterpay_store(company_name: str, merchant_url: str) -> bool:
    """
    Keep only likely Afterpay merchant links.
    We want URLs like:
    https://www.afterpay.com/en-us/stores/nike
    """

    if not company_name:
        return False

    name_lower = company_name.lower().strip()
    parsed = urlparse(merchant_url)
    path = parsed.path.rstrip("/")

    # Keep only Afterpay US store detail pages
    if parsed.netloc not in {"www.afterpay.com", "afterpay.com"}:
        return False

    if not path.startswith("/en-us/stores/"):
        return False

    if path == "/en-us/stores":
        return False

    slug = extract_slug_from_url(merchant_url)
    if not slug:
        return False

    blocked_names = {
        "stores",
        "retailer pages",
        "help",
        "guide",
        "afterpay access",
        "retailer resources",
        "how afterpay works",
        "all categories",
        "for retailers",
        "get the app",
        "login",
        "sign up",
        "account login",
        "privacy",
        "terms",
        "contact",
        "careers",
        "security",
        "licenses",
        "mobile app",
        "investors",
        "media",
        "about us",
        "partner program",
        "responsible spending",
        "installment agreement",
    }

    if name_lower in blocked_names:
        return False

    # Remove region/location links
    blocked_region_names = {
        "united states",
        "australia",
        "new zealand",
        "united kingdom",
        "canada - english",
        "canada - français",
    }

    if name_lower in blocked_region_names:
        return False

    # Remove pure numbers or strange empty labels
    if re.fullmatch(r"\d+", company_name):
        return False

    return True


def scrape_afterpay_stores() -> pd.DataFrame:
    rows = []
    seen = set()

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

            links = page.locator("a").all()
            print(f"Total links found on page: {len(links)}")

            for link in links:
                try:
                    raw_text = link.inner_text(timeout=2000)
                    company_name = clean_text(raw_text)

                    href = link.get_attribute("href") or ""
                    merchant_url = urljoin(URL, href)

                    if not is_likely_afterpay_store(company_name, merchant_url):
                        continue

                    merchant_slug = extract_slug_from_url(merchant_url)

                    key = (company_name.lower(), merchant_url.lower())
                    if key in seen:
                        continue

                    seen.add(key)

                    rows.append(
                        {
                            "company_name": company_name,
                            "merchant_slug": merchant_slug,
                            "merchant_url": merchant_url,
                            "bnpl_provider": BNPL_PROVIDER,
                            "source_page": URL,
                            "availability_type": "Afterpay US store directory",
                            "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                    )

                except Exception:
                    continue

        except PlaywrightTimeoutError:
            print("Timeout while opening Afterpay page.")
        except Exception as e:
            print(f"Error while scraping Afterpay page: {e}")

        browser.close()

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.drop_duplicates(subset=["company_name", "merchant_url"])
        df = df.sort_values("company_name").reset_index(drop=True)

    return df


def save_outputs(df: pd.DataFrame) -> None:
    if df.empty:
        print("No Afterpay stores extracted. Check Data_Raw/afterpay_page_debug.html.")
        return

    df.to_excel(OUTPUT_EXCEL, index=False)
    df.to_csv(OUTPUT_CSV, index=False)

    print(f"\nSaved Excel file to: {OUTPUT_EXCEL}")
    print(f"Saved CSV file to: {OUTPUT_CSV}")
    print(f"Total unique Afterpay store rows extracted: {len(df)}")

    print("\nPreview:")
    print(df.head(30).to_string(index=False))


if __name__ == "__main__":
    stores_df = scrape_afterpay_stores()
    save_outputs(stores_df)