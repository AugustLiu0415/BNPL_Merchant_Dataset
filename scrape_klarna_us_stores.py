import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


BASE_URL = "https://www.klarna.com/us/store/"
BNPL_PROVIDER = "Klarna"

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "Data_Raw"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_DIR / "klarna_us_stores.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "klarna_us_stores.csv"

# First test only 3 pages. If it works, we can scrape more.
TEST_MAX_PAGES = 3


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    return text


def parse_card_text(raw_text: str) -> dict:
    """
    Klarna store cards usually look like:
    Nike
    Sports
    Klarna at checkout - Google Pay
    """
    lines = [clean_text(x) for x in raw_text.split("\n") if clean_text(x)]

    company_name = lines[0] if len(lines) >= 1 else ""
    category = lines[1] if len(lines) >= 2 else ""
    availability_type = " - ".join(lines[2:]) if len(lines) >= 3 else ""

    return {
        "company_name": company_name,
        "category": category,
        "availability_type": availability_type,
    }


def is_likely_store(company_name: str, raw_text: str, href: str) -> bool:
    if not company_name:
        return False

    name_lower = company_name.lower().strip()
    text_lower = raw_text.lower()
    href_lower = href.lower()

    blocked_names = {
        "klarna",
        "help",
        "about us",
        "careers",
        "privacy",
        "terms",
        "legal",
        "press",
        "security",
        "sign in",
        "for business",
        "for shoppers",
        "customer service",
        "business login",
        "developer portal",
        "all categories",
        "united states",
        "show more",
    }

    if name_lower in blocked_names:
        return False

    if re.fullmatch(r"\d+", company_name):
        return False

    # These words usually appear inside real Klarna store cards.
    store_signals = [
        "klarna at checkout",
        "shop in the klarna app",
        "google pay",
        "apple pay",
        "onepay later",
    ]

    if any(signal in text_lower for signal in store_signals):
        return True

    # Store links often contain /us/store/
    if "/us/store/" in href_lower and href_lower.rstrip("/") != BASE_URL.rstrip("/"):
        return True

    return False


def scrape_one_page(page, page_number: int) -> list[dict]:
    url = f"{BASE_URL}?page={page_number}"
    print(f"Opening Klarna page {page_number}: {url}")

    page.goto(url, wait_until="domcontentloaded", timeout=120000)
    page.wait_for_timeout(5000)

    rows = []

    # Klarna store cards are usually clickable link blocks.
    links = page.locator("a").all()
    print(f"Page {page_number}: total links found = {len(links)}")

    for link in links:
        try:
            raw_text = link.inner_text(timeout=2000)
            raw_text = raw_text.strip()
            href = link.get_attribute("href") or ""
            full_url = urljoin(BASE_URL, href)

            parsed = parse_card_text(raw_text)
            company_name = parsed["company_name"]

            if not is_likely_store(company_name, raw_text, full_url):
                continue

            rows.append(
                {
                    "company_name": company_name,
                    "category": parsed["category"],
                    "availability_type": parsed["availability_type"],
                    "merchant_url": full_url,
                    "bnpl_provider": BNPL_PROVIDER,
                    "source_page": url,
                    "page_number": page_number,
                    "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

        except Exception:
            continue

    print(f"Page {page_number}: extracted stores = {len(rows)}")
    return rows


def scrape_klarna_stores() -> pd.DataFrame:
    all_rows = []
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

        for page_number in range(1, TEST_MAX_PAGES + 1):
            try:
                rows = scrape_one_page(page, page_number)

                for row in rows:
                    key = (row["company_name"].lower(), row["merchant_url"].lower())
                    if key not in seen:
                        seen.add(key)
                        all_rows.append(row)

            except PlaywrightTimeoutError:
                print(f"Timeout on page {page_number}. Skipping.")
            except Exception as e:
                print(f"Error on page {page_number}: {e}")

        browser.close()

    df = pd.DataFrame(all_rows)

    if not df.empty:
        df = df.drop_duplicates(subset=["company_name", "merchant_url"])
        df = df.sort_values(["page_number", "company_name"]).reset_index(drop=True)

    return df


def save_outputs(df: pd.DataFrame) -> None:
    if df.empty:
        print("No Klarna stores extracted.")
        return

    df.to_excel(OUTPUT_EXCEL, index=False)
    df.to_csv(OUTPUT_CSV, index=False)

    print(f"\nSaved Excel file to: {OUTPUT_EXCEL}")
    print(f"Saved CSV file to: {OUTPUT_CSV}")
    print(f"Total Klarna store rows extracted: {len(df)}")

    print("\nPreview:")
    print(df.head(30).to_string(index=False))


if __name__ == "__main__":
    stores_df = scrape_klarna_stores()
    save_outputs(stores_df)