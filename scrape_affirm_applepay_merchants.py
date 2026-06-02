import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


URL = "https://www.affirm.com/wallet/shopping/applepaymerchants"

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "Data_Raw"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_DIR / "affirm_applepay_merchants.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "affirm_applepay_merchants.csv"
DEBUG_HTML = OUTPUT_DIR / "affirm_page_debug.html"


def clean_text(text: str) -> str:
    """Clean extracted text."""
    if not text:
        return ""

    text = text.strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"^Continue to\s+", "", text, flags=re.IGNORECASE)
    return text.strip()


def is_likely_merchant(name: str, href: str) -> bool:
    """
    Filter out navigation/footer/legal/help links and keep likely merchant names.
    This is intentionally conservative.
    """
    if not name:
        return False

    lower_name = name.lower()

    blocked_names = {
        "affirm",
        "shop",
        "help",
        "help center",
        "log in",
        "sign in",
        "sign up",
        "get started",
        "for shoppers",
        "for businesses",
        "business",
        "developers",
        "careers",
        "investors",
        "press",
        "security",
        "privacy",
        "terms",
        "licenses",
        "accessibility",
        "your privacy choices",
        "download the app",
        "check my purchasing power",
        "learn more",
        "see all",
        "load more",
    }

    if lower_name in blocked_names:
        return False

    # Avoid very long text blocks accidentally captured from page sections.
    if len(name) > 80:
        return False

    # Avoid pure punctuation / numbers
    if not re.search(r"[A-Za-z]", name):
        return False

    blocked_url_parts = [
        "/help",
        "/business",
        "/careers",
        "/investors",
        "/press",
        "/licenses",
        "/terms",
        "/privacy",
        "support.apple.com",
        "fdic.gov",
    ]

    href_lower = (href or "").lower()
    if any(part in href_lower for part in blocked_url_parts):
        return False

    return True


def click_load_more_until_done(page, max_clicks: int = 100) -> int:
    """
    Click the 'Load more' button until it disappears or cannot be clicked.
    """
    clicks = 0

    for _ in range(max_clicks):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1000)

            button = page.get_by_text("Load more", exact=True)

            if button.count() == 0:
                print("No more 'Load more' button found.")
                break

            button.first.click(timeout=5000)
            clicks += 1
            print(f"Clicked Load more: {clicks}")

            page.wait_for_timeout(1500)

        except PlaywrightTimeoutError:
            print("Load more button exists but is not clickable anymore.")
            break
        except Exception as e:
            print(f"Stopped clicking Load more because: {e}")
            break

    return clicks


def scrape_affirm_merchants() -> pd.DataFrame:
    rows = []
    seen_names = set()

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

        print(f"Opening: {URL}")
        page.goto(URL, wait_until="domcontentloaded", timeout=120000)
        page.wait_for_timeout(5000)

        # Save debug HTML in case we need to inspect page structure later.
        DEBUG_HTML.write_text(page.content(), encoding="utf-8")

        click_load_more_until_done(page)

        # Extract all links after full page expansion.
        links = page.locator("a").all()
        print(f"Total links found on page: {len(links)}")

        for link in links:
            try:
                name = clean_text(link.inner_text(timeout=2000))
                href = link.get_attribute("href") or ""
                full_url = urljoin(URL, href)

                if not is_likely_merchant(name, full_url):
                    continue

                key = name.lower()
                if key in seen_names:
                    continue

                seen_names.add(key)

                rows.append(
                    {
                        "company_name": name,
                        "merchant_url": full_url,
                        "bnpl_provider": "Affirm",
                        "source_page": URL,
                        "availability_type": "Affirm through Apple Pay merchant list",
                        "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )

            except Exception:
                continue

        browser.close()

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.sort_values("company_name").reset_index(drop=True)

    return df


def save_outputs(df: pd.DataFrame) -> None:
    if df.empty:
        print("No merchants extracted. Check Data_Raw/affirm_page_debug.html.")
        return

    df.to_excel(OUTPUT_EXCEL, index=False)
    df.to_csv(OUTPUT_CSV, index=False)

    print(f"\nSaved Excel file to: {OUTPUT_EXCEL}")
    print(f"Saved CSV file to: {OUTPUT_CSV}")
    print(f"Total unique merchants extracted: {len(df)}")

    print("\nPreview:")
    print(df.head(20).to_string(index=False))


if __name__ == "__main__":
    merchants_df = scrape_affirm_merchants()
    save_outputs(merchants_df)