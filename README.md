# BNPL Merchant Dataset

This project collects publicly available merchant-level information related to Buy Now, Pay Later (BNPL) adoption.

The first scraper focuses on Affirm's Apple Pay merchant page and extracts merchant names and merchant links.

## Research Motivation

This project supports a merchant-level analysis of whether BNPL adoption is associated with changes in sales, revenue, and financial performance.

The broader research question is whether BNPL adoption improves merchant outcomes through conversion lift, price discrimination, demand expansion, or other channels.

## Current Data Source

- Affirm Apple Pay merchant list:  
  https://www.affirm.com/wallet/shopping/applepaymerchants
- Klarna US store directory:
  https://www.klarna.com/us/store/
- Afterpay US store directory: 
  https://www.afterpay.com/en-us/stores
- Clearpay UK retailer pages:
  https://www.clearpay.co.uk/en-GB/stores/retailer-pages
- Zip US retailer pages: 
  https://zip.co/us/store/directory



## Main Script

- `scrape_affirm_applepay_merchants.py`

This script:

1. Opens the Affirm Apple Pay merchant page.
2. Clicks "Load more" until no more merchants are available.
3. Extracts merchant names and URLs.
4. Saves the output as Excel and CSV files.

- `scrape_klarna_us_stores.py`

This script: 

1. Opens the Klarna US store list page.
2. Clicks "Load more" until no more merchants are available.
3. Extracts merchant names and URLs.
4. Saves the output as Excel and CSV files.

- `scrape_Afterpay_us_stores.py`

This script: 

1. Opens the Afterpay US store list page.
2. clicks "Load more" until no more merchants are available.
3. Extracts merchant names and URLs. 
4. Saves the output as Excel and CSV files

-  `scrape_clearpay_uk_stores.py` 

This script: 

1. Opens the Clearpay UK store list page. 
2. Extracts merchant names and URLs.
3. Saves the output as Excel and CSV files.

- `scrape_zip_us_stores.py` 

This script: 

1. Opens the Zip US store list page. 
2. Clicks "Load more" until no more merchantsd are available.
3. Extracts merchant names and URLs.
4. Saves the output as Excel and CSV files. 

## Output

The script generates output files in:

- `Data_Raw/affirm_applepay_merchants.xlsx`
- `Data_Raw/affirm_applepay_merchants.csv`
- `Data_Raw/affirm_page_debug.html`
- `Data_Raw/klarna_us_stores.xlsx`
- `Data_Raw/klarna_us_stores.csv`
- `Data_Raw/afterpay_us_stores.xlsx`
- `Data_Raw/afterpay_us_stores.csv`
- `Data_Raw/afterpay_page_debug.html`
- `Data_Raw/clearpay_uk_retailer_pages.xlsx`
- `Data_Raw/clearpay_uk_retailer_pages.csv`
- `Data_Raw/clearpay_retailer_pages_debug.html`
- `Data_Raw/zip_page_debug.html`
- `Data_Raw/zip_us_stores.csv`
- `Data_Raw/zip_us_stores.xlsx`


These output files are not uploaded to GitHub by default.

## Setup

Create and activate a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate