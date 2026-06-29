# BNPL Merchant Dataset

This project collects publicly available merchant-level information related to Buy Now, Pay Later (BNPL) adoption.

The first scraper focuses on Affirm's Apple Pay merchant page and extracts merchant names and merchant links.

## Research Question and Hypotheses

This project studies BNPL adoption from a merchant-side perspective. While much of the existing BNPL literature focuses on consumer borrowing, liquidity constraints, and spending behavior, this project shifts attention to merchant outcomes.

The central research question is:

**Does BNPL adoption improve merchant performance, and is the effect different across merchant categories?**

### H1: Revenue Growth Hypothesis

BNPL adoption increases merchant revenue growth by improving checkout conversion and expanding access to credit-constrained consumers.

The intuition is that BNPL reduces the immediate payment burden faced by consumers at checkout. This may increase purchase completion rates, average order value, or total sales, especially for discretionary retail categories.

### H2: Profitability Hypothesis

BNPL adoption improves merchant profitability only when incremental gross profit from new sales exceeds BNPL-related fee and substitution costs.

This hypothesis recognizes that sales growth alone is not sufficient. BNPL providers often charge merchants higher fees than standard card payment rails. Therefore, BNPL may raise revenue but still reduce net profitability if the fee burden is too high or if existing customers simply switch from cheaper payment methods to BNPL.

### H3: Industry Heterogeneity Hypothesis

The BNPL effect is stronger in discretionary, high-ticket, high-margin, and e-commerce-intensive industries.

BNPL is expected to matter more for categories where consumers are more likely to delay or abandon purchases because of liquidity constraints. Examples include apparel and fashion, beauty and personal care, electronics, sports and outdoor goods, home and furniture, and other discretionary retail sectors.

### H4: Firm Size Hypothesis

BNPL effects are more detectable among small- and mid-cap focused retailers than among large diversified merchants.

For large diversified companies, BNPL may represent only a small part of total sales. For smaller or more focused merchants, BNPL adoption may have a more visible relationship with revenue growth, conversion, or customer acquisition.

### H5: Dynamic Effect Hypothesis

BNPL adoption produces a short-run sales boost, but long-run effects depend on whether BNPL generates repeat customers rather than merely pulling future demand forward.

A short-run increase in sales may reflect true demand expansion, improved conversion, or intertemporal substitution. The long-run effect depends on whether BNPL creates persistent customer acquisition and repeat purchases, or simply shifts future purchases into the present.

## Current Empirical Strategy

The first stage of this project constructs a public merchant-level BNPL dataset from major BNPL provider directories, including Affirm, Klarna, Afterpay, Clearpay, and Zip.

The dataset is used to document:

1. Which merchants are listed by each BNPL provider;
2. Which industries are most represented among BNPL-supported merchants;
3. Which merchants appear across multiple BNPL platforms;
4. Whether high-overlap merchants are concentrated in specific industries;
5. Which BNPL-covered merchants can be linked to public-company financial data.

The current dataset supports descriptive analysis and public-company candidate identification. Future work may add BNPL adoption timing and firm-level financial outcomes to study whether BNPL adoption is associated with changes in revenue growth, margins, or other merchant performance measures.

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