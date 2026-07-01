# BNPL Merchant Dataset

This repository builds a merchant-level dataset for studying Buy Now, Pay Later
(BNPL) adoption from the merchant side. The current GitHub checkpoint covers the
pipeline through merchant scraping, provider overlap construction,
industry/sub-industry classification, and public/private listing-status
classification.

## Research Question

**Does BNPL adoption improve merchant performance, and is the effect different
across merchant categories?**

The project is motivated by the fact that much of the existing BNPL literature
focuses on consumers. This dataset is designed to support merchant-side analysis,
including future difference-in-differences/event-study work once adoption dates
and company financial metrics are finalized.

## Working Hypotheses

### H1: Revenue Growth Hypothesis

BNPL adoption increases merchant revenue growth by improving checkout conversion
and expanding access to credit-constrained consumers.

### H2: Profitability Hypothesis

BNPL adoption improves merchant profitability only when incremental gross profit
from new sales exceeds BNPL-related fee and substitution costs.

### H3: Industry Heterogeneity Hypothesis

The BNPL effect is stronger in discretionary, high-ticket, high-margin, and
e-commerce-intensive industries.

### H4: Firm Size Hypothesis

BNPL effects are more detectable among small- and mid-cap focused retailers than
among large diversified merchants.

### H5: Dynamic Effect Hypothesis

BNPL adoption may produce a short-run sales boost, but long-run effects depend on
whether BNPL generates repeat customers rather than merely pulling future demand
forward.

## Current Progress

Completed in this checkpoint:

1. Scraped merchant lists from major BNPL provider directories.
2. Combined provider-specific lists into a long merchant table.
3. Built provider-overlap tables to identify merchants listed by multiple BNPL
   services.
4. Classified merchants into broad industries and sub-industries using
   web-based evidence.
5. Built a cleaned final overlap table with industry labels.
6. Classified merchants into public/private listing status.
7. Added a manual override layer for brands whose listed financial statements
   belong to a public parent company rather than the merchant brand itself.

This checkpoint now also includes a first adoption-date and DiD-analysis layer:
confirmed BNPL adoption evidence for the public-company sample, industry labels
for confirmed adopters, SEC Companyfacts quarterly/annual financial panels, and
first-pass H1/H2 outputs. See
`Analysis/BNPL_DiD_H1_H2/`.

## Data Sources

Merchant directories currently used:

- Affirm Apple Pay merchant list:
  <https://www.affirm.com/wallet/shopping/applepaymerchants>
- Klarna US store directory:
  <https://www.klarna.com/us/store/>
- Afterpay US store directory:
  <https://www.afterpay.com/en-us/stores>
- Clearpay UK retailer pages:
  <https://www.clearpay.co.uk/en-GB/stores/retailer-pages>
- Zip US store directory:
  <https://zip.co/us/store/directory>

Public listing classification uses public reference sources such as SEC company
tickers, Yahoo Finance search, and Wikidata. Automated matches are treated as
candidates and are corrected through the manual override layer when needed.

## Repository Structure

```text
.
├── scrape_*.py
├── combine_bnpl_merchant_lists.py
├── apply_manual_industry_labels.py
├── classify_merchants_from_web_only_calibrated_v7_parse_safe.py
├── build_final_bnpl_overlap_tables.py
├── classify_merchants_public_listing_status.py
├── Analysis/
│   └── BNPL_DiD_H1_H2/
│       ├── data/
│       ├── outputs/
│       ├── scripts/
│       └── README.md
├── Data_Clean/
│   └── final_outputs/
│       ├── bnpl_merchant_master_long_labeled.xlsx
│       ├── bnpl_merchant_master_web_only_classified_v7_parse_safe.xlsx
│       ├── bnpl_merchant_overlap_summary_final_v7.xlsx
│       ├── bnpl_merchant_public_listing_status.xlsx
│       └── public_listing_manual_overrides.csv
└── requirements.txt
```

Most raw/intermediate data, web caches, run logs, and large report archives are
ignored by Git. Only curated checkpoint outputs are tracked under
`Data_Clean/final_outputs/`.

## Main Scripts

### 1. Provider Scrapers

- `scrape_affirm_applepay_merchants.py`
- `scrape_klarna_us_stores.py`
- `scrape_Afterpay_us_stores.py`
- `scrape_clear_pay_uk_stores.py`
- `scrape_zip_us_stores.py`

These scripts collect merchant names and URLs from provider directories and save
provider-specific CSV/XLSX outputs locally.

### 2. Merchant Combination

```bash
python combine_bnpl_merchant_lists.py
```

Combines provider-level merchant files into a long table and overlap summaries.

### 3. Industry and Sub-Industry Classification

```bash
python classify_merchants_from_web_only_calibrated_v7_parse_safe.py
```

Classifies merchants using website and search evidence. The parse-safe version is
the recommended classifier because it falls back to search evidence when a
merchant website is unreachable or hard to parse.

### 4. Final Overlap Tables

```bash
python build_final_bnpl_overlap_tables.py
```

Refreshes the provider-overlap table using the final web-based industry labels.
The main output is:

```text
Data_Clean/bnpl_merchant_overlap_summary_final_v7.xlsx
```

### 5. Public Listing Status Classification

Recommended first-pass test:

```bash
python classify_merchants_public_listing_status.py --limit 100 --sleep 0.2
```

Full run:

```bash
python classify_merchants_public_listing_status.py --sleep 0.2
```

The script classifies each merchant as one of the following broad cases:

- directly public company;
- subsidiary/brand with public parent company;
- private/unlisted;
- ambiguous or needs manual review.

The output includes listing status, public parent name, ticker, exchange,
financial-data availability level, match confidence, source notes, and manual
review flags.

## Current Final Outputs

The curated checkpoint files currently tracked in Git are:

- `Data_Clean/final_outputs/bnpl_merchant_master_long_labeled.xlsx`
- `Data_Clean/final_outputs/bnpl_merchant_master_web_only_classified_v7_parse_safe.xlsx`
- `Data_Clean/final_outputs/bnpl_merchant_overlap_summary_final_v7.xlsx`
- `Data_Clean/final_outputs/bnpl_merchant_public_listing_status.xlsx`
- `Data_Clean/final_outputs/public_listing_manual_overrides.csv`

These files support the current research checkpoint through public-company
classification.

## Analysis Checkpoint: Adoption Dates and H1/H2 Panels

The first merchant-performance analysis checkpoint is tracked under:

```text
Analysis/BNPL_DiD_H1_H2/
```

It includes:

- confirmed public-company BNPL adoption-date evidence;
- a confirmed adoption sample merged with existing industry/sub-industry labels;
- SEC Companyfacts quarterly and annual financial panels;
- first-pass H1/H2 baseline regression results;
- first-pass H1/H2 baseline regression outputs.

Key current counts:

- 38 confirmed public-company adoption rows;
- 23 direct-public, merchant-level main-sample rows;
- 15 parent-level robustness rows;
- 1,317 SEC firm-quarter observations for 25 matched firms;
- 340 SEC firm-year observations for 25 matched firms.

The baseline H1/H2 outputs are pipeline validation and exploratory estimates,
not final causal DiD evidence. The next major research step is constructing
matched never-adopter public-company controls and then applying a modern
staggered-adoption DiD/event-study estimator.

## Setup

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Some scrapers use Playwright. If needed, install browser binaries with:

```bash
playwright install chromium
```

## Notes

Automated listing-status classification is noisy because many merchants are
brands, subsidiaries, or recently acquired companies. The manual override file is
therefore part of the research design, not just a convenience file. It records
cases where public financial data should be linked to a parent company rather
than to the merchant brand itself.
