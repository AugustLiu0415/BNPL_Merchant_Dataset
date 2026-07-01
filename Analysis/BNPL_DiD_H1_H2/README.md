# BNPL DiD Analysis Checkpoint: Adoption Evidence and H1/H2 Setup

This folder contains the first analysis checkpoint for studying whether BNPL
adoption affects public merchant performance.

## What is included

- Confirmed BNPL adoption-date evidence for the public-company sample.
- Industry and sub-industry labels merged from the existing merchant
  classification workbook.
- SEC Companyfacts annual and quarterly financial panels for U.S. SEC-matched
  public companies.
- First-pass H1/H2 baseline regressions.

## Key sample counts

- Confirmed public-company adoption rows: 38.
- Main direct-public, merchant-level rows: 23.
- Parent-level robustness rows: 15.
- SEC quarterly panel: 1,317 firm-quarter rows for 25 matched firms.
- SEC annual panel: 340 firm-year rows for 25 matched firms.

## H1/H2 interpretation

The baseline H1/H2 output is a first-pass treated-firm panel scaffold. It uses:

```text
Outcome_it = firm FE + quarter FE + beta PostBNPL_it + error_it
```

Current outcomes include revenue growth, gross margin, operating margin, net
margin, ROA, and operating-income growth.

Important caveat: this is not the final causal DiD design because the current
checkpoint does not yet include a matched never-adopter control group. It is
best interpreted as a pipeline validation and exploratory baseline.

## Recommended next step

Construct matched public-company controls by industry. For each treated firm or
treated industry, candidate controls should be public retailers that have no
confirmed Affirm, Klarna, Afterpay/Clearpay, or Zip adoption evidence and are
matched on:

- broad industry and sub-industry;
- pre-adoption revenue scale;
- pre-adoption revenue growth trend;
- margin profile;
- total assets or market-cap size;
- SEC filing availability and exchange/listing type.

After controls are added, prefer a modern staggered-adoption DiD estimator such
as Callaway-Sant'Anna, Sun-Abraham, did2s, or a stacked event-study design.

## Reproduction commands

From the repository root, rebuild the H1/H2 panel with explicit paths:

```bash
python Analysis/BNPL_DiD_H1_H2/scripts/build_bnpl_did_panel.py \
  --adoption-csv Analysis/BNPL_DiD_H1_H2/data/adoption_evidence/bnpl_public_company_adoption_evidence_core_enriched_20260701_165450.csv \
  --public-workbook Data_Clean/final_outputs/bnpl_merchant_public_listing_status.xlsx \
  --output-dir Analysis/BNPL_DiD_H1_H2/data/panels \
  --cache-dir Analysis/BNPL_DiD_H1_H2/sec_cache
```

## File map

```text
bnpl_h1_h2_research_design_20260701_174503.md

data/adoption_evidence/
  bnpl_public_company_adoption_evidence_core_enriched_20260701_165450.csv
  bnpl_public_company_adoption_evidence_enriched_20260701_165450.csv
  bnpl_adoption_date_methodology_notes_20260701_165450.md

data/panels/
  bnpl_confirmed_adoption_industry_sample_20260701_174503.csv
  bnpl_sec_financial_panel_quarterly_20260701_174503.csv
  bnpl_sec_financial_panel_annual_20260701_174503.csv
  bnpl_sec_companyfacts_fetch_audit_20260701_174503.csv

outputs/
  bnpl_h1_h2_baseline_results_20260701_174503.csv

scripts/
  apply_manual_bnpl_evidence.py
  build_bnpl_did_panel.py
  run_h1_h2_baseline_regressions.py
```
