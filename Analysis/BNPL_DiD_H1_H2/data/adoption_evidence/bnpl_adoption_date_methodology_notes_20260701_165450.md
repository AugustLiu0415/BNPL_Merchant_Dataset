# BNPL adoption-date evidence notes

Generated from the prior enriched public-company evidence table plus `work/manual_bnpl_press_evidence.csv`.

## Output files

- `bnpl_public_company_adoption_evidence_enriched_20260701_165450.csv`
  - 191 rows, matching the workbook's public analysis sample.
  - Includes public-match audit status, confirmed/first-observed adoption evidence, weak provider-directory evidence, source URLs, and notes.
- `bnpl_public_company_adoption_evidence_core_enriched_20260701_165450.csv`
  - 75 core rows after keeping manually mapped/confirmed public rows and multi-provider automatic matches.
  - Excludes obvious false public-company name collisions and single-provider candidates needing public-match review.

## Evidence columns

- `confirmed_bnpl_adoption_date`
  - Filled only when a usable official announcement, credible report, or Wayback official merchant/help/payment page was found.
  - Check `confirmed_date_type` before using it as an event date.
- `confirmed_date_type`
  - `official_*`: stronger adoption/availability announcement.
  - `first_observed_official_page`: Wayback first-observed date on an official merchant/help/payment page; this is not necessarily the true launch date.
  - `reported_*`: credible news report; use as medium-strength evidence.
- `weak_provider_directory_first_observed_date`
  - Provider directory first-observed date from Wayback/Availability API.
  - Do not use this alone as formal adoption date.

## Current coverage

- 191 total public-analysis-sample rows.
- 75 core rows.
- 38 full-sample rows have confirmed or official-page first-observed BNPL dates.
- 38 core rows have confirmed or official-page first-observed BNPL dates.
- 92 rows have weak provider-directory first-observed dates.
- 34 rows are marked obvious false public-company matches.
