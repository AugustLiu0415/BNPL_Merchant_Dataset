# BNPL H1/H2 first-pass research design

Generated: 20260701_174503

## Sample construction

- Main sample: confirmed BNPL adoption rows with `listing_status == direct_public` and `financial_data_level == merchant_level`.
- Robustness sample: `subsidiary_of_public` / `parent_level` rows, because parent-level financials may dilute merchant-level BNPL effects.
- Current SEC panel only covers U.S. SEC registrants matched through the SEC ticker-CIK map. Non-U.S. firms need a separate provider or annual/interim reports.

## H1: revenue growth

Primary outcome:

```text
RevenueGrowth_it = firm FE + quarter FE + beta PostBNPL_it + error_it
```

Event-study version:

```text
RevenueGrowth_it = firm FE + quarter FE + sum_k beta_k EventTime_k + error_it
```

## H2: profitability

Outcomes:

- `gross_margin`
- `operating_margin`
- `net_margin`
- `roa`
- `operating_income_growth_yoy`

Interpretation caveat: SEC statements do not separately disclose BNPL fees for most merchants, so H2 is a reduced-form profitability effect, not a direct BNPL-fee test.

## DiD caveat

The current template is a baseline TWFE scaffold. Because BNPL adoption is staggered, final inference should add not-yet-treated / never-adopter controls and preferably use a modern DiD estimator such as Callaway-Sant'Anna, Sun-Abraham, did2s, or a stacked event study.
