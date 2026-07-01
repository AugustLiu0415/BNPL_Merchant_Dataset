#!/usr/bin/env python3
"""Build BNPL adoption samples and first-pass SEC financial panels for H1/H2.

The panel intentionally uses structured SEC XBRL Companyfacts data for U.S.
registrants instead of scraping 10-K/10-Q PDFs. This gives auditable annual
and quarterly statement variables for the DiD/event-study workflow.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import re
import time
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import numpy as np
import pandas as pd


SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik10}.json"

UA = "BNPL merchant performance research augustliu0415@gmail.com"

US_EXCHANGES = {"NASDAQ", "NYSE", "NYSE ADR", "OTC Markets"}

QUARTER_FRAME_RE = re.compile(r"^CY(?P<year>\d{4})Q(?P<quarter>[1-4])I?$")
ANNUAL_FRAME_RE = re.compile(r"^CY(?P<year>\d{4})$")

METRIC_TAGS = {
    "revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "SalesRevenueNet",
        "Revenues",
        "SalesRevenueGoodsNet",
    ],
    "gross_profit": ["GrossProfit"],
    "operating_income": ["OperatingIncomeLoss"],
    "net_income": ["NetIncomeLoss", "ProfitLoss"],
    "total_assets": ["Assets"],
    "sga": [
        "SellingGeneralAndAdministrativeExpense",
        "SellingAndMarketingExpense",
    ],
    "capex": [
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "PaymentsToAcquireProductiveAssets",
    ],
    "cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    ],
}

INSTANT_METRICS = {"total_assets", "cash"}


def http_json(url: str, cache_path: Path | None = None, sleep: float = 0.11) -> dict:
    if cache_path and cache_path.exists():
        return json.loads(cache_path.read_text(encoding="utf-8"))
    req = Request(url, headers={"User-Agent": UA, "Accept-Encoding": "identity"})
    with urlopen(req, timeout=30) as response:
        data = json.loads(response.read().decode("utf-8", "replace"))
    if cache_path:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(json.dumps(data), encoding="utf-8")
    time.sleep(sleep)
    return data


def latest_file(output_dir: Path, pattern: str) -> Path:
    paths = sorted(output_dir.glob(pattern))
    if not paths:
        raise FileNotFoundError(f"No files found for pattern {pattern!r} in {output_dir}")
    return paths[-1]


def clean_key(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_adoption_date(value: object) -> pd.Timestamp:
    if pd.isna(value) or not str(value).strip():
        return pd.NaT
    return pd.to_datetime(str(value), errors="coerce")


def event_quarter(date: pd.Timestamp) -> str:
    if pd.isna(date):
        return ""
    q = int((date.month - 1) // 3 + 1)
    return f"{date.year}Q{q}"


def quarter_index(year: int, quarter: int) -> int:
    return int(year) * 4 + int(quarter)


def ticker_base(ticker: object) -> str:
    ticker = clean_key(ticker).upper()
    if not ticker:
        return ""
    # SEC ticker map uses share-class tickers with hyphen, e.g. BRK-B.
    return ticker.replace(".", "-")


def build_confirmed_sample(adoption_path: Path, public_workbook: Path) -> pd.DataFrame:
    adoption = pd.read_csv(adoption_path)
    public = pd.read_excel(public_workbook, sheet_name="Public_Analysis_Sample")

    keep_cols = [
        "clean_company_key",
        "verified_broad_industry",
        "verified_sub_industry",
        "classification_confidence",
        "verification_status",
        "country_market",
        "cik",
        "listing_match_confidence",
        "needs_listing_review",
        "public_company_name",
        "public_parent_name",
    ]
    merged = adoption.merge(
        public[keep_cols],
        left_on="merchant_key",
        right_on="clean_company_key",
        how="left",
    )
    merged["adoption_date"] = merged["confirmed_bnpl_adoption_date"].apply(parse_adoption_date)
    merged = merged[merged["adoption_date"].notna()].copy()
    merged["adoption_year"] = merged["adoption_date"].dt.year
    merged["adoption_quarter"] = merged["adoption_date"].apply(event_quarter)
    merged["adoption_quarter_index"] = merged["adoption_date"].apply(
        lambda x: quarter_index(x.year, int((x.month - 1) // 3 + 1))
    )
    merged["sample_main_direct_public"] = (
        (merged["listing_status"] == "direct_public")
        & (merged["financial_data_level"] == "merchant_level")
        & (merged["public_match_audit_status"] != "exclude_false_public_match")
    )
    merged["sample_parent_level_robustness"] = (
        (merged["listing_status"] == "subsidiary_of_public")
        & (merged["financial_data_level"] == "parent_level")
        & (merged["public_match_audit_status"] != "exclude_false_public_match")
    )
    merged["is_us_sec_candidate"] = merged["exchange"].isin(US_EXCHANGES)
    merged["ticker_for_sec_match"] = merged["ticker"].apply(ticker_base)
    merged["window_start_year"] = merged["adoption_year"] - 10
    merged["window_end_year"] = merged["adoption_year"] + 10
    return merged


def build_ticker_map(cache_dir: Path) -> pd.DataFrame:
    data = http_json(SEC_TICKERS_URL, cache_dir / "company_tickers.json")
    rows = []
    for item in data.values():
        rows.append(
            {
                "sec_cik": int(item["cik_str"]),
                "sec_ticker": str(item["ticker"]).upper(),
                "sec_title": item["title"],
            }
        )
    return pd.DataFrame(rows).drop_duplicates(subset=["sec_ticker"])


def fact_units(companyfacts: dict, tag: str) -> list[dict]:
    facts = companyfacts.get("facts", {}).get("us-gaap", {}).get(tag, {})
    units = facts.get("units", {})
    rows: list[dict] = []
    for unit_name in ("USD", "USD/shares", "shares"):
        rows.extend(units.get(unit_name, []))
    return rows


def usable_fact(row: dict) -> bool:
    return (
        row.get("val") is not None
        and row.get("frame")
        and row.get("form") in {"10-Q", "10-K", "10-Q/A", "10-K/A"}
    )


def choose_fact(rows: Iterable[dict]) -> dict | None:
    rows = [row for row in rows if usable_fact(row)]
    if not rows:
        return None
    # Prefer original filings, latest filed date, then highest accession number.
    rows = sorted(
        rows,
        key=lambda row: (
            0 if row.get("form", "").endswith("/A") else 1,
            str(row.get("filed", "")),
            str(row.get("accn", "")),
        ),
        reverse=True,
    )
    return rows[0]


def extract_metric_facts(companyfacts: dict, metric: str, tags: list[str]) -> tuple[list[dict], list[dict]]:
    quarterly: dict[tuple[int, int], list[dict]] = {}
    annual: dict[int, list[dict]] = {}
    for tag in tags:
        for row in fact_units(companyfacts, tag):
            frame = row.get("frame") or ""
            qmatch = QUARTER_FRAME_RE.match(frame)
            amatch = ANNUAL_FRAME_RE.match(frame)
            enriched = dict(row)
            enriched["metric"] = metric
            enriched["xbrl_tag"] = tag
            if qmatch:
                key = (int(qmatch.group("year")), int(qmatch.group("quarter")))
                quarterly.setdefault(key, []).append(enriched)
                if metric in INSTANT_METRICS and int(qmatch.group("quarter")) == 4:
                    annual.setdefault(int(qmatch.group("year")), []).append(enriched)
            elif amatch:
                key = int(amatch.group("year"))
                annual.setdefault(key, []).append(enriched)

    q_rows = []
    for (year, quarter), candidates in quarterly.items():
        chosen = choose_fact(candidates)
        if chosen:
            q_rows.append(
                {
                    "calendar_year": year,
                    "calendar_quarter": quarter,
                    "period": f"{year}Q{quarter}",
                    "period_index": quarter_index(year, quarter),
                    "metric": metric,
                    "value": chosen["val"],
                    "xbrl_tag": chosen["xbrl_tag"],
                    "form": chosen.get("form", ""),
                    "filed": chosen.get("filed", ""),
                    "accn": chosen.get("accn", ""),
                    "fp": chosen.get("fp", ""),
                    "fy": chosen.get("fy", ""),
                    "frame": chosen.get("frame", ""),
                    "end": chosen.get("end", ""),
                }
            )

    a_rows = []
    for year, candidates in annual.items():
        chosen = choose_fact(candidates)
        if chosen:
            a_rows.append(
                {
                    "calendar_year": year,
                    "period": str(year),
                    "metric": metric,
                    "value": chosen["val"],
                    "xbrl_tag": chosen["xbrl_tag"],
                    "form": chosen.get("form", ""),
                    "filed": chosen.get("filed", ""),
                    "accn": chosen.get("accn", ""),
                    "fp": chosen.get("fp", ""),
                    "fy": chosen.get("fy", ""),
                    "frame": chosen.get("frame", ""),
                    "end": chosen.get("end", ""),
                }
            )
    return q_rows, a_rows


def pivot_panel(rows: list[dict], index_cols: list[str]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    long = pd.DataFrame(rows)
    values = long.pivot_table(
        index=index_cols,
        columns="metric",
        values="value",
        aggfunc="first",
    ).reset_index()
    tag_sources = long.pivot_table(
        index=index_cols,
        columns="metric",
        values="xbrl_tag",
        aggfunc="first",
    ).reset_index()
    tag_sources = tag_sources.rename(
        columns={col: f"{col}_xbrl_tag" for col in tag_sources.columns if col not in index_cols}
    )
    return values.merge(tag_sources, on=index_cols, how="left")


def add_outcomes(panel: pd.DataFrame, period_col: str) -> pd.DataFrame:
    if panel.empty:
        return panel
    panel = panel.sort_values(["merchant_key", period_col]).copy()
    for col in [
        "revenue",
        "gross_profit",
        "operating_income",
        "net_income",
        "total_assets",
        "sga",
        "capex",
        "cash",
    ]:
        if col not in panel.columns:
            panel[col] = np.nan

    panel["gross_margin"] = panel["gross_profit"] / panel["revenue"]
    panel["operating_margin"] = panel["operating_income"] / panel["revenue"]
    panel["net_margin"] = panel["net_income"] / panel["revenue"]
    panel["sga_to_revenue"] = panel["sga"] / panel["revenue"]
    panel["capex_to_revenue"] = panel["capex"] / panel["revenue"]
    panel["roa"] = panel["net_income"] / panel["total_assets"]
    panel["log_revenue"] = np.log(panel["revenue"].where(panel["revenue"] > 0))

    lag_n = 4 if period_col == "period_index" else 1
    grouped = panel.groupby("merchant_key", group_keys=False)
    panel["revenue_lag"] = grouped["revenue"].shift(lag_n)
    panel["revenue_growth_yoy"] = panel["revenue"] / panel["revenue_lag"] - 1
    panel["operating_income_lag"] = grouped["operating_income"].shift(lag_n)
    panel["operating_income_growth_yoy"] = panel["operating_income"] / panel["operating_income_lag"].abs() - 1
    panel.loc[panel["operating_income_lag"].isna() | (panel["operating_income_lag"] == 0), "operating_income_growth_yoy"] = np.nan
    return panel


def merge_event_fields(panel: pd.DataFrame, sample: pd.DataFrame, frequency: str) -> pd.DataFrame:
    if panel.empty:
        return panel
    cols = [
        "merchant_key",
        "merchant_name",
        "public_entity",
        "ticker",
        "exchange",
        "listing_status",
        "financial_data_level",
        "public_match_audit_status",
        "sample_main_direct_public",
        "sample_parent_level_robustness",
        "verified_broad_industry",
        "verified_sub_industry",
        "adoption_date",
        "adoption_year",
        "adoption_quarter",
        "adoption_quarter_index",
        "confirmed_date_type",
        "confirmed_evidence_strength",
        "confirmed_evidence_url",
    ]
    out = panel.merge(sample[cols], on="merchant_key", how="left")
    if frequency == "quarterly":
        out["event_time"] = out["period_index"] - out["adoption_quarter_index"]
        out["post_bnpl"] = (out["event_time"] >= 0).astype(int)
        out = out[(out["event_time"] >= -40) & (out["event_time"] <= 40)].copy()
    else:
        out["event_time"] = out["calendar_year"] - out["adoption_year"]
        out["post_bnpl"] = (out["event_time"] >= 0).astype(int)
        out = out[(out["event_time"] >= -10) & (out["event_time"] <= 10)].copy()
    return out


def fetch_sec_panels(sample: pd.DataFrame, cache_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    ticker_map = build_ticker_map(cache_dir)
    sec_sample = sample[sample["is_us_sec_candidate"]].copy()
    sec_sample = sec_sample.merge(
        ticker_map,
        left_on="ticker_for_sec_match",
        right_on="sec_ticker",
        how="left",
    )
    q_fact_rows: list[dict] = []
    a_fact_rows: list[dict] = []
    audit_rows = []

    for _, row in sec_sample.iterrows():
        merchant_key = row["merchant_key"]
        ticker = row["ticker_for_sec_match"]
        cik = row.get("sec_cik")
        if pd.isna(cik):
            audit_rows.append(
                {
                    "merchant_key": merchant_key,
                    "ticker": row["ticker"],
                    "sec_cik": "",
                    "sec_title": "",
                    "sec_companyfacts_url": "",
                    "fetch_status": "no_sec_ticker_match",
                    "error": "",
                }
            )
            continue

        cik_int = int(cik)
        cik10 = f"{cik_int:010d}"
        try:
            data = http_json(
                SEC_COMPANYFACTS_URL.format(cik10=cik10),
                cache_dir / "companyfacts" / f"CIK{cik10}.json",
            )
            company_name = data.get("entityName", "")
            found_metrics = []
            for metric, tags in METRIC_TAGS.items():
                q_rows, a_rows = extract_metric_facts(data, metric, tags)
                for item in q_rows:
                    item.update({"merchant_key": merchant_key, "sec_cik": cik_int, "sec_entity_name": company_name})
                for item in a_rows:
                    item.update({"merchant_key": merchant_key, "sec_cik": cik_int, "sec_entity_name": company_name})
                q_fact_rows.extend(q_rows)
                a_fact_rows.extend(a_rows)
                if q_rows or a_rows:
                    found_metrics.append(metric)
            audit_rows.append(
                {
                    "merchant_key": merchant_key,
                    "ticker": row["ticker"],
                    "sec_cik": cik_int,
                    "sec_title": row.get("sec_title", ""),
                    "sec_companyfacts_url": SEC_COMPANYFACTS_URL.format(cik10=cik10),
                    "sec_entity_name": company_name,
                    "fetch_status": "ok",
                    "metrics_found": "; ".join(found_metrics),
                    "error": "",
                }
            )
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
            audit_rows.append(
                {
                    "merchant_key": merchant_key,
                    "ticker": row["ticker"],
                    "sec_cik": cik_int,
                    "sec_title": row.get("sec_title", ""),
                    "sec_companyfacts_url": SEC_COMPANYFACTS_URL.format(cik10=cik10),
                    "fetch_status": "fetch_error",
                    "metrics_found": "",
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )

    q_panel = pivot_panel(
        q_fact_rows,
        ["merchant_key", "sec_cik", "sec_entity_name", "calendar_year", "calendar_quarter", "period", "period_index"],
    )
    a_panel = pivot_panel(
        a_fact_rows,
        ["merchant_key", "sec_cik", "sec_entity_name", "calendar_year", "period"],
    )
    if not q_panel.empty:
        q_panel["sec_companyfacts_url"] = q_panel["sec_cik"].apply(
            lambda cik: SEC_COMPANYFACTS_URL.format(cik10=f"{int(cik):010d}")
        )
    if not a_panel.empty:
        a_panel["sec_companyfacts_url"] = a_panel["sec_cik"].apply(
            lambda cik: SEC_COMPANYFACTS_URL.format(cik10=f"{int(cik):010d}")
        )
    return q_panel, a_panel, pd.DataFrame(audit_rows)


def write_regression_templates(output_dir: Path, stamp: str) -> tuple[Path, Path]:
    py_path = output_dir / f"bnpl_h1_h2_regression_template_{stamp}.py"
    md_path = output_dir / f"bnpl_h1_h2_research_design_{stamp}.md"
    py_path.write_text(
        '''#!/usr/bin/env python3
"""Baseline H1/H2 regressions for the BNPL merchant panel.

Run from the project root after building the panel:
python outputs/bnpl_h1_h2_regression_template_<stamp>.py
"""

import math
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PANEL = ROOT / "outputs" / "bnpl_sec_financial_panel_quarterly_<STAMP>.csv"
OUT = ROOT / "outputs" / "bnpl_h1_h2_baseline_results_<STAMP>.csv"


def normal_pvalue(t_stat: float) -> float:
    if not np.isfinite(t_stat):
        return np.nan
    return math.erfc(abs(float(t_stat)) / math.sqrt(2.0))


def fit_twfe(data: pd.DataFrame, outcome: str) -> dict:
    sub = data[[outcome, "post_bnpl", "merchant_key", "period"]].dropna().copy()
    if sub["post_bnpl"].nunique() < 2 or sub["merchant_key"].nunique() < 3:
        return {
            "outcome": outcome,
            "nobs": len(sub),
            "n_firms": sub["merchant_key"].nunique(),
            "coef_post_bnpl": np.nan,
            "se_cluster_firm": np.nan,
            "t_stat": np.nan,
            "p_value_normal": np.nan,
            "status": "not_enough_variation",
        }

    firm_dummies = pd.get_dummies(sub["merchant_key"], prefix="firm", drop_first=True, dtype=float)
    time_dummies = pd.get_dummies(sub["period"], prefix="time", drop_first=True, dtype=float)
    x_df = pd.concat(
        [
            pd.Series(1.0, index=sub.index, name="const"),
            sub["post_bnpl"].astype(float),
            firm_dummies,
            time_dummies,
        ],
        axis=1,
    )
    y = sub[outcome].astype(float).to_numpy()
    x = x_df.to_numpy(dtype=float)
    colnames = list(x_df.columns)
    post_idx = colnames.index("post_bnpl")

    beta = np.linalg.pinv(x.T @ x) @ x.T @ y
    resid = y - x @ beta
    xtx_inv = np.linalg.pinv(x.T @ x)

    meat = np.zeros((x.shape[1], x.shape[1]))
    for _, idx in sub.groupby("merchant_key").groups.items():
        loc = sub.index.get_indexer(idx)
        xg = x[loc, :]
        ug = resid[loc]
        score = xg.T @ ug
        meat += np.outer(score, score)

    nobs, k = x.shape
    clusters = sub["merchant_key"].nunique()
    cov = xtx_inv @ meat @ xtx_inv
    if clusters > 1 and nobs > k:
        cov *= (clusters / (clusters - 1)) * ((nobs - 1) / (nobs - k))
    se = float(np.sqrt(max(cov[post_idx, post_idx], 0.0)))
    coef = float(beta[post_idx])
    t_stat = coef / se if se else np.nan
    return {
        "outcome": outcome,
        "nobs": nobs,
        "n_firms": clusters,
        "coef_post_bnpl": coef,
        "se_cluster_firm": se,
        "t_stat": t_stat,
        "p_value_normal": normal_pvalue(t_stat),
        "status": "ok",
    }


def main() -> None:
    df = pd.read_csv(PANEL)
    df = df[df["sample_main_direct_public"].astype(str).str.lower().isin(["true", "1"])].copy()
    df = df[df["revenue"].notna()].copy()
    df["firm_fe"] = df["merchant_key"].astype(str)
    df["time_fe"] = df["period"].astype(str)

    outcomes = [
        "revenue_growth_yoy",
        "gross_margin",
        "operating_margin",
        "net_margin",
        "roa",
        "operating_income_growth_yoy",
    ]
    rows = []
    for outcome in outcomes:
        rows.append(fit_twfe(df, outcome))
    result = pd.DataFrame(rows)
    result.to_csv(OUT, index=False)
    print(result.to_string(index=False))
    print(f"\\nwrote {OUT}")


if __name__ == "__main__":
    main()
'''.replace("<STAMP>", stamp),
        encoding="utf-8",
    )
    md_path.write_text(
        f"""# BNPL H1/H2 first-pass research design

Generated: {stamp}

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
""",
        encoding="utf-8",
    )
    return py_path, md_path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default="outputs")
    parser.add_argument("--adoption-csv", default="")
    parser.add_argument(
        "--public-workbook",
        default="work/BNPL_Merchant_Dataset/Data_Clean/final_outputs/bnpl_merchant_public_listing_status.xlsx",
    )
    parser.add_argument("--cache-dir", default="work/sec_cache")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    adoption_path = Path(args.adoption_csv) if args.adoption_csv else latest_file(
        output_dir,
        "bnpl_public_company_adoption_evidence_core_enriched_*.csv",
    )
    public_workbook = Path(args.public_workbook)
    cache_dir = Path(args.cache_dir)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    sample = build_confirmed_sample(adoption_path, public_workbook)
    sample_path = output_dir / f"bnpl_confirmed_adoption_industry_sample_{stamp}.csv"
    sample.to_csv(sample_path, index=False)

    q_panel, a_panel, audit = fetch_sec_panels(sample, cache_dir)
    q_panel = add_outcomes(merge_event_fields(q_panel, sample, "quarterly"), "period_index")
    a_panel = add_outcomes(merge_event_fields(a_panel, sample, "annual"), "calendar_year")

    q_path = output_dir / f"bnpl_sec_financial_panel_quarterly_{stamp}.csv"
    a_path = output_dir / f"bnpl_sec_financial_panel_annual_{stamp}.csv"
    audit_path = output_dir / f"bnpl_sec_companyfacts_fetch_audit_{stamp}.csv"
    q_panel.to_csv(q_path, index=False)
    a_panel.to_csv(a_path, index=False)
    audit.to_csv(audit_path, index=False)
    reg_path, design_path = write_regression_templates(output_dir, stamp)

    print(f"adoption source: {adoption_path}")
    print(f"wrote {sample_path} rows={len(sample)}")
    print(f"wrote {q_path} rows={len(q_panel)} firms={q_panel['merchant_key'].nunique() if not q_panel.empty else 0}")
    print(f"wrote {a_path} rows={len(a_panel)} firms={a_panel['merchant_key'].nunique() if not a_panel.empty else 0}")
    print(f"wrote {audit_path} rows={len(audit)}")
    print(f"wrote {reg_path}")
    print(f"wrote {design_path}")


if __name__ == "__main__":
    main()
