#!/usr/bin/env python3
"""Run first-pass H1/H2 TWFE baseline regressions.

This runner is a pipeline check, not the final causal specification. It uses the
treated firm-quarter panel and firm/quarter fixed effects with firm-clustered
standard errors.
"""

from __future__ import annotations

import argparse
import datetime as dt
import math
from pathlib import Path

import numpy as np
import pandas as pd


OUTCOMES = [
    "revenue_growth_yoy",
    "gross_margin",
    "operating_margin",
    "net_margin",
    "roa",
    "operating_income_growth_yoy",
]


def latest_file(path: Path, pattern: str) -> Path:
    matches = sorted(path.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No files found under {path} matching {pattern}")
    return matches[-1]


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
    analysis_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--panel",
        default=str(latest_file(analysis_root / "data" / "panels", "bnpl_sec_financial_panel_quarterly_*.csv")),
    )
    parser.add_argument("--output", default="")
    args = parser.parse_args()

    panel_path = Path(args.panel)
    output_path = Path(args.output) if args.output else (
        analysis_root
        / "outputs"
        / f"bnpl_h1_h2_baseline_results_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(panel_path)
    df = df[df["sample_main_direct_public"].astype(str).str.lower().isin(["true", "1"])].copy()
    df = df[df["revenue"].notna()].copy()

    result = pd.DataFrame([fit_twfe(df, outcome) for outcome in OUTCOMES])
    result.to_csv(output_path, index=False)
    print(result.to_string(index=False))
    print(f"\nwrote {output_path}")


if __name__ == "__main__":
    main()
