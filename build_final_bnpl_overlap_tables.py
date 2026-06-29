#!/usr/bin/env python3
"""
build_final_bnpl_overlap_tables.py

Purpose
-------
Refresh the BNPL merchant overlap table with the final web-based industry
classification output, then generate clean research-ready overlap and audit sheets.

Inputs expected in Data_Clean/
-----------------------------
1. bnpl_merchant_overlap_summary_labeled.xlsx
   - Existing overlap table with provider_set/provider_count/has_* columns.
   - Current industry labels in this file may be outdated.

2. One final classified merchant table, preferably:
   - bnpl_merchant_master_web_only_classified_v7_parse_safe_recovered.xlsx
   - or bnpl_merchant_master_web_only_classified_v7_parse_safe.xlsx
   - or bnpl_merchant_master_web_only_classified_v7.xlsx

Output
------
Data_Clean/bnpl_merchant_overlap_summary_final_v7.xlsx

Main sheets
-----------
- README
- Clean_Overlap
- Multi_Provider
- High_Overlap
- Provider_Matrix
- Industry_Overall
- Industry_By_Provider
- Industry_High_Overlap
- Audit_Needs_Review
- Audit_Excluded
- Audit_Merge_Missing
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


# ============================================================
# Paths
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_CLEAN = PROJECT_ROOT / "Data_Clean"

OVERLAP_PATH = DATA_CLEAN / "bnpl_merchant_overlap_summary_labeled.xlsx"

CLASSIFIED_CANDIDATES = [
    DATA_CLEAN / "bnpl_merchant_master_web_only_classified_v7_parse_safe_recovered.xlsx",
    DATA_CLEAN / "bnpl_merchant_master_web_only_classified_v7_parse_safe.xlsx",
    DATA_CLEAN / "bnpl_merchant_master_web_only_classified_v7.xlsx",
]

OUTPUT_PATH = DATA_CLEAN / "bnpl_merchant_overlap_summary_final_v7.xlsx"

PROVIDERS = ["affirm", "afterpay", "clearpay", "klarna", "zip"]


# ============================================================
# Basic utilities
# ============================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_key(value: Any) -> str:
    value = clean_text(value).lower()
    value = value.replace("&", " and ")
    value = value.replace("™", "").replace("®", "")
    value = re.sub(r"\b(official store|official|online store|store|shop)\b", "", value)
    value = re.sub(r"\b(inc|llc|ltd|limited|co|company|corp|corporation|plc)\b", "", value)
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def split_unique_semicolon(values: list[Any]) -> str:
    out: list[str] = []
    seen: set[str] = set()

    for value in values:
        if value is None:
            continue
        text = clean_text(value)
        if not text:
            continue

        parts = re.split(r"[;|,]", text)
        for part in parts:
            item = clean_text(part)
            if not item:
                continue
            key = item.lower()
            if key not in seen:
                seen.add(key)
                out.append(item)

    return "; ".join(out)


def clean_country_set(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""

    replacements = {
        "UnitedStates": "United States",
        "United States": "United States",
        "USA": "United States",
        "US": "United States",
        "UnitedKingdom": "United Kingdom",
        "United Kingdom": "United Kingdom",
        "UK": "United Kingdom",
        "Canada(En)": "Canada",
        "Canada(Fr)": "Canada",
    }

    parts = re.split(r"[;|,]", text)
    out: list[str] = []
    seen: set[str] = set()

    for part in parts:
        item = clean_text(part)
        if not item:
            continue
        item = replacements.get(item, item)
        key = item.lower()
        if key not in seen:
            seen.add(key)
            out.append(item)

    return "; ".join(out)


def bool_from_value(value: Any) -> bool:
    text = clean_text(value).lower()
    return text in {"true", "1", "yes", "y", "t"}


def sanitize_for_excel_value(value: Any) -> Any:
    """
    Prevent Excel from interpreting scraped text as a formula.
    This avoids repair prompts caused by cells starting with =, +, -, or @.
    """
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    object_cols = out.select_dtypes(include=["object"]).columns
    for col in object_cols:
        out[col] = out[col].map(sanitize_for_excel_value)
    return out


def find_classified_path() -> Path:
    for path in CLASSIFIED_CANDIDATES:
        if path.exists():
            return path

    msg = "Could not find final classified workbook. Expected one of:\n"
    msg += "\n".join(str(p) for p in CLASSIFIED_CANDIDATES)
    raise FileNotFoundError(msg)


def read_first_non_summary_sheet(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    sheet = None
    for name in xls.sheet_names:
        if "summary" not in name.lower() and "readme" not in name.lower():
            sheet = name
            break
    if sheet is None:
        sheet = xls.sheet_names[0]

    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ============================================================
# Alias / page cleaning
# ============================================================

# Merchant aliases that should be merged before overlap analysis.
ALIAS_TO_CANONICAL = {
    "applefromklarna": "Apple",
    "google": "GoogleStore",
    "googlestore": "GoogleStore",
    "loweshomeimprovement": "Lowe's",
    "lowes": "Lowe's",
    "samsclub": "Sam'sClub",
    "thelego": "LEGO",
    "legostore": "LEGO",
    "bookingcom": "Booking.com",
    "hotelscom": "Hotels.com",
    "expediaflights": "ExpediaFlights",
}

# Rows that are known non-merchant/provider-internal pages.
NON_MERCHANT_KEYS = {
    "australia",
    "canadaen",
    "canadafr",
    "canada",
    "unitedkingdom",
    "unitedstates",
    "moreinformationaboutourcookiepolicy",
    "skiptomaincontent",
    "shopnow",
    "howitworks",
    "fordevelopers",
    "diversityandinclusion",
    "learnmore",
    "readmore",
    "cookiepolicy",
    "privacychoices",
}

BNPL_INTERNAL_KEYS = {
    "affirmcomlenders",
    "affirmlenders",
    "affirmcard",
    "affirmmoney",
    "affirmcares",
}


def choose_name_for_key(row: pd.Series) -> str:
    for col in ["canonical_company_name", "normalized_company_name", "merchant_name", "company_name", "raw_name_variants"]:
        if col in row.index and clean_text(row.get(col)):
            return clean_text(row.get(col))
    return ""


def clean_company_key_from_row(row: pd.Series) -> str:
    original_name = choose_name_for_key(row)
    key = normalize_key(original_name)
    canonical = ALIAS_TO_CANONICAL.get(key)
    if canonical:
        return normalize_key(canonical)
    return key


def canonical_name_from_group(group: pd.DataFrame) -> str:
    keys = [clean_company_key_from_row(row) for _, row in group.iterrows()]
    key = keys[0] if keys else ""

    for alias_key, canonical in ALIAS_TO_CANONICAL.items():
        if normalize_key(canonical) == key:
            return canonical

    candidates: list[str] = []
    for col in ["canonical_company_name", "normalized_company_name", "merchant_name", "company_name"]:
        if col in group.columns:
            candidates += [clean_text(v) for v in group[col].tolist() if clean_text(v)]

    if candidates:
        # Prefer the shortest readable display name.
        candidates = sorted(set(candidates), key=lambda x: (len(x), x.lower()))
        return candidates[0]

    return key


# ============================================================
# Overlap aggregation
# ============================================================

def infer_has_provider(row: pd.Series, provider: str) -> bool:
    col = f"has_{provider}"
    if col in row.index:
        return bool_from_value(row.get(col))

    provider_set = clean_text(row.get("provider_set", "")).lower()
    return provider in provider_set


def aggregate_overlap(overlap_df: pd.DataFrame) -> pd.DataFrame:
    df = overlap_df.copy()
    df["clean_company_key"] = df.apply(clean_company_key_from_row, axis=1)

    rows: list[dict[str, Any]] = []

    for key, group in df.groupby("clean_company_key", dropna=False):
        canonical_name = canonical_name_from_group(group)

        has_provider = {}
        for provider in PROVIDERS:
            has_provider[provider] = any(infer_has_provider(row, provider) for _, row in group.iterrows())

        provider_list = [p for p in PROVIDERS if has_provider[p]]
        provider_count = len(provider_list)

        clean_country = split_unique_semicolon([clean_country_set(v) for v in group.get("country_set", pd.Series(dtype=str)).tolist()])

        raw_name_cols = []
        for col in ["raw_name_variants", "canonical_company_name", "normalized_company_name"]:
            if col in group.columns:
                raw_name_cols.extend(group[col].tolist())

        merchant_urls = group["merchant_urls"].tolist() if "merchant_urls" in group.columns else []
        source_files = group["source_files"].tolist() if "source_files" in group.columns else []

        row = {
            "clean_company_key": key,
            "canonical_company_name_clean": canonical_name,
            "raw_name_variants_clean": split_unique_semicolon(raw_name_cols),
            "provider_set_clean": "; ".join(provider_list),
            "provider_count_clean": provider_count,
            "clean_country_set": clean_country,
            "merchant_urls_clean": split_unique_semicolon(merchant_urls),
            "source_files_clean": split_unique_semicolon(source_files),
        }

        for provider in PROVIDERS:
            row[f"has_{provider}"] = has_provider[provider]

        # Mark deterministic exclusions before merge.
        if key in NON_MERCHANT_KEYS:
            row["premerge_exclusion"] = "Exclude / Non-Merchant Page"
            row["premerge_exclusion_reason"] = "Known navigation / locale / policy page."
        elif key in BNPL_INTERNAL_KEYS:
            row["premerge_exclusion"] = "Exclude / BNPL Internal Page"
            row["premerge_exclusion_reason"] = "Known BNPL provider internal page."
        else:
            row["premerge_exclusion"] = ""
            row["premerge_exclusion_reason"] = ""

        rows.append(row)

    out = pd.DataFrame(rows)
    out = out.sort_values(["provider_count_clean", "canonical_company_name_clean"], ascending=[False, True])
    return out


# ============================================================
# Classification aggregation
# ============================================================

CONFIDENCE_SCORE = {
    "high": 3,
    "medium": 2,
    "low": 1,
    "": 0,
}

STATUS_SCORE = {
    "official_website_verified": 4,
    "official_website_verified_plus_search_fallback": 3,
    "search_verified_website_unreachable": 2,
    "website_parse_failed_search_fallback": 2,
    "search_only_no_official_url": 1,
    "needs_manual_review": 0,
}


def get_classification_key(row: pd.Series) -> str:
    return clean_company_key_from_row(row)


def choose_best_classification(group: pd.DataFrame) -> pd.Series:
    g = group.copy()

    if "verified_broad_industry" not in g.columns:
        raise ValueError("Classified workbook must contain verified_broad_industry.")

    # Prefer non-empty classifications.
    g["broad_clean"] = g["verified_broad_industry"].map(clean_text)
    g = g[g["broad_clean"] != ""].copy()
    if g.empty:
        return group.iloc[0]

    # If a group contains a real merchant classification and an excluded page,
    # prefer the real merchant classification. If all are excluded, keep the best excluded row.
    non_excluded = g[~g["broad_clean"].str.startswith("Exclude", na=False)].copy()
    if not non_excluded.empty:
        g = non_excluded

    g["confidence_score"] = g.get("classification_confidence", "").map(lambda x: CONFIDENCE_SCORE.get(clean_text(x).lower(), 0))
    g["status_score"] = g.get("verification_status", "").map(lambda x: STATUS_SCORE.get(clean_text(x), 0))
    g["review_score"] = g.get("needs_manual_review", False).map(lambda x: 0 if bool_from_value(x) else 1)

    g = g.sort_values(["confidence_score", "status_score", "review_score"], ascending=[False, False, False])
    return g.iloc[0]


def aggregate_classifications(classified_df: pd.DataFrame) -> pd.DataFrame:
    df = classified_df.copy()
    df["clean_company_key"] = df.apply(get_classification_key, axis=1)

    rows = []
    for key, group in df.groupby("clean_company_key", dropna=False):
        best = choose_best_classification(group)

        row = {
            "clean_company_key": key,
            "verified_broad_industry": clean_text(best.get("verified_broad_industry", "")),
            "verified_sub_industry": clean_text(best.get("verified_sub_industry", "")),
            "classification_confidence": clean_text(best.get("classification_confidence", "")),
            "verification_status": clean_text(best.get("verification_status", "")),
            "needs_manual_review": bool_from_value(best.get("needs_manual_review", False)),
            "official_url": clean_text(best.get("official_url", "")),
            "final_url": clean_text(best.get("final_url", "")),
            "classification_reason": clean_text(best.get("classification_reason", "")),
            "matched_classification_rows": len(group),
        }
        rows.append(row)

    return pd.DataFrame(rows)


# ============================================================
# Final table and summaries
# ============================================================

def apply_premerge_exclusions(row: pd.Series) -> pd.Series:
    if clean_text(row.get("premerge_exclusion", "")):
        row["verified_broad_industry"] = clean_text(row.get("premerge_exclusion"))
        row["verified_sub_industry"] = "Navigation / provider-internal page"
        row["classification_confidence"] = "high"
        row["verification_status"] = (
            "exclude_provider_internal_page"
            if "BNPL" in clean_text(row.get("premerge_exclusion", ""))
            else "exclude_non_merchant_page"
        )
        row["needs_manual_review"] = False
        row["classification_reason"] = clean_text(row.get("premerge_exclusion_reason", ""))
    return row


def build_final_overlap(overlap_clean: pd.DataFrame, class_agg: pd.DataFrame) -> pd.DataFrame:
    final = overlap_clean.merge(class_agg, on="clean_company_key", how="left", indicator=True)
    final = final.apply(apply_premerge_exclusions, axis=1)

    final["is_excluded"] = final["verified_broad_industry"].fillna("").astype(str).str.startswith("Exclude")
    final["is_multi_provider"] = final["provider_count_clean"] >= 2
    final["is_high_overlap"] = final["provider_count_clean"] >= 3

    final["classification_needs_review"] = (
        final["needs_manual_review"].fillna(False).astype(bool)
        | final["classification_confidence"].fillna("").astype(str).str.lower().eq("low")
        | final["verified_broad_industry"].fillna("").astype(str).eq("Other / Unclassified")
        | final["verified_broad_industry"].fillna("").astype(str).eq("")
    )

    # Final column order.
    base_cols = [
        "clean_company_key",
        "canonical_company_name_clean",
        "raw_name_variants_clean",
        "provider_set_clean",
        "provider_count_clean",
        "clean_country_set",
    ]
    provider_cols = [f"has_{p}" for p in PROVIDERS]
    class_cols = [
        "verified_broad_industry",
        "verified_sub_industry",
        "classification_confidence",
        "verification_status",
        "needs_manual_review",
        "classification_needs_review",
        "is_excluded",
        "is_multi_provider",
        "is_high_overlap",
        "official_url",
        "final_url",
        "classification_reason",
        "matched_classification_rows",
        "_merge",
    ]
    misc_cols = ["merchant_urls_clean", "source_files_clean"]

    ordered = [c for c in base_cols + provider_cols + class_cols + misc_cols if c in final.columns]
    remaining = [c for c in final.columns if c not in ordered and not c.startswith("premerge_")]
    final = final[ordered + remaining]

    final = final.sort_values(
        ["is_excluded", "provider_count_clean", "canonical_company_name_clean"],
        ascending=[True, False, True],
    )

    return final


def provider_overlap_matrix(final: pd.DataFrame) -> pd.DataFrame:
    active = final[~final["is_excluded"]].copy()

    matrix = pd.DataFrame(index=PROVIDERS, columns=PROVIDERS, dtype=int)

    for p1 in PROVIDERS:
        for p2 in PROVIDERS:
            matrix.loc[p1, p2] = int((active[f"has_{p1}"] & active[f"has_{p2}"]).sum())

    matrix.index.name = "provider"
    return matrix.reset_index()


def industry_overall(final: pd.DataFrame) -> pd.DataFrame:
    active = final[~final["is_excluded"]].copy()

    out = (
        active.groupby(["verified_broad_industry"], dropna=False)
        .size()
        .reset_index(name="merchant_count")
        .sort_values("merchant_count", ascending=False)
    )
    out["share_of_active_merchants"] = out["merchant_count"] / out["merchant_count"].sum()
    return out


def industry_by_provider(final: pd.DataFrame) -> pd.DataFrame:
    active = final[~final["is_excluded"]].copy()

    rows = []
    for provider in PROVIDERS:
        temp = active[active[f"has_{provider}"]].copy()
        counts = temp.groupby("verified_broad_industry", dropna=False).size().reset_index(name="merchant_count")
        total = counts["merchant_count"].sum()
        counts["provider"] = provider
        counts["provider_total"] = total
        counts["share_within_provider"] = counts["merchant_count"] / total if total else 0
        rows.append(counts)

    if not rows:
        return pd.DataFrame(columns=["provider", "verified_broad_industry", "merchant_count", "provider_total", "share_within_provider"])

    out = pd.concat(rows, ignore_index=True)
    out = out[["provider", "verified_broad_industry", "merchant_count", "provider_total", "share_within_provider"]]
    out = out.sort_values(["provider", "merchant_count"], ascending=[True, False])
    return out


def industry_high_overlap(final: pd.DataFrame) -> pd.DataFrame:
    active = final[(~final["is_excluded"]) & (final["is_high_overlap"])].copy()
    out = (
        active.groupby(["verified_broad_industry"], dropna=False)
        .size()
        .reset_index(name="high_overlap_merchant_count")
        .sort_values("high_overlap_merchant_count", ascending=False)
    )
    total = out["high_overlap_merchant_count"].sum()
    out["share_of_high_overlap"] = out["high_overlap_merchant_count"] / total if total else 0
    return out


def build_readme(final: pd.DataFrame, classified_path: Path) -> pd.DataFrame:
    active = final[~final["is_excluded"]]
    rows = [
        ["output_created_by", "build_final_bnpl_overlap_tables.py"],
        ["source_overlap_file", str(OVERLAP_PATH)],
        ["source_classified_file", str(classified_path)],
        ["total_overlap_rows_after_alias_merge", len(final)],
        ["active_merchant_rows", len(active)],
        ["excluded_rows", int(final["is_excluded"].sum())],
        ["multi_provider_merchants", int(active["is_multi_provider"].sum())],
        ["high_overlap_merchants_provider_count_ge_3", int(active["is_high_overlap"].sum())],
        ["classification_needs_review_rows", int(final["classification_needs_review"].sum())],
        ["notes", "Industry labels refreshed from final web-based classifier; alias and non-merchant filters applied before overlap analysis."],
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def write_workbook(
    output_path: Path,
    readme: pd.DataFrame,
    final: pd.DataFrame,
    provider_matrix: pd.DataFrame,
    ind_overall: pd.DataFrame,
    ind_provider: pd.DataFrame,
    ind_high: pd.DataFrame,
) -> None:
    multi = final[(~final["is_excluded"]) & (final["is_multi_provider"])].copy()
    high = final[(~final["is_excluded"]) & (final["is_high_overlap"])].copy()
    audit_review = final[final["classification_needs_review"]].copy()
    audit_excluded = final[final["is_excluded"]].copy()
    audit_missing = final[final["_merge"].eq("left_only")].copy() if "_merge" in final.columns else pd.DataFrame()

    sheets = {
        "README": readme,
        "Clean_Overlap": final,
        "Multi_Provider": multi,
        "High_Overlap": high,
        "Provider_Matrix": provider_matrix,
        "Industry_Overall": ind_overall,
        "Industry_By_Provider": ind_provider,
        "Industry_High_Overlap": ind_high,
        "Audit_Needs_Review": audit_review,
        "Audit_Excluded": audit_excluded,
        "Audit_Merge_Missing": audit_missing,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_df = sanitize_dataframe_for_excel(df)
            safe_df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.book[sheet_name]

            # Freeze top row
            ws.freeze_panes = "A2"

            # Basic autofilter
            if ws.max_row >= 1 and ws.max_column >= 1:
                ws.auto_filter.ref = ws.dimensions

            # Conservative column width formatting.
            for col_cells in ws.columns:
                col_letter = col_cells[0].column_letter
                max_len = 0
                for cell in col_cells[:200]:
                    value = cell.value
                    if value is None:
                        continue
                    max_len = max(max_len, len(str(value)))
                width = min(max(max_len + 2, 10), 45)
                ws.column_dimensions[col_letter].width = width

    print(f"Saved final overlap workbook: {output_path}")


def main() -> None:
    if not OVERLAP_PATH.exists():
        raise FileNotFoundError(f"Overlap workbook not found: {OVERLAP_PATH}")

    classified_path = find_classified_path()

    print(f"Reading overlap workbook: {OVERLAP_PATH}")
    overlap_df = read_first_non_summary_sheet(OVERLAP_PATH)
    print(f"Overlap rows: {len(overlap_df)}")

    print(f"Reading classified workbook: {classified_path}")
    classified_df = read_first_non_summary_sheet(classified_path)
    print(f"Classified rows: {len(classified_df)}")

    print("Aggregating overlap table with alias cleaning...")
    overlap_clean = aggregate_overlap(overlap_df)
    print(f"Rows after alias merge: {len(overlap_clean)}")

    print("Aggregating final classification labels...")
    class_agg = aggregate_classifications(classified_df)
    print(f"Unique classification keys: {len(class_agg)}")

    print("Merging final classification into overlap table...")
    final = build_final_overlap(overlap_clean, class_agg)

    print("Building summary sheets...")
    readme = build_readme(final, classified_path)
    matrix = provider_overlap_matrix(final)
    ind_overall = industry_overall(final)
    ind_provider = industry_by_provider(final)
    ind_high = industry_high_overlap(final)

    write_workbook(
        output_path=OUTPUT_PATH,
        readme=readme,
        final=final,
        provider_matrix=matrix,
        ind_overall=ind_overall,
        ind_provider=ind_provider,
        ind_high=ind_high,
    )

    print("\nDone.")
    print(f"Output: {OUTPUT_PATH}")
    print("\nKey counts:")
    print(readme.to_string(index=False))


if __name__ == "__main__":
    main()
