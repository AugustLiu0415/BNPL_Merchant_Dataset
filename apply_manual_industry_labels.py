from pathlib import Path
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_CLEAN = PROJECT_ROOT / "Data_Clean"

MASTER_FILE = DATA_CLEAN / "bnpl_merchant_master_long.xlsx"
LABEL_FILE = DATA_CLEAN / "manual_industry_labels.xlsx"

LABELED_MASTER_EXCEL = DATA_CLEAN / "bnpl_merchant_master_long_labeled.xlsx"
LABELED_MASTER_CSV = DATA_CLEAN / "bnpl_merchant_master_long_labeled.csv"

LABELED_OVERLAP_EXCEL = DATA_CLEAN / "bnpl_merchant_overlap_summary_labeled.xlsx"
LABELED_OVERLAP_CSV = DATA_CLEAN / "bnpl_merchant_overlap_summary_labeled.csv"

LABELED_INDUSTRY_EXCEL = DATA_CLEAN / "bnpl_industry_by_provider_labeled.xlsx"
LABELED_INDUSTRY_CSV = DATA_CLEAN / "bnpl_industry_by_provider_labeled.csv"


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def mode_or_default(series: pd.Series, default: str = "Other / Unclassified") -> str:
    values = [clean_text(x) for x in series if clean_text(x)]
    if not values:
        return default

    mode = pd.Series(values).mode()
    if mode.empty:
        return values[0]

    return mode.iloc[0]


def build_overlap_summary(master_long: pd.DataFrame) -> pd.DataFrame:
    providers = sorted(master_long["provider"].dropna().unique().tolist())

    summary = (
        master_long.groupby("normalized_company_name")
        .agg(
            canonical_company_name=("canonical_company_name", "first"),
            raw_name_variants=(
                "raw_company_name",
                lambda x: " | ".join(sorted(set(map(clean_text, x)))),
            ),
            provider_set=(
                "provider",
                lambda x: "; ".join(sorted(set(map(clean_text, x)))),
            ),
            provider_count=("provider", lambda x: len(set(map(clean_text, x)))),
            country_set=(
                "country",
                lambda x: "; ".join(sorted(set(map(clean_text, x)))),
            ),
            broad_industry=("broad_industry", mode_or_default),
            sub_industry=("sub_industry", lambda x: mode_or_default(x, default="")),
            merchant_urls=(
                "merchant_url",
                lambda x: " | ".join(sorted(set([clean_text(v) for v in x if clean_text(v)]))),
            ),
            source_files=(
                "source_file",
                lambda x: "; ".join(sorted(set(map(clean_text, x)))),
            ),
        )
        .reset_index()
    )

    for provider in providers:
        provider_col = f"has_{provider.lower().replace(' ', '_')}"
        temp = (
            master_long[master_long["provider"] == provider]
            .groupby("normalized_company_name")
            .size()
            .reset_index(name=provider_col)
        )
        temp[provider_col] = 1

        summary = summary.merge(
            temp[["normalized_company_name", provider_col]],
            on="normalized_company_name",
            how="left",
        )
        summary[provider_col] = summary[provider_col].fillna(0).astype(int)

    summary = summary.sort_values(
        ["provider_count", "canonical_company_name"],
        ascending=[False, True],
    ).reset_index(drop=True)

    return summary


def build_industry_by_provider(master_long: pd.DataFrame) -> pd.DataFrame:
    table = (
        master_long.drop_duplicates(
            subset=["normalized_company_name", "provider", "broad_industry"]
        )
        .pivot_table(
            index="broad_industry",
            columns="provider",
            values="normalized_company_name",
            aggfunc="nunique",
            fill_value=0,
        )
        .reset_index()
    )

    provider_cols = [col for col in table.columns if col != "broad_industry"]
    table["total_unique_merchants"] = table[provider_cols].sum(axis=1)

    table = table.sort_values(
        "total_unique_merchants",
        ascending=False,
    ).reset_index(drop=True)

    return table


def main() -> None:
    if not MASTER_FILE.exists():
        raise FileNotFoundError(f"Cannot find master file: {MASTER_FILE}")

    if not LABEL_FILE.exists():
        raise FileNotFoundError(f"Cannot find manual label file: {LABEL_FILE}")

    master = pd.read_excel(MASTER_FILE)
    labels = pd.read_excel(LABEL_FILE, sheet_name="manual_labels")

    required_cols = [
        "normalized_company_name",
        "final_broad_industry",
        "final_sub_industry",
    ]

    for col in required_cols:
        if col not in labels.columns:
            raise ValueError(f"Missing column in manual label file: {col}")

    labels = labels[
        [
            "normalized_company_name",
            "final_broad_industry",
            "final_sub_industry",
        ]
    ].copy()

    labels["final_broad_industry"] = labels["final_broad_industry"].apply(clean_text)
    labels["final_sub_industry"] = labels["final_sub_industry"].apply(clean_text)

    labels = labels[labels["final_broad_industry"] != ""].copy()
    labels = labels.drop_duplicates(subset=["normalized_company_name"])

    labeled = master.merge(
        labels,
        on="normalized_company_name",
        how="left",
    )

    broad_mask = labeled["final_broad_industry"].notna() & (
        labeled["final_broad_industry"].astype(str).str.strip() != ""
    )

    sub_mask = labeled["final_sub_industry"].notna() & (
        labeled["final_sub_industry"].astype(str).str.strip() != ""
    )

    labeled.loc[broad_mask, "broad_industry"] = labeled.loc[
        broad_mask, "final_broad_industry"
    ]

    labeled.loc[sub_mask, "sub_industry"] = labeled.loc[
        sub_mask, "final_sub_industry"
    ]

    labeled = labeled.drop(columns=["final_broad_industry", "final_sub_industry"])

    overlap = build_overlap_summary(labeled)
    industry = build_industry_by_provider(labeled)

    labeled.to_excel(LABELED_MASTER_EXCEL, index=False)
    labeled.to_csv(LABELED_MASTER_CSV, index=False)

    overlap.to_excel(LABELED_OVERLAP_EXCEL, index=False)
    overlap.to_csv(LABELED_OVERLAP_CSV, index=False)

    industry.to_excel(LABELED_INDUSTRY_EXCEL, index=False)
    industry.to_csv(LABELED_INDUSTRY_CSV, index=False)

    print("Manual industry labels applied.")
    print(f"Saved: {LABELED_MASTER_EXCEL}")
    print(f"Saved: {LABELED_OVERLAP_EXCEL}")
    print(f"Saved: {LABELED_INDUSTRY_EXCEL}")

    print("\nUpdated industry distribution:")
    print(industry.to_string(index=False))


if __name__ == "__main__":
    main()