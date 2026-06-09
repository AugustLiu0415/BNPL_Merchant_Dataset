import re
from pathlib import Path
from typing import Optional

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_RAW = PROJECT_ROOT / "Data_Raw"
DATA_CLEAN = PROJECT_ROOT / "Data_Clean"
DATA_CLEAN.mkdir(exist_ok=True)


INPUT_FILES = [
    {
        "provider": "Affirm",
        "country": "United States",
        "paths": [
            DATA_RAW / "affirm_applepay_merchants.xlsx",
            DATA_RAW / "affirm_applepay_merchants.csv",
        ],
    },
    {
        "provider": "Klarna",
        "country": "United States",
        "paths": [
            DATA_RAW / "klarna_us_stores.xlsx",
            DATA_RAW / "klarna_us_stores.csv",
        ],
    },
    {
        "provider": "Afterpay",
        "country": "United States",
        "paths": [
            DATA_RAW / "afterpay_us_stores.xlsx",
            DATA_RAW / "afterpay_us_stores.csv",
        ],
    },
    {
        "provider": "Clearpay",
        "country": "United Kingdom",
        "paths": [
            DATA_RAW / "clearpay_uk_retailer_pages.xlsx",
            DATA_RAW / "clearpay_uk_retailer_pages.csv",
        ],
    },
    {
        "provider": "Zip",
        "country": "United States",
        "paths": [
            DATA_RAW / "zip_us_stores.xlsx",
            DATA_RAW / "zip_us_stores.csv",
        ],
    },
]


LONG_OUTPUT_CSV = DATA_CLEAN / "bnpl_merchant_master_long.csv"
LONG_OUTPUT_EXCEL = DATA_CLEAN / "bnpl_merchant_master_long.xlsx"

OVERLAP_OUTPUT_CSV = DATA_CLEAN / "bnpl_merchant_overlap_summary.csv"
OVERLAP_OUTPUT_EXCEL = DATA_CLEAN / "bnpl_merchant_overlap_summary.xlsx"

INDUSTRY_PROVIDER_OUTPUT_CSV = DATA_CLEAN / "bnpl_industry_by_provider.csv"
INDUSTRY_PROVIDER_OUTPUT_EXCEL = DATA_CLEAN / "bnpl_industry_by_provider.xlsx"

PROVIDER_MATRIX_OUTPUT_CSV = DATA_CLEAN / "bnpl_provider_overlap_matrix.csv"
PROVIDER_MATRIX_OUTPUT_EXCEL = DATA_CLEAN / "bnpl_provider_overlap_matrix.xlsx"


def find_existing_file(paths: list[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None


def read_input_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsx":
        return pd.read_excel(path)

    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)

    raise ValueError(f"Unsupported file format: {path}")


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(col)
        .strip()
        .lower()
        .replace(" ", "_")
        .replace("-", "_")
        for col in df.columns
    ]
    return df


def clean_text(text: object) -> str:
    if text is None:
        return ""
    
    try: 
        if pd.isna(text):
                return""
    except (TypeError, ValueError):
        pass

    text = str(text).strip()
    text = re.sub(r"\s+", "", text)
    return text.strip()

    text = str(text).strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_company_name(name: object) -> str:
    """
    Create a normalized matching key.

    Example:
    "DICK'S Sporting Goods" -> "dickssportinggoods"
    "Apple from Klarna" -> "apple"
    """
    name = clean_text(name).lower()

    # Remove common platform-specific wording.
    remove_phrases = [
        " from klarna",
        " official store",
        " official",
        " online store",
        " store",
        " usa",
        " us",
    ]

    for phrase in remove_phrases:
        if name.endswith(phrase):
            name = name[: -len(phrase)]

    # Standardize common symbols.
    name = name.replace("&", " and ")

    # Remove trademark symbols.
    name = name.replace("™", "")
    name = name.replace("®", "")
    name = name.replace("©", "")

    # Remove common corporate suffixes.
    suffixes = [
        " inc",
        " incorporated",
        " llc",
        " ltd",
        " limited",
        " co",
        " company",
        " corp",
        " corporation",
        " plc",
    ]

    for suffix in suffixes:
        if name.endswith(suffix):
            name = name[: -len(suffix)]

    # Remove punctuation and spaces.
    name = re.sub(r"[^a-z0-9]+", "", name)

    return name


def choose_canonical_name(names: pd.Series) -> str:
    """
    Pick a readable display name for each normalized merchant group.
    We choose the shortest common-looking name, not necessarily perfect.
    """
    clean_names = sorted(
        {clean_text(x) for x in names if clean_text(x)},
        key=lambda x: (len(x), x.lower()),
    )

    if not clean_names:
        return ""

    return clean_names[0]


def get_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def classify_industry(company_name: str, category: str = "", availability_type: str = "") -> str:
    """
    Initial rule-based industry classification.

    This is intentionally simple and transparent.
    Later we can improve it manually or with an API/classification model.
    """
    text = f"{company_name} {category} {availability_type}".lower()

    rules = {
        "Marketplace / General Retail": [
            "marketplace", "amazon", "walmart", "target", "ebay", "best buy",
            "costco", "sam's club", "sams club", "department store", "general retail",
        ],
        "Apparel & Fashion": [
            "fashion", "apparel", "clothing", "clothes", "shoes", "sneakers",
            "footwear", "boutique", "dress", "jewelry", "jewellery", "watch",
            "bags", "handbag", "accessories", "nike", "adidas", "lululemon",
            "american eagle", "asos", "zara", "h&m", "aritzia", "anthropologie",
            "urban outfitters", "levi", "gap", "old navy", "shein", "boohoo",
            "allbirds", "under armour",
        ],
        "Sports & Outdoor": [
            "sports", "outdoor", "camping", "hiking", "cycling", "bike",
            "fitness", "golf", "fishing", "hunting", "rei", "dick's",
            "dicks sporting goods", "patagonia", "arc'teryx", "arcteryx",
            "the north face", "columbia", "academy sports",
        ],
        "Electronics": [
            "electronics", "computer", "laptop", "phone", "mobile", "camera",
            "audio", "headphone", "gaming", "apple", "google store", "samsung",
            "anker", "bose", "audeze", "newegg", "microsoft",
        ],
        "Home & Furniture": [
            "home", "furniture", "mattress", "bed", "sofa", "decor",
            "kitchen", "bath", "lighting", "wayfair", "ikea", "article",
            "crate and barrel", "west elm", "pottery barn", "overstock",
            "home depot", "lowe's", "lowes",
        ],
        "Beauty & Personal Care": [
            "beauty", "cosmetic", "cosmetics", "skincare", "skin care",
            "hair", "makeup", "fragrance", "perfume", "sephora", "ulta",
            "dermalogica", "estee lauder", "glossier",
        ],
        "Travel & Ticketing": [
            "travel", "hotel", "flight", "airline", "booking", "expedia",
            "airbnb", "delta", "ticket", "ticketmaster", "event",
        ],
        "Food & Delivery": [
            "food", "restaurant", "delivery", "doordash", "ubereats",
            "uber eats", "grubhub", "meal", "grocery", "coffee",
        ],
        "Baby & Kids": [
            "baby", "kids", "children", "child", "toy", "toys", "carter",
            "buybuy baby", "stroller", "nursery",
        ],
        "Automotive": [
            "auto", "car", "tires", "vehicle", "motor", "automotive",
            "autozone", "advance auto", "autow",
        ],
        "Pet": [
            "pet", "pets", "dog", "cat", "chewy", "petco", "petsmart",
        ],
        "Health & Wellness": [
            "health", "wellness", "pharmacy", "medical", "fitness", "vitamin",
            "supplement", "contacts", "1-800 contacts", "1800 contacts",
        ],
        "Education / Books": [
            "book", "books", "education", "course", "school", "university",
            "textbook", "barnes", "noble",
        ],
        "Entertainment / Gaming": [
            "game", "gaming", "entertainment", "streaming", "music", "movie",
            "playstation", "xbox", "nintendo",
        ],
    }

    for industry, keywords in rules.items():
        if any(keyword in text for keyword in keywords):
            return industry

    return "Other / Unclassified"


def load_and_standardize_all_files() -> pd.DataFrame:
    all_frames = []

    for item in INPUT_FILES:
        provider = item["provider"]
        default_country = item["country"]
        path = find_existing_file(item["paths"])

        if path is None:
            print(f"WARNING: No input file found for {provider}. Skipping.")
            continue

        print(f"Reading {provider}: {path}")

        df = read_input_file(path)
        df = clean_column_names(df)

        company_col = get_first_existing_column(
            df,
            [
                "company_name",
                "merchant_name",
                "name",
                "retailer_name",
                "store_name",
            ],
        )

        if company_col is None:
            raise ValueError(
                f"Cannot find company name column in {path}. Columns are: {list(df.columns)}"
            )

        merchant_url_col = get_first_existing_column(
            df,
            [
                "merchant_url",
                "merchant",
                "store_url",
                "url",
                "link",
            ],
        )

        category_col = get_first_existing_column(
            df,
            [
                "category",
                "merchant_category",
                "industry",
            ],
        )

        source_page_col = get_first_existing_column(
            df,
            [
                "source_page",
                "source",
            ],
        )

        availability_col = get_first_existing_column(
            df,
            [
                "availability_type",
                "availability",
                "type",
            ],
        )

        scraped_at_col = get_first_existing_column(
            df,
            [
                "scraped_at",
                "scrape_time",
                "date_scraped",
            ],
        )

        country_col = get_first_existing_column(
            df,
            [
                "country",
                "market",
            ],
        )

        standardized = pd.DataFrame()
        standardized["raw_company_name"] = df[company_col].apply(clean_text)
        standardized["normalized_company_name"] = standardized["raw_company_name"].apply(
            normalize_company_name
        )

        standardized["provider"] = provider
        standardized["country"] = (
            df[country_col].apply(clean_text) if country_col else default_country
        )

        standardized["merchant_url"] = (
            df[merchant_url_col].apply(clean_text) if merchant_url_col else ""
        )
        standardized["category_raw"] = (
            df[category_col].apply(clean_text) if category_col else ""
        )
        standardized["source_page"] = (
            df[source_page_col].apply(clean_text) if source_page_col else ""
        )
        standardized["availability_type"] = (
            df[availability_col].apply(clean_text) if availability_col else ""
        )
        standardized["scraped_at"] = (
            df[scraped_at_col].apply(clean_text) if scraped_at_col else ""
        )

        standardized["source_file"] = path.name

        standardized["broad_industry"] = standardized.apply(
            lambda row: classify_industry(
                company_name=row["raw_company_name"],
                category=row["category_raw"],
                availability_type=row["availability_type"],
            ),
            axis=1,
        )

        standardized["sub_industry"] = standardized["category_raw"].where(
            standardized["category_raw"] != "",
            standardized["broad_industry"],
        )

        # Remove empty or clearly bad rows.
        standardized = standardized[
            standardized["normalized_company_name"].str.len() > 0
        ].copy()

        all_frames.append(standardized)

    if not all_frames:
        raise ValueError("No input files were loaded. Please check Data_Raw file names.")

    master_long = pd.concat(all_frames, ignore_index=True)

    # Drop exact duplicate rows.
    master_long = master_long.drop_duplicates(
        subset=[
            "normalized_company_name",
            "provider",
            "country",
            "merchant_url",
        ]
    ).reset_index(drop=True)

    # Add canonical name based on normalized groups.
    canonical_map = (
        master_long.groupby("normalized_company_name")["raw_company_name"]
        .apply(choose_canonical_name)
        .to_dict()
    )

    master_long["canonical_company_name"] = master_long[
        "normalized_company_name"
    ].map(canonical_map)

    # Reorder columns.
    ordered_cols = [
        "canonical_company_name",
        "raw_company_name",
        "normalized_company_name",
        "provider",
        "country",
        "broad_industry",
        "sub_industry",
        "category_raw",
        "availability_type",
        "merchant_url",
        "source_page",
        "source_file",
        "scraped_at",
    ]

    master_long = master_long[ordered_cols]

    return master_long


def build_overlap_summary(master_long: pd.DataFrame) -> pd.DataFrame:
    providers = sorted(master_long["provider"].dropna().unique().tolist())

    grouped = master_long.groupby("normalized_company_name")

    summary = grouped.agg(
        canonical_company_name=("canonical_company_name", "first"),
        raw_name_variants=("raw_company_name", lambda x: " | ".join(sorted(set(x)))),
        provider_set=("provider", lambda x: "; ".join(sorted(set(x)))),
        provider_count=("provider", lambda x: len(set(x))),
        country_set=("country", lambda x: "; ".join(sorted(set(x)))),
        broad_industry=("broad_industry", lambda x: x.mode().iloc[0] if not x.mode().empty else "Other / Unclassified"),
        sub_industry=("sub_industry", lambda x: x.mode().iloc[0] if not x.mode().empty else ""),
        merchant_urls=("merchant_url", lambda x: " | ".join(sorted(set([v for v in x if v])))),
        source_files=("source_file", lambda x: "; ".join(sorted(set(x)))),
    ).reset_index()

    # Provider dummy variables.
    for provider in providers:
        provider_norm = provider.lower().replace(" ", "_")
        provider_members = (
            master_long[master_long["provider"] == provider]
            .groupby("normalized_company_name")
            .size()
            .reset_index(name=f"has_{provider_norm}")
        )
        provider_members[f"has_{provider_norm}"] = 1

        summary = summary.merge(
            provider_members[["normalized_company_name", f"has_{provider_norm}"]],
            on="normalized_company_name",
            how="left",
        )

        summary[f"has_{provider_norm}"] = (
            summary[f"has_{provider_norm}"].fillna(0).astype(int)
        )

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


def build_provider_overlap_matrix(overlap_summary: pd.DataFrame) -> pd.DataFrame:
    provider_cols = [
        col for col in overlap_summary.columns
        if col.startswith("has_")
    ]

    labels = [col.replace("has_", "") for col in provider_cols]

    matrix = pd.DataFrame(index=labels, columns=labels, dtype=int)

    for i, col_i in enumerate(provider_cols):
        for j, col_j in enumerate(provider_cols):
            matrix.iloc[i, j] = int(
                ((overlap_summary[col_i] == 1) & (overlap_summary[col_j] == 1)).sum()
            )

    matrix = matrix.reset_index().rename(columns={"index": "provider"})

    return matrix


def save_outputs(
    master_long: pd.DataFrame,
    overlap_summary: pd.DataFrame,
    industry_by_provider: pd.DataFrame,
    provider_matrix: pd.DataFrame,
) -> None:
    master_long.to_csv(LONG_OUTPUT_CSV, index=False)
    master_long.to_excel(LONG_OUTPUT_EXCEL, index=False)

    overlap_summary.to_csv(OVERLAP_OUTPUT_CSV, index=False)
    overlap_summary.to_excel(OVERLAP_OUTPUT_EXCEL, index=False)

    industry_by_provider.to_csv(INDUSTRY_PROVIDER_OUTPUT_CSV, index=False)
    industry_by_provider.to_excel(INDUSTRY_PROVIDER_OUTPUT_EXCEL, index=False)

    provider_matrix.to_csv(PROVIDER_MATRIX_OUTPUT_CSV, index=False)
    provider_matrix.to_excel(PROVIDER_MATRIX_OUTPUT_EXCEL, index=False)

    print("\nSaved cleaned outputs:")
    print(f"- {LONG_OUTPUT_CSV}")
    print(f"- {LONG_OUTPUT_EXCEL}")
    print(f"- {OVERLAP_OUTPUT_CSV}")
    print(f"- {OVERLAP_OUTPUT_EXCEL}")
    print(f"- {INDUSTRY_PROVIDER_OUTPUT_CSV}")
    print(f"- {INDUSTRY_PROVIDER_OUTPUT_EXCEL}")
    print(f"- {PROVIDER_MATRIX_OUTPUT_CSV}")
    print(f"- {PROVIDER_MATRIX_OUTPUT_EXCEL}")


def print_basic_summary(master_long: pd.DataFrame, overlap_summary: pd.DataFrame) -> None:
    print("\n========== BASIC SUMMARY ==========")

    print("\nRows by provider:")
    print(master_long.groupby("provider").size().sort_values(ascending=False))

    print("\nUnique merchants by provider:")
    print(
        master_long.groupby("provider")["normalized_company_name"]
        .nunique()
        .sort_values(ascending=False)
    )

    print("\nDistribution of number of BNPL platforms per merchant:")
    print(overlap_summary["provider_count"].value_counts().sort_index())

    print("\nTop multi-platform merchants:")
    cols = [
        "canonical_company_name",
        "provider_set",
        "provider_count",
        "broad_industry",
        "country_set",
    ]
    print(
        overlap_summary[overlap_summary["provider_count"] >= 2][cols]
        .head(30)
        .to_string(index=False)
    )


def main() -> None:
    master_long = load_and_standardize_all_files()
    overlap_summary = build_overlap_summary(master_long)
    industry_by_provider = build_industry_by_provider(master_long)
    provider_matrix = build_provider_overlap_matrix(overlap_summary)

    save_outputs(
        master_long=master_long,
        overlap_summary=overlap_summary,
        industry_by_provider=industry_by_provider,
        provider_matrix=provider_matrix,
    )

    print_basic_summary(master_long, overlap_summary)


if __name__ == "__main__":
    main()