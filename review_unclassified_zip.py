from pathlib import Path
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_CLEAN = PROJECT_ROOT / "Data_Clean"

input_file = DATA_CLEAN / "bnpl_merchant_master_long.xlsx"
output_file = DATA_CLEAN / "zip_unclassified_review.xlsx"

df = pd.read_excel(input_file)

zip_unclassified = df[
    (df["provider"] == "Zip")
    & (df["broad_industry"] == "Other / Unclassified")
].copy()

cols = [
    "canonical_company_name",
    "raw_company_name",
    "normalized_company_name",
    "provider",
    "country",
    "merchant_url",
    "source_page",
    "availability_type",
    "category_raw",
]

existing_cols = [col for col in cols if col in zip_unclassified.columns]

zip_unclassified = zip_unclassified[existing_cols].drop_duplicates()

zip_unclassified.to_excel(output_file, index=False)

print(f"Saved Zip unclassified review file to: {output_file}")
print(f"Total Zip unclassified rows: {len(zip_unclassified)}")
print(zip_unclassified.head(30).to_string(index=False))