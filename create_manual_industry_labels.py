from pathlib import Path
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_CLEAN = PROJECT_ROOT / "Data_Clean"

INPUT_FILE = DATA_CLEAN / "bnpl_merchant_overlap_summary.xlsx"
OUTPUT_FILE = DATA_CLEAN / "manual_industry_labels.xlsx"

MAX_ZIP_UNCLASSIFIED = 500


TAXONOMY = [
    {
        "broad_industry": "Apparel & Fashion",
        "example_sub_industries": "Fashion; Athletic wear; Footwear; Eyewear; Bags & Accessories; Jewelry",
        "examples": "Nike; SHEIN; ASOS; A.P.C.; Ray-Ban; Adidas",
    },
    {
        "broad_industry": "Sports & Outdoor",
        "example_sub_industries": "Outdoor gear; Fitness; Tactical; Cycling; Golf; Sporting goods",
        "examples": "REI; 511Tactical; 3VGear; Academy Sports; Dick's Sporting Goods",
    },
    {
        "broad_industry": "Electronics",
        "example_sub_industries": "Consumer electronics; Audio; Gaming; Computers; Mobile accessories",
        "examples": "Apple; Best Buy; GameStop; Bose; Anker",
    },
    {
        "broad_industry": "Home & Furniture",
        "example_sub_industries": "Furniture; Home decor; Kitchenware; Cookware; Mattress; Appliances",
        "examples": "Wayfair; Article; 360Cookware; Crate & Barrel",
    },
    {
        "broad_industry": "Beauty & Personal Care",
        "example_sub_industries": "Cosmetics; Skincare; Haircare; Fragrance; Personal care",
        "examples": "Sephora; Estee Lauder; Dermalogica; Ulta",
    },
    {
        "broad_industry": "Marketplace / General Retail",
        "example_sub_industries": "Marketplace; Department store; General retail; Discount retail",
        "examples": "Amazon; Walmart; Target; eBay",
    },
    {
        "broad_industry": "Travel & Ticketing",
        "example_sub_industries": "Flights; Hotels; Vacation rental; Event tickets; Travel booking",
        "examples": "Airbnb; Expedia; Hotels.com; Ticketmaster",
    },
    {
        "broad_industry": "Food & Delivery",
        "example_sub_industries": "Restaurant; Food delivery; Grocery; Meal kits; Coffee",
        "examples": "DoorDash; Instacart",
    },
    {
        "broad_industry": "Baby & Kids",
        "example_sub_industries": "Baby products; Kids clothing; Toys; Nursery",
        "examples": "Carter's; buybuy BABY",
    },
    {
        "broad_industry": "Automotive",
        "example_sub_industries": "Auto parts; Tires; Car accessories; Motorcycle",
        "examples": "4WheelParts; AutoZone",
    },
    {
        "broad_industry": "Pet",
        "example_sub_industries": "Pet supplies; Pet food; Pet healthcare",
        "examples": "Chewy; Petco; PetSmart",
    },
    {
        "broad_industry": "Health & Wellness",
        "example_sub_industries": "Nutrition; Fitness supplements; Pharmacy; Wellness",
        "examples": "5StarNutrition; Walgreens",
    },
    {
        "broad_industry": "Education / Books",
        "example_sub_industries": "Online training; Books; Courses; Textbooks",
        "examples": "360Training; Barnes & Noble",
    },
    {
        "broad_industry": "Entertainment / Gaming",
        "example_sub_industries": "Gaming; Streaming; Music; Events; Hobbies",
        "examples": "GameStop; Nintendo",
    },
    {
        "broad_industry": "Gifts & Flowers",
        "example_sub_industries": "Flowers; Gifts; Greeting cards; Personalized gifts",
        "examples": "1-800-Flowers",
    },
    {
        "broad_industry": "Specialty / Hobby",
        "example_sub_industries": "Crafts; Collectibles; Niche hobby; Specialty retail",
        "examples": "Small niche brands",
    },
    {
        "broad_industry": "Other / Unclassified",
        "example_sub_industries": "Unknown",
        "examples": "Use only when unclear",
    },
]


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Cannot find input file: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)

    df["provider_count"] = pd.to_numeric(
        df["provider_count"],
        errors="coerce",
    ).fillna(0).astype(int)

    for col in [
        "normalized_company_name",
        "canonical_company_name",
        "raw_name_variants",
        "provider_set",
        "country_set",
        "broad_industry",
        "sub_industry",
        "merchant_urls",
    ]:
        if col not in df.columns:
            df[col] = ""

    # Priority 1: merchants appearing on at least two BNPL platforms.
    multi_platform = df[df["provider_count"] >= 2].copy()
    multi_platform["review_reason"] = "multi-platform merchant"

    # Priority 2: a sample of Zip-only unclassified long-tail merchants.
    zip_unclassified = df[
        (df["provider_count"] == 1)
        & (df["provider_set"].astype(str).str.contains("Zip", case=False, na=False))
        & (df["broad_industry"].astype(str).str.strip().eq("Other / Unclassified"))
    ].copy()

    zip_unclassified = zip_unclassified.sort_values("canonical_company_name").head(
        MAX_ZIP_UNCLASSIFIED
    )
    zip_unclassified["review_reason"] = "Zip-only unclassified sample"

    manual = pd.concat([multi_platform, zip_unclassified], ignore_index=True)
    manual = manual.drop_duplicates(subset=["normalized_company_name"]).copy()

    manual["current_broad_industry"] = manual["broad_industry"]
    manual["current_sub_industry"] = manual["sub_industry"]

    # Pre-fill final labels if current label is already informative.
    manual["final_broad_industry"] = manual["current_broad_industry"].apply(
        lambda x: "" if clean_text(x) == "Other / Unclassified" else clean_text(x)
    )
    manual["final_sub_industry"] = manual["current_sub_industry"].apply(
        lambda x: "" if clean_text(x) == "Other / Unclassified" else clean_text(x)
    )

    manual["manual_notes"] = ""

    output_cols = [
        "normalized_company_name",
        "canonical_company_name",
        "raw_name_variants",
        "provider_set",
        "provider_count",
        "country_set",
        "current_broad_industry",
        "current_sub_industry",
        "final_broad_industry",
        "final_sub_industry",
        "review_reason",
        "merchant_urls",
        "manual_notes",
    ]

    manual = manual[output_cols]

    taxonomy_df = pd.DataFrame(TAXONOMY)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        manual.to_excel(writer, sheet_name="manual_labels", index=False)
        taxonomy_df.to_excel(writer, sheet_name="taxonomy_reference", index=False)

    print(f"Saved manual industry label file to: {OUTPUT_FILE}")
    print(f"Rows for manual review: {len(manual)}")
    print("\nNext step: open the Excel file and fill final_broad_industry and final_sub_industry.")


if __name__ == "__main__":
    main()