# Food Delivery Data Cleaning | Python Portfolio Project
# Tools: Python, Pandas, NumPy
# Input: Food_Dashboard.xlsx (place beside this script or update input_file)
# Output: Cleaned_Food_Delivery_Data.xlsx

from pathlib import Path
import pandas as pd
import numpy as np

input_file = Path("Food_Dashboard.xlsx")
output_file = Path("Cleaned_Food_Delivery_Data.xlsx")

if not input_file.exists():
    raise FileNotFoundError(
        f"Could not find {input_file}. Place the Excel file beside this script "
        "or update input_file with the correct path."
    )

df = pd.read_excel(input_file)
print(f"Loaded: {df.shape[0]:,} rows × {df.shape[1]} columns")

# 1. Remove exact duplicate rows
rows_before = len(df)
df = df.drop_duplicates().copy()
print(f"Removed {rows_before - len(df):,} exact duplicate rows.")

# 2. Standardize date and time fields
if "order_date" in df.columns:
    df["order_date"] = pd.to_datetime(df["order_date"], format="mixed", errors="coerce")

for col in ["order_time", "delivered_at"]:
    if col in df.columns:
        parsed = pd.to_datetime(df[col].astype("string"), errors="coerce")
        df[col] = parsed.dt.strftime("%H:%M:%S")

# 3. Convert numeric columns
numeric_columns = [
    "item_count", "order_subtotal_inr", "discount_inr", "delivery_fee_inr",
    "tax_inr", "tip_inr", "order_total_inr", "distance_km",
    "prep_time_min", "delivery_time_min", "order_rating"
]
for col in numeric_columns:
    if col in df.columns:
        extracted = df[col].astype("string").str.extract(r"(-?\d+\.?\d*)", expand=False)
        df[col] = pd.to_numeric(extracted, errors="coerce")

# 4. Exclude zero or negative item counts and delivery distances.
# Missing values are retained for review; they are not assumed to be zero.
for col in ["item_count", "distance_km"]:
    if col in df.columns:
        before = len(df)
        df = df[df[col].isna() | (df[col] > 0)].copy()
        print(f"{col}: removed {before - len(df):,} invalid rows.")

# 5. Trim text and handle selected missing values
for col in df.select_dtypes(include=["object", "string"]).columns:
    df[col] = df[col].astype("string").str.strip()

if "promo_code" in df.columns:
    df["promo_code"] = df["promo_code"].fillna("No_Promo_Code")
if "customer_note" in df.columns:
    df["customer_note"] = df["customer_note"].fillna("No_Note")

for col in df.select_dtypes(include=["object", "string"]).columns:
    df[col] = df[col].fillna("Unknown")

# 6. Standardize status text only; do not infer business outcomes.
for col in ["order_status", "payment_status"]:
    if col in df.columns:
        df[col] = df[col].astype("string").str.strip().str.title()
        print(f"\n{col} categories:\n{df[col].value_counts(dropna=False)}")

# 7. Impute missing order ratings with the dataset mean.
if "order_rating" in df.columns:
    rating_mean = df["order_rating"].mean()
    if pd.notna(rating_mean):
        df["order_rating"] = df["order_rating"].fillna(round(rating_mean, 1))

# 8. Quality check and export
print(f"Final shape: {df.shape}")
print("Remaining duplicates:", df.duplicated().sum())
print("Missing values by column:\n", df.isna().sum())

df.to_excel(output_file, index=False)
print(f"Cleaned workbook saved to: {output_file.resolve()}")
