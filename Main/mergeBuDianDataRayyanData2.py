import pandas as pd
import re
from datetime import datetime
import time

program_start_time = time.time()
print("=== STARTING MERGE PROCESS ===")
print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

# =============================================================================
# STEP 1: Data Loading
# =============================================================================
print("STEP 1: Loading data files...")
start_time = time.time()

try:
    df_penjualan = pd.read_excel("./BAEKMI/Penjualan2024.xlsx")
    df_beli = pd.read_excel("./PembelianBuDian2024.xlsx")

    print(f"Data loaded successfully. Penjualan: {len(df_penjualan)} rows, Pembelian: {len(df_beli)} rows")
except Exception as e:
    print(f"ERROR loading files: {str(e)}")
    raise

print(f"STEP 1 completed in {time.time() - start_time:.2f} seconds\n")

# =============================================================================
# STEP 2: Data Preprocessing
# =============================================================================
print("STEP 2: Preprocessing data (cleaning, date formatting)...")
start_time = time.time()

def preprocess_data(df_penjualan, df_beli):
    print("  - Cleaning whitespace and standardizing text columns...")
    for col in ["Nama Barang", "Satuan"]:
        df_penjualan[col] = df_penjualan[col].astype(str).str.strip()
        df_beli[col] = df_beli[col].astype(str).str.strip()
    
    print("  - Formatting dates to '02 Mar 2024' format...")
    for df, name in [(df_penjualan, "Penjualan"), (df_beli, "Pembelian")]:
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors='coerce')
        invalid_dates = df["Tanggal"].isna().sum()
        if invalid_dates > 0:
            print(f"    Warning: {invalid_dates} invalid dates found in {name} data")
        df["Tanggal"] = df["Tanggal"].dt.strftime("%d %b %Y").fillna("Unknown Date")
    
    return df_penjualan, df_beli

df_penjualan, df_beli = preprocess_data(df_penjualan, df_beli)
print(f"STEP 2 completed in {time.time() - start_time:.2f} seconds\n")

# =============================================================================
# STEP 3: Partial Matching Setup
# =============================================================================
print("STEP 3: Setting up partial matching function...")
start_time = time.time()

def safe_contains(haystack, needle):
    """Case-sensitive partial matching that escapes regex special characters"""
    try:
        if pd.isna(needle) or pd.isna(haystack):
            return False
        escaped_needle = re.escape(str(needle))
        return bool(re.search(escaped_needle, str(haystack)))
    except Exception as e:
        print(f"Warning: Matching error for '{needle}' vs '{haystack}': {str(e)}")
        return False

print(f"STEP 3 completed in {time.time() - start_time:.2f} seconds\n")

# =============================================================================
# STEP 4: Merging Data
# =============================================================================
print("STEP 4: Merging datasets (this may take time)...")
start_time = time.time()
total_rows = len(df_penjualan)
merged_rows = []
match_counts = 0
no_match_counts = 0

print(f"  Processing {total_rows} sales records...")
for i, (_, penjualan_row) in enumerate(df_penjualan.iterrows()):
    if i % 100 == 0 or i == total_rows - 1:
        print(f"  Processing row {i+1}/{total_rows} ({((i+1)/total_rows)*100:.1f}%)...")
    
    # Find matching rows in df_beli
    mask = (
        (df_beli["Tanggal"] == penjualan_row["Tanggal"]) &
        (df_beli["Satuan"] == penjualan_row["Satuan"]) &
        df_beli["Nama Barang"].apply(
            lambda x: safe_contains(x, penjualan_row["Nama Barang"])
        )
    )
    matches = df_beli[mask]
    
    if not matches.empty:
        match_counts += 1
        for _, match_row in matches.iterrows():
            merged_row = {
                "Kode #": match_row.get("Kode #", None),
                **{f"{col} Jual": penjualan_row[col] for col in df_penjualan.columns},
                **{f"{col} Beli": match_row[col] for col in df_beli.columns 
                   if col not in ["Tanggal", "Nama Barang", "Satuan", "Kode #"]}
            }
            merged_rows.append(merged_row)
    else:
        no_match_counts += 1
        merged_row = {
            "Kode #": None,
            **{f"{col} Jual": penjualan_row[col] for col in df_penjualan.columns},
            **{f"{col} Beli": None for col in df_beli.columns 
               if col not in ["Tanggal", "Nama Barang", "Satuan", "Kode #"]}
        }
        merged_rows.append(merged_row)

print(f"  Matching results: {match_counts} with matches, {no_match_counts} without matches")
print(f"STEP 4 completed in {time.time() - start_time:.2f} seconds\n")

# =============================================================================
# STEP 5: Create Final DataFrame
# =============================================================================
print("STEP 5: Creating final merged DataFrame...")
start_time = time.time()

merged = pd.DataFrame(merged_rows)
print(f"  Merged DataFrame created with {len(merged)} rows")

# Reorder columns
jual_cols = [col for col in merged.columns 
             if col.endswith("Jual") and col not in ["Tanggal Jual", "Nama Barang Jual", "Satuan Jual"]]
beli_cols = [col for col in merged.columns 
             if col.endswith("Beli") and col != "Kode #"]

columns_order = ["Kode #"] + \
                ["Tanggal Jual", "Nama Barang Jual", "Satuan Jual"] + \
                jual_cols + \
                beli_cols

merged = merged[columns_order]
print("  Columns reordered successfully")
print(f"STEP 5 completed in {time.time() - start_time:.2f} seconds\n")

# =============================================================================
# STEP 6: Export Results
# =============================================================================
print("STEP 6: Exporting to Excel...")
start_time = time.time()

output_path = "./MergedPenjualanPembelianReport2024_DEEPSEEK.xlsx"
try:
    merged.to_excel(output_path, index=False)
    print(f"  File saved successfully to {output_path}")
    print(f"  Final dimensions: {merged.shape[0]} rows x {merged.shape[1]} columns")
except Exception as e:
    print(f"ERROR saving file: {str(e)}")
    raise

print(f"STEP 6 completed in {time.time() - start_time:.2f} seconds\n")

# =============================================================================
# Final Summary
# =============================================================================
total_time = time.time() - program_start_time
print("=== MERGE COMPLETED SUCCESSFULLY ===")
print(f"Total processing time: {total_time:.2f} seconds")
print(f"End Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(f"Summary:")
print(f"- Sales records processed: {total_rows}")
print(f"- Records with matches: {match_counts} ({match_counts/total_rows:.1%})")
print(f"- Records without matches: {no_match_counts} ({no_match_counts/total_rows:.1%})")