import pandas as pd
from datetime import datetime

# ================== CONFIGURATION ==================
file_penjualan = "./penjualan raw mei 2025.xlsx"
file_pembelian = "./pembelian raw 2023 2025.xlsx"
output_file = "MBUPembelianPenjualan2025_Lembur2.xlsx"

# ================== UTILITY FUNCTION FOR VERBOSE LOGGING ==================
def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

log("Starting data cleaning and merging process...")

# ================== STEP 1: CLEAN PENJUALAN ==================
def clean_penjualan(path):
    log("Reading penjualan file...")
    df = pd.read_excel(path)

    log("Setting headers from row 3 and removing first 4 rows...")
    df.columns = df.iloc[3]
    df = df.loc[:, ~df.columns.isna()]
    df = df.drop(range(0,4)).reset_index(drop=True)

    # Clean string columns
    log("Cleaning text fields in penjualan...")
    df['Nama Barang'] = df['Nama Barang'].astype(str).str.strip()
    df['Satuan'] = df['Satuan'].astype(str).str.strip()

    # Convert Tanggal safely
    log("Converting 'Tanggal' to datetime in penjualan...")
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='mixed', errors='coerce')

    # Drop invalid dates
    df = df[df['Tanggal'].notna()]

    log(f"Found {len(df)} valid rows in penjualan.")
    return df

# ================== STEP 2: CLEAN PEMBELIAN ==================
def clean_pembelian(path):
    log("Reading pembelian file...")
    df = pd.read_excel(path)

    log("Dropping empty columns...")
    df = df.dropna(axis="columns", how="all")

    log("Setting headers from row 3 and removing first 4 rows...")
    df.columns = df.iloc[3]
    df = df.drop(range(0,4)).reset_index(drop=True)
    df = df.loc[:, ~df.columns.isna()]

    log("Cleaning 'Kode #' values...")
    df['Kode #'] = df['Kode #'].astype(str).apply(lambda x: x.lstrip('0') if x else '0')

    # Clean string columns
    log("Cleaning text fields in pembelian...")
    df['Nama Barang'] = df['Nama Barang'].astype(str).str.strip()
    df['Satuan'] = df['Satuan'].astype(str).str.strip()

    # Convert Tanggal safely
    log("Converting 'Tanggal' to datetime in pembelian...")
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='mixed', errors='coerce')

    # Drop invalid dates
    df = df[df['Tanggal'].notna()]

    log(f"Found {len(df)} valid rows in pembelian.")
    return df

# ================== CLEAN DATA ==================
log("Cleaning penjualan data...")
df_jual = clean_penjualan(file_penjualan)

log("Cleaning pembelian data...")
df_beli = clean_pembelian(file_pembelian)

# ================== MERGE LOGIC ==================
log("Starting merge process...")
merged_rows = []
unmatched_rows = []

# Show available columns for debugging
log("Available columns in penjualan: " + ", ".join(df_jual.columns.astype(str)))
log("Available columns in pembelian: " + ", ".join(df_beli.columns.astype(str)))

for idx, row_jual in df_jual.iterrows():
    jual_nama = row_jual['Nama Barang']
    jual_satuan = row_jual['Satuan']
    jual_tanggal = row_jual['Tanggal']

    # Filter by Satuan and partial Nama Barang match
    matches = df_beli[
        (df_beli['Satuan'] == jual_satuan) &
        (df_beli['Nama Barang'].str.contains(jual_nama, case=False, na=False, regex=False))
    ]

    if not matches.empty:
        # Filter purchases before the sales date
        matches = matches[matches['Tanggal'] <= jual_tanggal]

        if not matches.empty:
            best_match = matches.loc[matches['Tanggal'].idxmax()]

            merged_row = {
                'Nama Barang': jual_nama,
                'Satuan': jual_satuan,
                'Tanggal_Jual': jual_tanggal.strftime('%d %b %Y'),
                'Kode #_Jual': row_jual.get('Kode #', ''),
            }

            # Add all relevant penjualan fields with _Jual suffix
            for col in df_jual.columns:
                if col not in ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']:
                    merged_row[f"{col}_Jual"] = row_jual.get(col, '')

            # Add all relevant pembelian fields with _Beli suffix
            if not best_match.empty:
                merged_row['Kode #_Beli'] = best_match.get('Kode #', '')
                merged_row['Tanggal_Beli'] = best_match.get('Tanggal').strftime('%d %b %Y') if pd.notna(best_match.get('Tanggal')) else ''

                for col in df_beli.columns:
                    if col not in ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']:
                        merged_row[f"{col}_Beli"] = best_match.get(col, '')
            else:
                merged_row['Kode #_Beli'] = ''
                merged_row['Tanggal_Beli'] = ''

            merged_rows.append(merged_row)
        else:
            unmatched_row = {
                'Nama Barang': jual_nama,
                'Satuan': jual_satuan,
                'Tanggal_Jual': jual_tanggal.strftime('%d %b %Y'),
                'Reason': 'No purchase before sale date'
            }
            unmatched_rows.append(unmatched_row)
    else:
        unmatched_row = {
            'Nama Barang': jual_nama,
            'Satuan': jual_satuan,
            'Tanggal_Jual': jual_tanggal.strftime('%d %b %Y'),
            'Reason': 'No matching item found'
        }
        unmatched_rows.append(unmatched_row)

# Convert to DataFrames
log("Building merged DataFrame...")
df_merged = pd.DataFrame(merged_rows)
df_unmatched = pd.DataFrame(unmatched_rows)


# ---- NEW: Add DUAL HPP Calculation ----
log("Starting dual HPP calculation...")

# List to hold missing column names (if any)
missing_columns = []

# Check if required columns exist for both calculations
has_hpp_jual = all(col in df_merged.columns for col in ['Penjualan_Jual', 'Laba_Jual', 'Kuantitas_Jual'])
has_hpp_beli = all(col in df_merged.columns for col in ['Penjualan_Jual', 'Laba_Jual', 'Kuantitas_Beli'])

if has_hpp_jual:
    # Calculate HPP_Jual = (Penjualan_Jual - Laba_Jual) / Kuantitas_Jual
    df_merged['Kuantitas_Jual'] = df_merged['Kuantitas_Jual'].replace(0, pd.NA)
    df_merged['HPP_Jual'] = (df_merged['Penjualan_Jual'] - df_merged['Laba_Jual']) / df_merged['Kuantitas_Jual']
    df_merged['HPP_Jual'] = df_merged['HPP_Jual'].round(2).fillna(0)
else:
    missing = [col for col in ['Penjualan_Jual', 'Laba_Jual', 'Kuantitas_Jual'] if col not in df_merged.columns]
    missing_columns.extend(missing)
    log(f"⚠️ Missing columns for HPP_Jual: {', '.join(missing)}")

if has_hpp_beli:
    # Calculate HPP_Beli = (Penjualan_Jual - Laba_Jual) / Kuantitas_Beli
    df_merged['Kuantitas_Beli'] = df_merged['Kuantitas_Beli'].replace(0, pd.NA)
    df_merged['HPP_Beli'] = (df_merged['Penjualan_Jual'] - df_merged['Laba_Jual']) / df_merged['Kuantitas_Beli']
    df_merged['HPP_Beli'] = df_merged['HPP_Beli'].round(2).fillna(0)
else:
    missing = [col for col in ['Penjualan_Jual', 'Laba_Jual', 'Kuantitas_Beli'] if col not in df_merged.columns]
    missing_columns.extend(missing)
    log(f"⚠️ Missing columns for HPP_Beli: {', '.join(missing)}")

if not missing_columns:
    log("✅ Both HPP columns calculated successfully.")
elif len(set(missing_columns)) == len(missing_columns):
    log(f"❌ Some columns are still missing: {', '.join(set(missing_columns))}")

# Sort merged data
log("Sorting merged data...")
df_merged = df_merged.sort_values(by=['Tanggal_Jual', 'Nama Barang'])
df_unmatched = df_unmatched.sort_values(by=['Tanggal_Jual', 'Nama Barang'])

# ================== EXPORT TO EXCEL WITH MULTIPLE SHEETS ==================
log(f"Exporting results to {output_file}...")
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df_merged.to_excel(writer, sheet_name='Merged', index=False)
    df_unmatched.to_excel(writer, sheet_name='Unmatched', index=False)

log("Process completed successfully!")
log(f"✔️ Merged rows: {len(df_merged)}")
log(f"❌ Unmatched rows: {len(df_unmatched)}")
log(f"📄 Output saved to: {output_file}")