import pandas as pd

# Step 1: Read the files
print("Reading sales and purchase files...")
df_penjualan = pd.read_excel("./BAEKMI/Penjualan2024.xlsx")
df_beli = pd.read_excel("./PembelianBuDian2024.xlsx")

# Step 2: Trim whitespaces and enforce consistent format
print("Cleaning whitespace and formatting columns...")
df_penjualan['Nama Barang'] = df_penjualan['Nama Barang'].astype(str).str.strip()
df_beli['Nama Barang'] = df_beli['Nama Barang'].astype(str).str.strip()
df_penjualan['Satuan'] = df_penjualan['Satuan'].astype(str).str.strip()
df_beli['Satuan'] = df_beli['Satuan'].astype(str).str.strip()

# Convert 'Tanggal' to datetime
print("Converting Tanggal to datetime format...")
df_penjualan['Tanggal'] = pd.to_datetime(df_penjualan['Tanggal'])
df_beli['Tanggal'] = pd.to_datetime(df_beli['Tanggal'])

# Step 3: Prepare an empty list to collect merged rows
merged_rows = []

# Step 4: Iterate over penjualan rows to match
print("Matching sales with purchases using partial Nama Barang match...")
for idx, row_jual in df_penjualan.iterrows():
    jual_nama = row_jual['Nama Barang']
    jual_satuan = row_jual['Satuan']

    # Filter df_beli by matching 'Satuan' and partial 'Nama Barang'
    match_beli = df_beli[
        (df_beli['Satuan'] == jual_satuan) &
        (df_beli['Nama Barang'].str.contains(jual_nama, na=False, regex=False))
    ]

    if not match_beli.empty:
        # Pick the latest purchase date (can be multiple rows if same date)
        latest_date = match_beli['Tanggal'].max()
        match_beli = match_beli[match_beli['Tanggal'] == latest_date]

        for _, row_beli in match_beli.iterrows():
            merged_row = {}

            # Prioritize 'Kode #_Jual' then 'Kode #_Beli'
            merged_row['Kode #_Jual'] = row_jual.get('Kode #')
            merged_row['Kode #_Beli'] = row_beli.get('Kode #')

            # Add Tanggal, Nama Barang, Satuan
            merged_row['Tanggal'] = row_jual['Tanggal']
            merged_row['Nama Barang'] = jual_nama
            merged_row['Satuan'] = jual_satuan

            # Add rest of penjualan columns with suffix 'Jual'
            for col in df_penjualan.columns:
                if col not in ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']:
                    merged_row[f"{col}_Jual"] = row_jual[col]

            # Add rest of pembelian columns with suffix 'Beli'
            for col in df_beli.columns:
                if col not in ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']:
                    merged_row[f"{col}_Beli"] = row_beli[col]

            merged_rows.append(merged_row)
    else:
        # No match: merge with empty beli data
        merged_row = {}

        merged_row['Kode #_Jual'] = row_jual.get('Kode #')
        merged_row['Kode #_Beli'] = None
        merged_row['Tanggal'] = row_jual['Tanggal']
        merged_row['Nama Barang'] = jual_nama
        merged_row['Satuan'] = jual_satuan

        for col in df_penjualan.columns:
            if col not in ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']:
                merged_row[f"{col}_Jual"] = row_jual[col]
        for col in df_beli.columns:
            if col not in ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']:
                merged_row[f"{col}_Beli"] = None

        merged_rows.append(merged_row)

# Step 5: Create merged DataFrame
print("Creating final merged DataFrame...")
df_merged = pd.DataFrame(merged_rows)

# Step 6: Sort by Tanggal then Nama Barang
print("Sorting merged data by Tanggal and Nama Barang...")
df_merged = df_merged.sort_values(by=['Tanggal', 'Nama Barang'], ascending=[True, True])

# Step 7: Format Tanggal as 'DD Mon YYYY'
df_merged['Tanggal'] = df_merged['Tanggal'].dt.strftime('%d %b %Y')

# Step 8: Export to Excel
output_path = "./MergedDianRayyan2024.xlsx"
print(f"Exporting merged data to {output_path}...")
df_merged.to_excel(output_path, index=False)

print("Done! Merged report saved.")
