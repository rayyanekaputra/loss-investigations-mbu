import pandas as pd
from datetime import datetime

print("🔄 STEP 1: Loading Excel files...")
df_penjualan = pd.read_excel("./BAEKMI/Penjualan2024.xlsx")
df_beli = pd.read_excel("./PembelianBuDian2024.xlsx")

print("✅ Loaded:")
print(f"   - df_penjualan: {df_penjualan.shape[0]} rows")
print(f"   - df_beli     : {df_beli.shape[0]} rows\n")

print("🧹 STEP 2: Cleaning and formatting key columns...")

def clean_text(s):
    return str(s).strip()

# Clean columns
print("   - Cleaning df_penjualan columns: 'Nama Barang', 'Satuan', 'Tanggal'")
df_penjualan['Nama Barang'] = df_penjualan['Nama Barang'].apply(clean_text)
df_penjualan['Satuan'] = df_penjualan['Satuan'].apply(clean_text)
df_penjualan['Tanggal'] = pd.to_datetime(df_penjualan['Tanggal'], errors='coerce').dt.strftime('%d %b %Y')

print("   - Cleaning df_beli columns: 'Nama Barang', 'Satuan', 'Tanggal'")
df_beli['Nama Barang'] = df_beli['Nama Barang'].apply(clean_text)
df_beli['Satuan'] = df_beli['Satuan'].apply(clean_text)
df_beli['Tanggal'] = pd.to_datetime(df_beli['Tanggal'], errors='coerce').dt.strftime('%d %b %Y')

print("✅ Columns cleaned and dates formatted to dd Mon yyyy format.\n")

print("🔁 STEP 3: Starting merge logic with partial string matching...")

merged_rows = []

for idx, jual_row in df_penjualan.iterrows():
    if idx % 50 == 0 or idx == len(df_penjualan) - 1:
        print(f"   > Processing row {idx + 1} of {len(df_penjualan)}")

    jual_nama = jual_row['Nama Barang']
    jual_tanggal = jual_row['Tanggal']
    jual_satuan = jual_row['Satuan']

    matching_beli = df_beli[
        (df_beli['Tanggal'] == jual_tanggal) &
        (df_beli['Satuan'] == jual_satuan) &
        (df_beli['Nama Barang'].str.contains(jual_nama, na=False, regex=False))
    ]

    if matching_beli.empty:
        print(f"     ⚠️  No match for: '{jual_nama}' on {jual_tanggal} [{jual_satuan}]")
        row_data = {'Kode #': None}
        row_data.update(jual_row.drop(['Tanggal', 'Nama Barang', 'Satuan']).add_suffix(' Jual'))
        row_data.update({
            'Tanggal': jual_tanggal,
            'Nama Barang': jual_nama,
            'Satuan': jual_satuan
        })
        merged_rows.append(row_data)
    else:
        print(f"     ✅ {len(matching_beli)} match(es) found for: '{jual_nama}' on {jual_tanggal} [{jual_satuan}]")
        for _, beli_row in matching_beli.iterrows():
            row_data = {'Kode #': beli_row['Kode #']}
            row_data.update(jual_row.drop(['Tanggal', 'Nama Barang', 'Satuan']).add_suffix(' Jual'))
            row_data.update(beli_row.drop(['Tanggal', 'Nama Barang', 'Satuan', 'Kode #']).add_suffix(' Beli'))
            row_data.update({
                'Tanggal': jual_tanggal,
                'Nama Barang': jual_nama,
                'Satuan': jual_satuan
            })
            merged_rows.append(row_data)

print("\n📦 STEP 4: Building final merged DataFrame...")
df_merged = pd.DataFrame(merged_rows)
print(f"✅ Merged DataFrame created with {df_merged.shape[0]} rows.\n")

print("🧾 STEP 5: Reordering columns...")
first_cols = ['Kode #', 'Tanggal', 'Nama Barang', 'Satuan']
other_cols = [col for col in df_merged.columns if col not in first_cols]
df_merged = df_merged[first_cols + other_cols]
print("✅ Columns reordered with 'Kode #' as the first column.\n")

print("💾 STEP 6: Exporting to Excel file...")
output_path = "./MergedPenjualanPembelianReport2024_CHATGPT.xlsx"
df_merged.to_excel(output_path, index=False)
print(f"✅ Done! File saved to: {output_path}")
