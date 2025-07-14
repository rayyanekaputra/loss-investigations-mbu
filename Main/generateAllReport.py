import os
import pandas as pd
from openpyxl import Workbook

# Configuration
input_folder = "./BAEKMI"
output_file = "Laporan_Laba_Bulanan.xlsx"

# Get list of input files
input_files = [f for f in os.listdir(input_folder) if f.startswith('_merge_with_purchasing')]

# Create a Pandas Excel writer
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    
    # Summary dataframe to collect monthly totals
    summary_data = []

    for file in input_files:
        file_path = os.path.join(input_folder, file)

        # Read the Excel file
        df = pd.read_excel(file_path)

        # Clean column names (if needed)
        df.columns = df.columns.str.strip()

        # --- Calculate HPP per item ---
        df['HPP'] = df['@Harga Beli'] * df['Kuantitas Beli']

        # --- Calculate difference between HPP and Harga Beli ---
        df['Selisih HPP vs Harga Beli'] = df['HPP'] - df['@Harga Beli']

        # --- Recalculate Laba (Profit) = Penjualan - HPP - Diskon ---
        df['Hitung Ulang Laba'] = df['Penjualan'] - df['HPP'] - df['Diskon']

        # --- Optional: Drop unnecessary columns to make it cleaner ---
        cols_to_show = [
            'Nama Barang', 'Tanggal', 'Kuantitas', 'Satuan',
            '@Harga', 'Total Harga', 'Penjualan', 'Diskon',
            'HPP', 'Selisih HPP vs Harga Beli', 'Hitung Ulang Laba'
        ]
        df_summary = df[cols_to_show]

        # --- Monthly summary ---
        total_penjualan = df['Penjualan'].sum()
        total_hpp = df['HPP'].sum()
        total_laba_baru = df['Hitung Ulang Laba'].sum()
        laba_selisih = df['Hitung Ulang Laba'].sum() - df['Laba'].sum()

        # Extract month from filename
        month_name = file.split("_")[-1].replace(".xlsx", "").capitalize()

        summary_data.append({
            'Bulan': month_name,
            'Total Penjualan': total_penjualan,
            'Total HPP': total_hpp,
            'Total Laba (Baru)': total_laba_baru,
            'Selisih Laba Lama-Baru': laba_selisih
        })

        # Write detailed sheet for this month
        sheet_name = month_name[:31]  # Excel sheet name limit
        df_summary.to_excel(writer, sheet_name=sheet_name, index=False)

    # --- Save the summary sheet ---
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name="Ringkasan Bulanan", index=False)

print(f"✅ Laporan berhasil dibuat: {output_file}")