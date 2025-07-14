import pandas as pd
import os
from glob import glob

def analyze_supplier_profits(directory="./BAEKMI"):
    # Find all merged files in the directory
    merged_files = glob(os.path.join(directory, "*_merge_*.xlsx"))
    
    if not merged_files:
        print("No merged files found in directory:", directory)
        return
    
    all_results = {}

    for file_path in merged_files:
        try:
            # Extract month and number from filename
            file_name = os.path.basename(file_path)
            number, month = file_name.split('_merge_')
            month = month.replace('.xlsx', '')

            print(f"\n🔍 Processing {file_name}...")

            # Read the Excel file
            df = pd.read_excel(file_path)

            # Ensure required columns exist
            if 'Pemasok' not in df.columns or 'Laba' not in df.columns:
                print(f"⚠️ Skipping {file_name} - missing 'Pemasok' or 'Laba'")
                continue

            # Convert 'Laba' to numeric, handle errors
            df['Laba'] = pd.to_numeric(df['Laba'], errors='coerce')
            df = df.dropna(subset=['Laba'])

            # Handle missing suppliers
            df['Pemasok'] = df['Pemasok'].fillna('(Tidak Diketahui)')

            # Get the lowest-profit item per supplier
            lowest_profit_items = df.loc[df.groupby('Pemasok')['Laba'].idxmin()]

            # Sort by profit (ascending)
            lowest_profit_items = lowest_profit_items.sort_values('Laba')

            # Select columns to output
            output_cols = [
                'Pemasok', 'Nama Barang', 'Laba', 'Total Harga', 'Kuantitas',
                'Satuan', '@Harga', 'Nama Kategori Barang Barang & Jasa'
            ]
            output_cols = [col for col in output_cols if col in lowest_profit_items.columns]

            result_df = lowest_profit_items[output_cols]

            # Count losses
            num_losses = (result_df['Laba'] < 0).sum()
            print(f"✅ Found {len(result_df)} suppliers, {num_losses} with losses (Laba < 0)")

            # Store results
            all_results[f"{number}_{month}"] = result_df

        except Exception as e:
            print(f"❌ Error processing {file_name}: {str(e)}")

    if not all_results:
        print("🚫 No valid data found in any files.")
        return

    # Generate output file name
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(directory, f"supplier_lowest_profit_{timestamp}.xlsx")

    # Write to Excel with auto-adjusted columns
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for sheet_name, df in all_results.items():
            sheet_name = sheet_name[:31]  # Excel sheet name limit
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Auto-adjust column width
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(i, i, max_len + 2)

    print(f"\n📁 Analysis complete. Results saved to: {output_file}")

if __name__ == "__main__":
    analyze_supplier_profits()
