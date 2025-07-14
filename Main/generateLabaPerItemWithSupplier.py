import pandas as pd
import os
from glob import glob

def export_item_profit_and_losses(directory="./BAEKMI"):
    merged_files = glob(os.path.join(directory, "*_merge_*.xlsx"))
    
    if not merged_files:
        print("No merged files found in directory:", directory)
        return
    
    all_results = {}
    combined_data = []

    for file_path in merged_files:
        try:
            file_name = os.path.basename(file_path)
            number, month = file_name.split('_merge_')
            month = month.replace('.xlsx', '')

            print(f"\nProcessing {file_name}...")
            df = pd.read_excel(file_path)

            if 'Nama Barang' not in df.columns or 'Laba' not in df.columns:
                print(f"Skipping {file_name} - missing 'Nama Barang' or 'Laba'")
                continue

            # Convert Laba to numeric
            df['Laba'] = pd.to_numeric(df['Laba'], errors='coerce')
            df = df.dropna(subset=['Laba'])

            # Fill missing 'Pemasok'
            df['Pemasok'] = df['Pemasok'].fillna('(Tidak Diketahui)')

            # Select relevant columns
            output_cols = ['Nama Barang', 'Pemasok', 'Laba']
            optional_cols = ['Total Harga', 'Kuantitas', 'Satuan', '@Harga', 'Nama Kategori Barang Barang & Jasa']
            for col in optional_cols:
                if col in df.columns:
                    output_cols.append(col)

            result_df = df[output_cols].copy()

            # Save per-file result
            all_results[f"{number}_{month}"] = result_df

            # Add to combined summary
            combined_data.append(result_df[['Nama Barang', 'Pemasok', 'Laba']])
        
        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")

    if not all_results:
        print("No valid data found in any files")
        return

    # Build summary: total profit per item per supplier
    summary_df = pd.concat(combined_data)
    summary_df = (
        summary_df.groupby(['Nama Barang', 'Pemasok'], as_index=False)
        .agg({'Laba': 'sum'})
        .sort_values(by='Laba')
    )

    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(directory, f"item_profit_loss_{timestamp}.xlsx")

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write each file's sheet
        for sheet_name, df in all_results.items():
            sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(i, i, max_len + 2)

        # Write summary
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        worksheet = writer.sheets['Summary']
        for i, col in enumerate(summary_df.columns):
            max_len = max(summary_df[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(i, i, max_len + 2)

    print(f"\n✅ Export complete. Results saved to: {output_file}")

if __name__ == "__main__":
    export_item_profit_and_losses()
