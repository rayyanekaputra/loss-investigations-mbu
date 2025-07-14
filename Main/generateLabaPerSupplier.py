import pandas as pd
import os
from glob import glob

def analyze_supplier_profits(directory="./BAEKMI"):
    merged_files = glob(os.path.join(directory, "*_merge_*.xlsx"))
    
    if not merged_files:
        print("No merged files found in directory:", directory)
        return
    
    all_results = {}
    combined_data = []  # For summary
    
    for file_path in merged_files:
        try:
            file_name = os.path.basename(file_path)
            number, month = file_name.split('_merge_')
            month = month.replace('.xlsx', '')
            
            print(f"\nProcessing {file_name}...")
            df = pd.read_excel(file_path)
            
            # Check required columns
            if 'Pemasok' not in df.columns or 'Laba' not in df.columns:
                print(f"Skipping {file_name} - missing required columns")
                continue
            
            # Convert 'Laba' to numeric
            df['Laba'] = pd.to_numeric(df['Laba'], errors='coerce')
            
            # Drop rows where 'Laba' is missing
            df = df.dropna(subset=['Laba'])
            
            # Fill missing 'Pemasok'
            df['Pemasok'] = df['Pemasok'].fillna('(Tidak Diketahui)')
            
            # Group and sum
            supplier_profits = df.groupby('Pemasok', as_index=False)['Laba'].sum()
            supplier_profits = supplier_profits.sort_values(by='Laba')
            
            print(supplier_profits.head(5))
            losses_count = (supplier_profits['Laba'] < 0).sum()
            print(f"→ Suppliers with losses: {losses_count}")
            
            # Store per-file results
            all_results[f"{number}_{month}"] = supplier_profits
            
            # Add to combined data for summary
            combined_data.append(supplier_profits)
        
        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")
    
    if not all_results:
        print("No valid data found in any files")
        return
    
    # Create combined summary
    summary_df = pd.concat(combined_data)
    summary_df = summary_df.groupby('Pemasok', as_index=False)['Laba'].sum()
    summary_df = summary_df.sort_values(by='Laba')
    
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(directory, f"supplier_total_profit_{timestamp}.xlsx")
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write each file’s sheet
        for sheet_name, df in all_results.items():
            sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(i, i, max_len + 2)
        
        # Write summary sheet
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        worksheet = writer.sheets['Summary']
        for i, col in enumerate(summary_df.columns):
            max_len = max(summary_df[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(i, i, max_len + 2)
    
    print(f"\n✅ Analysis complete. Summary and details saved to: {output_file}")

if __name__ == "__main__":
    analyze_supplier_profits()
