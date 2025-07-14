import pandas as pd
import os
from datetime import datetime

def process_sales_purchasing(sales_file, purchasing_file):
    """Process and merge sales data with purchasing data"""
    try:
        # Read files
        df_sales = pd.read_excel(sales_file, sheet_name='Data Penjualan')
        df_purchasing = pd.read_excel(purchasing_file)
        
        # Print column names
        print(f"Sales columns in {os.path.basename(sales_file)}: {df_sales.columns.tolist()}")
        print(f"Purchasing columns in {os.path.basename(purchasing_file)}: {df_purchasing.columns.tolist()}")
        
        # Verify 'Nama Barang' column
        if 'Nama Barang' not in df_sales.columns:
            raise ValueError(f"'Nama Barang' column missing in sales file: {sales_file}")
        if 'Nama Barang' not in df_purchasing.columns:
            raise ValueError(f"'Nama Barang' column missing in purchasing file: {purchasing_file}")
        
        # Validate data
        for df, name in [(df_sales, 'sales'), (df_purchasing, 'purchasing')]:
            if df['Nama Barang'].isna().any():
                print(f"Warning: {df['Nama Barang'].isna().sum()} missing values in 'Nama Barang' in {name} file")
                df = df.dropna(subset=['Nama Barang'])
            if df['Nama Barang'].str.strip().eq('').any():
                print(f"Warning: {df['Nama Barang'].str.strip().eq('').sum()} empty strings in 'Nama Barang' in {name} file")
                df = df[df['Nama Barang'].str.strip() != '']
            if name == 'purchasing' and df['Nama Barang'].duplicated().any():
                duplicates = df[df['Nama Barang'].duplicated(keep=False)]
                print(f"Warning: {df['Nama Barang'].duplicated().sum()} duplicate values in 'Nama Barang' in purchasing file")
                print("Duplicate 'Nama Barang' values:", duplicates['Nama Barang'].tolist())
                df = df.drop_duplicates(subset=['Nama Barang'], keep='first')
        
        # Update DataFrames after cleaning
        df_sales = df_sales[df_sales['Nama Barang'].str.strip() != '']
        df_purchasing = df_purchasing[df_purchasing['Nama Barang'].str.strip() != '']
        
        # Print sample data and types
        print(f"Sample 'Nama Barang' from sales (first 5): {df_sales['Nama Barang'].head().tolist()}")
        print(f"Sample 'Nama Barang' from purchasing (first 5): {df_purchasing['Nama Barang'].head().tolist()}")
        print("Sales 'Nama Barang' types:", df_sales['Nama Barang'].apply(type).value_counts().to_dict())
        print("Purchasing 'Nama Barang' types:", df_purchasing['Nama Barang'].apply(type).value_counts().to_dict())
        
        # Create purchasing mapping dictionary
        purchasing_map = {}
        for _, row in df_purchasing.iterrows():
            name = row['Nama Barang']
            try:
                if pd.notna(name) and name.strip():
                    cleaned_name = str(name).strip().lower()
                    if cleaned_name in purchasing_map:
                        print(f"Skipping duplicate 'Nama Barang' after cleaning: {cleaned_name}")
                        continue
                    purchasing_map[cleaned_name] = row.to_dict()
                else:
                    print(f"Skipping invalid 'Nama Barang' value in purchasing file: {name!r}")
            except Exception as e:
                print(f"Error processing 'Nama Barang' value '{name!r}' in purchasing file: {str(e)}")
                raise ValueError(f"Failed to process 'Nama Barang' value '{name!r}': {str(e)}")
        
        # Match purchasing data to sales data
        for col in df_purchasing.columns:
            try:
                df_sales[f"{col} Beli"] = df_sales['Nama Barang'].apply(
                    lambda x: next(
                        (purchasing_map[product][col] for product in purchasing_map 
                         if pd.notna(x) and x.strip() and str(x).strip().lower() in product),
                        None
                    )
                )
            except Exception as e:
                print(f"Error matching column '{col} Beli' for 'Nama Barang' values")
                raise ValueError(f"Error matching column '{col} Beli': {str(e)}")
        
        # Extract month and number
        base_name = os.path.basename(sales_file)
        parts = base_name.split('_')
        if len(parts) < 3:
            raise ValueError(f"Invalid sales filename format: {sales_file}")
        number = parts[0]
        month = parts[-1].replace('.xlsx', '')
        
        # Prepare output filename
        output_file = f"./BAEKMI/{number}_merge_with_purchasing_{month}.xlsx"
        
        # Export to Excel
        with pd.ExcelWriter(output_file) as writer:
            df_sales.to_excel(writer, sheet_name='Data Penjualan', index=False)
            matched = df_sales['Nama Pemasok Faktur Pembelian Beli'].notna().sum()
            summary = pd.DataFrame({
                'Metric': ['Total Rows', 'Matched Purchasing Records', 'Match Rate'],
                'Value': [len(df_sales), matched, f"{matched/len(df_sales):.1%}"]
            })
            summary.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"Successfully processed and saved: {output_file}")
        return output_file
    
    except Exception as e:
        raise Exception(f"Error processing {sales_file} with {purchasing_file}: {str(e)}")

def find_matching_files(directory):
    """Find matching sales and purchasing files in directory"""
    files = os.listdir(directory)
    processed = set()
    
    for f in files:
        if '_merge_' in f.lower() and f.endswith('.xlsx') and '_with_purchasing_' not in f.lower():
            parts = f.split('_')
            if len(parts) < 3:
                continue
            num = parts[0]
            month = parts[-1].replace('.xlsx', '')
            
            purchasing_pattern = f"{num} pembelian terbaru per barang hingga {month}"
            purchasing_match = [x for x in files if purchasing_pattern.lower() in x.lower()]
            
            if purchasing_match:
                sales_file = os.path.join(directory, f)
                purchasing_file = os.path.join(directory, purchasing_match[0])
                yield (sales_file, purchasing_file)
                processed.add(num)
    
    all_nums = {f.split('_')[0] for f in files if '_merge_' in f.lower() and '_with_purchasing_' not in f.lower() and f.endswith('.xlsx')}
    unmatched = all_nums - processed
    if unmatched:
        print(f"\nWarning: No purchasing file found for sales files with numbers: {', '.join(unmatched)}")

# Main execution
if __name__ == "__main__":
    directory = "./BAEKMI"
    
    print("Starting automated sales-purchasing matching...")
    print(f"Processing files in: {directory}\n")
    
    results = []
    for sales_file, purchasing_file in find_matching_files(directory):
        print(f"Processing:\n- Sales: {os.path.basename(sales_file)}\n- Purchasing: {os.path.basename(purchasing_file)}")
        try:
            output = process_sales_purchasing(sales_file, purchasing_file)
            results.append(output)
        except Exception as e:
            print(str(e))
    
    print("\nProcessing complete. Files created:")
    for r in results:
        print(f"- {r}")