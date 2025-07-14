import pandas as pd
import os
from datetime import datetime

def process_monthly_sales(supplier_file, sales_file):
    """Process and merge supplier data with sales data"""
    # Read files
    df_supplier = pd.read_excel(supplier_file)
    df_sales = pd.read_excel(sales_file)
    
    # Create supplier mapping dictionary (lowercase for case-insensitive matching)
    supplier_map = {
        str(name).strip().lower(): supplier 
        for name, supplier in zip(df_supplier['Nama Barang'], df_supplier['Pemasok'])
    }
    
    # Match suppliers to sales data
    df_sales['Pemasok'] = df_sales['Nama Barang'].apply(
        lambda x: next(
            (supplier for product, supplier in supplier_map.items() 
             if str(x).strip().lower() in product),
            None
        )
    )
    
    # Extract month and number from filename
    base_name = os.path.basename(sales_file)
    parts = base_name.split(' ')
    number = parts[0]
    month = parts[-1].replace('.xlsx', '')
    
    # Prepare output filename
    output_file = f"./BAEKMI/{number}_merge_{month}.xlsx"
    
    # Export to Excel
    with pd.ExcelWriter(output_file) as writer:
        df_sales.to_excel(writer, sheet_name='Data Penjualan', index=False)
        
        # Add summary sheet
        matched = df_sales['Pemasok'].notna().sum()
        summary = pd.DataFrame({
            'Metric': ['Total Rows', 'Matched Suppliers', 'Match Rate'],
            'Value': [len(df_sales), matched, f"{matched/len(df_sales):.1%}"]
        })
        summary.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"Successfully processed and saved: {output_file}")
    return output_file

def find_matching_files(directory):
    """Find matching supplier and sales files in directory"""
    files = os.listdir(directory)
    processed = set()
    
    for f in files:
        if 'penjualan' in f.lower() and f.endswith('.xlsx'):
            # Extract the number prefix (e.g., "01" from "01 penjualan januari.xlsx")
            num = f.split(' ')[0]
            
            # Find matching supplier file
            supplier_pattern = f"{num} supplier"
            supplier_match = [x for x in files if supplier_pattern.lower() in x.lower()]
            
            if supplier_match:
                supplier_file = os.path.join(directory, supplier_match[0])
                sales_file = os.path.join(directory, f)
                yield (supplier_file, sales_file)
                processed.add(num)
    
    # Print any unmatched files
    all_nums = {f.split(' ')[0] for f in files if f.endswith('.xlsx')}
    unmatched = all_nums - processed
    if unmatched:
        print(f"\nWarning: No supplier file found for files with numbers: {', '.join(unmatched)}")

# Main execution
if __name__ == "__main__":
    directory = "./BAEKMI"
    
    print("Starting automated sales-supplier matching...")
    print(f"Processing files in: {directory}\n")
    
    results = []
    for supplier_file, sales_file in find_matching_files(directory):
        print(f"Processing:\n- Sales: {os.path.basename(sales_file)}\n- Supplier: {os.path.basename(supplier_file)}")
        try:
            output = process_monthly_sales(supplier_file, sales_file)
            results.append(output)
        except Exception as e:
            print(f"Error processing {sales_file}: {str(e)}")
    
    print("\nProcessing complete. Files created:")
    for r in results:
        print(f"- {r}")