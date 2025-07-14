import pandas as pd
import os
from pathlib import Path

# Define month mappings
MONTHS = {
    '01': 'Januari',
    '02': 'Februari',
    '03': 'Maret',
    '04': 'April',
    '05': 'Mei',
    '06': 'Juni',
    '07': 'Juli',
    '08': 'Agustus',
    '09': 'September',
    '10': 'Oktober',
    '11': 'November',
    '12': 'Desember'
}

def check_duplicates(df, month_name):
    """Check for duplicate Nama Barang in the dataframe"""
    duplicates = df[df.duplicated(subset=['Nama Barang'], keep=False)]
    if not duplicates.empty:
        print(f"Warning: Found duplicate Nama Barang in {month_name}:")
        print(duplicates[['Kode #', 'Nama Barang', 'Tanggal']])
        print(f"Total duplicate entries: {len(duplicates)}\n")
    else:
        print(f"No duplicate Nama Barang found in {month_name}\n")

def process_month(month_num, month_name):
    # Create paths
    input_dir = Path(f"./{month_num} {month_name}")
    output_dir = Path("./BAEKMI")  # Single output directory
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    input_file = input_dir / f"Pembelian per Barang hingga {month_name}.xlsx"
    output_file = output_dir / f"{month_num} pembelian terbaru dan unit terkecil per barang hingga {month_name}.xlsx"
    
    print(f"Processing {input_file}...")
    
    try:
        # Read the excel file
        df = pd.read_excel(input_file)
        
        # Clean columns
        df.dropna(axis="columns", how="all", inplace=True)  # Drop empty columns
        df.rename(columns=df.iloc[3], inplace=True)
        df.drop(range(0,4), axis=0, inplace=True)  # Drop header rows
        df = df.loc[:, ~df.columns.isna()]
        df.reset_index(drop=True, inplace=True)
        
        # Clean Kode Barang
        df['Kode #'] = df['Kode #'].astype(str)
        df['Kode #'] = df['Kode #'].apply(
            lambda x: x.lstrip('0') if x.lstrip('0') else '0'
        )
        
        # Remove duplicate header rows
        duplicate = df['Tanggal'] == 'Tanggal'
        df = df[~duplicate]
        
        # Convert date column
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%Y-%m-%d %H:%M:%S')
        
        # Step 1: Define UOM hierarchy (lower number = smaller unit)
        uom_priority = {
            'PCS': 1, 'Pcs': 1, 'pcs': 1,
            'ML': 2, 'Ml': 2, 'ml': 2,
            'GR': 3, 'Gr': 3, 'gr': 3,
            'LTR': 4, 'Ltr': 4, 'ltr': 4,
            'PAI': 5, 'Pai': 5,
            'PCK': 6, 'Pck': 6, 'pck': 6,
            'CAN': 7, 'Can': 7,
            'BTL': 8, 'Btl': 8, 'btl': 8,
            'JAR': 9, 'Jar': 9,
            'JRG': 10, 'Jrg': 10,
            'BOX': 11, 'Box': 11, 'box': 11,
            'CTN': 12, 'Ctn': 12, 'ctn': 12,
            'GAL': 13, 'Gal': 13,
            'KG': 14, 'Kg': 14, 'kg': 14,
        }

        # Step 2: Normalize and assign priority
        df['Satuan'] = df['Satuan'].astype(str).str.strip()
        df['UOM_Priority'] = df['Satuan'].map(lambda x: uom_priority.get(x.upper(), 999))  # default 999 for unknowns

        # Step 3: Keep only rows with smallest UOM per 'Nama Barang'
        min_uom_idx = df.groupby('Nama Barang')['UOM_Priority'].idxmin()
        df = df.loc[min_uom_idx].reset_index(drop=True)

        # Step 4: From those, get the latest entry per item
        df_sorted = df.sort_values(by='Tanggal', ascending=True)
        df_barang_terbaru = df_sorted.drop_duplicates(subset='Nama Barang', keep='last').reset_index(drop=True)
        # Check for duplicates
        check_duplicates(df_barang_terbaru, month_name)
        
        # Save to output file
        df_barang_terbaru.to_excel(output_file, index=False)
        print(f"Successfully saved to {output_file}\n")
        
    except Exception as e:
        print(f"Error processing {month_name}: {str(e)}\n")

# Process all months
for month_num, month_name in MONTHS.items():
    process_month(month_num, month_name)

print("All months processed!")