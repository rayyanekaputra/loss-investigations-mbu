import pandas as pd
import os

# Month mapping (number to month name)
months = {
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

def clean_penjualan(df):
    """Clean penjualan dataframe"""
    # Clean columns
    df.dropna(axis="columns", how="all", inplace=True)
    df.rename(columns=df.iloc[3], inplace=True)
    df.drop(range(0,4), axis=0, inplace=True)
    df = df.loc[:, ~df.columns.isna()]
    df.reset_index(drop=True, inplace=True)
    
    # Clean data
    df['Tanggal'] = pd.to_datetime(df['Tanggal'])
    df['Kode #'] = df['Kode #'].astype(str)
    df['Kode #'] = df['Kode #'].apply(
        lambda x: x.lstrip('0') if x.lstrip('0') else '0'
    )
    return df

def clean_beli_supplier(df):
    """Clean pembelian supplier dataframe"""
    # Clean columns
    df.dropna(axis="columns", how="all", inplace=True)
    df.rename(columns=df.iloc[3], inplace=True)
    df.drop(range(0,4), axis=0, inplace=True)
    df = df.loc[:, ~df.columns.isna()]
    df.reset_index(drop=True, inplace=True)
    
    # Clean data
    duplicate = df['Nama Barang'] == 'Nama Barang'
    df = df[~duplicate]
    df.dropna(how='all', inplace=True)
    return df

def process_monthly_data():
    for num, month in months.items():
        folder_path = f"./{num} {month}/"
        
        # Check if folder exists
        if not os.path.exists(folder_path):
            print(f"Folder not found: {folder_path}")
            continue
            
        try:
            # Load and clean penjualan data
            penjualan_file = f"Penjualan per Barang {month}.xlsx"
            df_penjualan = pd.read_excel(folder_path + penjualan_file)
            df_penjualan = clean_penjualan(df_penjualan)
            
            # Load and clean supplier data
            supplier_file = f"Pembelian per Barang dan Supplier {month}.xlsx"
            df_beli_supplier = pd.read_excel(folder_path + supplier_file)
            df_beli_supplier = clean_beli_supplier(df_beli_supplier)
            
            # # Merge data
            # merged_df = pd.merge(
            #     left=df_penjualan, 
            #     right=df_beli_supplier, 
            #     on="Nama Barang", 
            #     how="left"
            # )
            
            # Save files with numbering prefix (01-12)
            df_penjualan.to_excel(f"{num} penjualan {month.lower()}.xlsx", index=False)
            df_beli_supplier.to_excel(f"{num} supplier {month.lower()}.xlsx", index=False)
            # merged_df.to_excel(f"{num} merged_penjualan_supplier {month.lower()}.xlsx", index=False)
            
            print(f"Successfully processed {month} data")
            
        except FileNotFoundError as e:
            print(f"File not found in {folder_path}: {e}")
        except Exception as e:
            print(f"Error processing {month}: {str(e)}")

if __name__ == "__main__":
    process_monthly_data()