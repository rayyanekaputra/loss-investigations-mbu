import pandas as pd
from fuzzywuzzy import process

# Read the excel file
df_penjualan = pd.read_excel("./BAEKMI/01 penjualan januari.xlsx")
print(df_penjualan.info())
df_beli_supplier = pd.read_excel("./BAEKMI/01 supplier januari.xlsx")
print(df_beli_supplier.info())

# Create a dictionary mapping from df_beli_supplier's complete names to suppliers
name_to_supplier = dict(zip(df_beli_supplier['Nama Barang'], df_beli_supplier['Pemasok']))

# Function to find the best match for each partial name
def find_supplier(partial_name):
    # Get the best match from df_beli_supplier's complete names
    result = process.extractOne(partial_name, name_to_supplier.keys())
    if result:  # If a match was found
        match, score = result
        # Only return the supplier if we have a good match (adjust threshold as needed)
        return name_to_supplier[match] if score >= 80 else None
    return None

# Apply the function to create a new Pemasok column in df_penjualan
df_penjualan['Pemasok'] = df_penjualan['Nama Barang'].apply(find_supplier)

# Check how many matches were successful
matched_count = df_penjualan['Pemasok'].notna().sum()
total_count = len(df_penjualan)
print(f"Successfully matched {matched_count} out of {total_count} rows ({matched_count/total_count:.1%})")

# Optional: Save unmatched items for review
unmatched = df_penjualan[df_penjualan['Pemasok'].isna()]
print(f"\nUnmatched items sample:")
print(unmatched['Nama Barang'].head(10).to_string(index=False))