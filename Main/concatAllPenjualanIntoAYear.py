import pandas as pd
import os

# Dictionary to map month names to their numbers (for sorting)
month_order = {
    'januari': 1, 'februari': 2, 'maret': 3, 'april': 4, 'mei': 5, 'juni': 6,
    'juli': 7, 'agustus': 8, 'september': 9, 'oktober': 10, 'november': 11, 'desember': 12
}

# List to store all dataframes
all_dfs = []

# Get list of files in the directory
directory = './BAEKMI/'
for filename in os.listdir(directory):
    if filename.endswith('.xlsx') and 'penjualan' in filename.lower():
        try:
            # Extract month name from filename
            parts = filename.lower().split()
            month = parts[-1].replace('.xlsx', '')
            
            # Get the file path
            filepath = os.path.join(directory, filename)
            
            # Read the Excel file
            df = pd.read_excel(filepath)
            
            # Add a column for the month
            df['Bulan'] = month.capitalize()
            df['Bulan_Num'] = month_order[month.lower()]
            
            # Add to our list
            all_dfs.append(df)
            
            print(f"Processed: {filename}")
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")

# Check if we found any files
if not all_dfs:
    print("No sales files found in the directory!")
else:
    # Concatenate all dataframes
    combined_df = pd.concat(all_dfs, ignore_index=True)
    
    # Sort by month number (optional)
    combined_df.sort_values('Bulan_Num', inplace=True)
    
    # Drop the month number column if                                                                                                                                                                                                                   you don't need it
    combined_df.drop('Bulan_Num', axis=1, inplace=True)
    
    # Export to Excel
    output_path = os.path.join(directory, 'Penjualan2024.xlsx')
    combined_df.to_excel(output_path, index=False)
    
    print(f"\nSuccessfully combined {len(all_dfs)} monthly files into {output_path}")