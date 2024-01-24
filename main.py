import pandas as pd

# Load the first Excel spreadsheet into a DataFrame
df1 = pd.read_excel('./A.xlsx')

# Load the second Excel spreadsheet into another DataFrame
df2 = pd.read_excel('./B.xlsx')

# Merge DataFrames based on the "Variant code / SKU" column
merged_df = pd.merge(df1, df2, on='Variant code / SKU', how='inner', suffixes=('_first', '_last'))

# Update columns 1-7 with the first value
merged_df = merged_df.assign(
    Name=merged_df['Name_first'],
    Category=merged_df['Category_first'],
    Default_supplier=merged_df['Default supplier_first'],
    Units_of_measure=merged_df['Units of measure_first'],
    Default_storage_bin=merged_df['Default storage bin_first'],
    Average_cost=merged_df['Average cost_first'],
    Value_in_stock=merged_df['Value in stock_first']
)

# Add values for columns 8-13
merged_df = merged_df.assign(
    In_stock=merged_df['In stock_first'] + merged_df['In stock_last'],
    Expected=merged_df['Expected_first'] + merged_df['Expected_last'],
    Committed=merged_df['Committed_first'] + merged_df['Committed_last'],
    Reorder_point=merged_df['Reorder point_first'] + merged_df['Reorder point_last'],
    Missing_or_Excess=merged_df['Missing / Excess_first'] + merged_df['Missing / Excess_last'],
)

# Update column 14 with the first value
merged_df = merged_df.assign(Location=merged_df['Location_first'])

# Drop the redundant columns (optional)
merged_df.drop(columns=merged_df.filter(like='_first').columns, inplace=True)

# Save the merged DataFrame to a new Excel file
merged_df.to_excel('merged_file.xlsx', index=False)
