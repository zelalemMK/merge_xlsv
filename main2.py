import pandas as pd

# Replace these paths with the actual paths to your Excel files
path_file1 = "./A.xlsx"
path_file2 = "./B.xlsx"


# Function to read an Excel file and standardize column names by stripping whitespace
def load_excel(file_path):
    df = pd.read_excel(file_path)
    df.columns = (
        df.columns.str.strip()
    )  # Remove any leading/trailing whitespace from column names
    return df


# Load the files into pandas DataFrames
df1 = load_excel(path_file1)
df2 = load_excel(path_file2)

# Define the columns to merge on and to calculate
merge_columns = [
    "Name",
    "Variant code / SKU",
    "Category",
    "Default supplier",
    "Units of measure",
    "Default storage bin",
]
calc_columns = [
    "Average cost",
    "Value in stock",
    "In stock",
    "Expected",
    "Committed",
    "Reorder point",
    "Missing / Excess",
]

# Merge the DataFrames on the specified columns
merged_df = pd.merge(
    df1, df2, on=merge_columns, how="outer", suffixes=("_file1", "_file2")
)


# Before calculation, let's ensure all calc_columns have numeric data, converting if necessary
for col in calc_columns:
    merged_df[col + "_file1"] = pd.to_numeric(
        merged_df[col + "_file1"], errors="coerce"
    )
    merged_df[col + "_file2"] = pd.to_numeric(
        merged_df[col + "_file2"], errors="coerce"
    )

    # Calculate the sum of the specified columns and add it to the merged DataFrame
    merged_df[col] = merged_df[col + "_file1"].fillna(0) + merged_df[
        col + "_file2"
    ].fillna(0)

# Drop the duplicated columns after the calculation
merged_df.drop(
    [col + suffix for col in calc_columns for suffix in ["_file1", "_file2"]],
    axis=1,
    inplace=True,
)

merged_df = merged_df[merged_df["Category"] != "Finish Supplies"]

# Optional: Check for any rows that have NaN values after the merge to diagnose issues
nan_rows = merged_df[merged_df.isnull().any(axis=1)]
if not nan_rows.empty:
    print(
        f"There are rows with NaN values after the merge, which may indicate unmatched keys: \n{nan_rows}"
    )

# Export the merged DataFrame to a new Excel file
output_path = "merged_output.xlsx"
merged_df.to_excel(output_path, index=False)

print(f"Merged file saved to {output_path}")
