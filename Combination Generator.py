#Combination upto 1M for dynamic row and columns

import pandas as pd
import itertools
from tqdm import tqdm
import math

EXCEL_ROW_LIMIT = 1048576  # Excel's max row limit
BATCH_SIZE = 1000000         # Rows to write at once

# Read and clean data
temp = pd.read_excel('InputValues.xlsx', sheet_name='Input Master')
df1 = temp.fillna("0")

# Drop dynamic rows and columns
arr_index = list(range(1, temp.index[-1] + 1))
trydf = temp.drop(arr_index).dropna(axis=1)

headerlist = list(trydf)
header_count = len(headerlist)

print(f"Detected {header_count} headers, generating combinations directly to Excel...")

# Get only non-'0' values for each column
value_lists = [df1[col][df1[col] != '0'].tolist() for col in headerlist]

# Calculate total combinations
total_combinations = math.prod(len(values) for values in value_lists)
print(f"Total combinations to generate: {total_combinations:,}")

current_time = pd.Timestamp.now().strftime('%Y_%m_%d_%H_%M_%S')
file_name = 'CombinedData_'+ current_time +'.xlsx'

# Prepare Excel writer
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

# Write input summary sheet
temp.to_excel(writer, sheet_name="Input Master", index=False)

# Generate combinations and write in chunks
combo_index = 1
sheet_number = 1
rows_written = 0
batch = []

for combo in tqdm(itertools.product(*value_lists), total=total_combinations, unit="rows"):
    combo_str = "_".join(map(str, combo))
    row = [combo_index] + list(combo) + [combo_str]
    batch.append(row)
    combo_index += 1
    rows_written += 1

    # If batch is full or Excel sheet limit reached
    if len(batch) >= BATCH_SIZE or rows_written >= EXCEL_ROW_LIMIT - 1:
        df_batch = pd.DataFrame(batch, columns=["Index"] + headerlist + ["Combined_String"])
        sheet_name = f"Output {sheet_number}"
        df_batch.to_excel(writer, sheet_name=sheet_name, index=False)
        batch.clear()
        sheet_number += 1
        rows_written = 0

# Write any remaining rows
if batch:
    df_batch = pd.DataFrame(batch, columns=["Index"] + headerlist + ["Combined_String"])
    sheet_name = f"Output {sheet_number}"
    df_batch.to_excel(writer, sheet_name=sheet_name, index=False)

writer.close()
print("Excel file generated successfully.")