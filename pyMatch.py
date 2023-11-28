import pandas as pd

# Load the Excel workbook from a path
workbook_path = r'Path/to/excel/file'
df = pd.read_excel(workbook_path, sheet_name='Visits')

# Assuming your columns are named 'keys', 'values', and 'valueToConvert'
keys_column = 'keys'
values_column = 'values'
column_to_convert = 'valueToConvert'

# Create a dictionary mapping the first two columsn
key_values = dict(zip(df[values_column], df[keys_column]))

# Create a new column 'convertedValues' with the converted values
df['convertedValues'] = df[column_to_convert].map(key_values)

# Save the updates to an excel file
updated_workbook_path = r'path/to/excel/output/file'
df.to_excel(updated_workbook_path, index=False)

print(f"Excel file with converted values saved to: {updated_workbook_path}")