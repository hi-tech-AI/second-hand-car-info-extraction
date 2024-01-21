import pandas as pd

# Read the Excel file
df = pd.read_excel("total.xlsx")

# Remove duplicates based on all columns
df = df.drop_duplicates()

# Save the updated DataFrame to a new Excel file
df.to_excel(f'{len(df)}.xlsx', index=False)