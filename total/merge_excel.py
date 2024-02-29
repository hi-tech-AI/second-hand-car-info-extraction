import pandas as pd
import os

merged_data = pd.DataFrame()
directory = 'excel'
for index, filename in enumerate(os.listdir(directory)):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):  # Filter only Excel files
        filepath = os.path.join(directory, filename)
        df = pd.read_excel(filepath)  # Read each Excel file into a DataFrame
        merged_data = pd.concat([merged_data, df], ignore_index = True)
        print(index + 1)

merged_data.to_excel('merge.xlsx', index = False)