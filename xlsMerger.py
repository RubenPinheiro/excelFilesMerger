import pandas as pd
import os

# Specify the directory where the Excel files are located
directory = ".venv/xlsFiles"

# List all Excel files in the directory
file_names = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith((".xlsx", ".xls"))]
# print(file_names))

# Create a writer object
with pd.ExcelWriter('ULSPV_CHPVVC.xlsx', engine='openpyxl') as writer:
    # Loop through each file
    for file_name in file_names:
        # Extract the sheet name from the file name
        sheet_name = os.path.splitext(os.path.basename(file_name))[0]
        # Read the Excel file
        df = pd.read_excel(file_name)
        # Write the DataFrame to a new sheet with file name as sheet name
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Merged file created successfully!")
