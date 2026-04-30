# import os
# import pandas as pd
# import win32com.client as win32

# def process_excel_files():
#     folder_path = r"C:\Users\Shreyas Shah\Downloads\ZSD_064"
    
#     # 1. Convert .xls to .xlsx using Excel Application (ensures data integrity)
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     excel.Visible = False
#     excel.DisplayAlerts = False

#     for i in range(1, 36):
#         file_name = f"{i:02d}"
#         old_path = os.path.join(folder_path, f"{file_name}.xls")
#         new_path = os.path.join(folder_path, f"{file_name}.xlsx")

#         if os.path.exists(old_path):
#             wb = excel.Workbooks.Open(old_path)
#             # FileFormat 51 is .xlsx
#             wb.SaveAs(new_path, FileFormat=51)
#             wb.Close()
#             print(f"Converted {file_name}.xls to .xlsx")

#     excel.Quit()

#     # 2. Process rows and columns (Ignore file '01')
#     for i in range(2, 36):
#         file_name = f"{i:02d}.xlsx"
#         file_path = os.path.join(folder_path, file_name)

#         if os.path.exists(file_path):
#             # Read the file
#             df = pd.read_excel(file_path, header=None)

#             # Remove 1st column (Index 0)
#             df.drop(df.columns[0], axis=1, inplace=True)

#             # Remove 1st, 2nd, 3rd rows (Indices 0, 1, 2) and 5th row (Index 4)
#             # We drop them by index. 
#             df.drop([0, 1, 2, 4], axis=0, inplace=True)

#             # Save the modified file
#             df.to_excel(file_path, index=False, header=False)
#             print(f"Processed data for {file_name}")

# if __name__ == "__main__":
#     process_excel_files()

# import pandas as pd
# import os

# # Define the directory path
# folder_path = r'C:\Users\Shreyas Shah\Downloads\ZSD_064'
# output_file = os.path.join(folder_path, 'MERGED.xlsx')

# # Get a list of all XLSX files in the folder
# file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# # Initialize an empty list to hold dataframes
# all_dataframes = []

# print(f"Found {len(file_list)} files. Starting merge...")

# for file in file_list:
#     file_path = os.path.join(folder_path, file)
    
#     # Read the excel file
#     # Note: This reads the first sheet by default
#     df = pd.read_excel(file_path)
    
#     # Optional: Add a column to track which file the data came from
#     df['Source_File'] = file
    
#     all_dataframes.append(df)

# # Concatenate all dataframes into one
# merged_df = pd.concat(all_dataframes, ignore_index=True)

# # Save the result to a new Excel file
# merged_df.to_excel(output_file, index=False)

# print(f"Success! Merged file saved at: {output_file}")

import pandas as pd
import os

# Define the directory path
folder_path = r'C:\Users\Shreyas Shah\Downloads\ZSD_064'

# Get a list of all XLSX files
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

total_rows = 0

print("Analyzing files...")

for file in file_list:
    file_path = os.path.join(folder_path, file)
    
    # Load the file
    df = pd.read_excel(file_path)
    
    # Get row count for this specific file (excluding header)
    current_file_rows = len(df)
    total_rows += current_file_rows
    
    print(f"{file}: {current_file_rows} rows")

# Apply your formula: ((Total Rows) - 36) / 2
result = (total_rows - 36) / 2

print("-" * 30)
print(f"Grand Total Rows: {total_rows}")
print(f"Calculation Result [({total_rows} - 36) / 2]: {result}")