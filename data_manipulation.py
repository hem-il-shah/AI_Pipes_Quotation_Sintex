import os
import pandas as pd
import win32com.client as win32

def process_excel_files():
    folder_path = r"C:\Users\Shreyas Shah\Downloads\ZSD_064"
    
    # 1. Convert .xls to .xlsx using Excel Application (ensures data integrity)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    for i in range(1, 36):
        file_name = f"{i:02d}"
        old_path = os.path.join(folder_path, f"{file_name}.xls")
        new_path = os.path.join(folder_path, f"{file_name}.xlsx")

        if os.path.exists(old_path):
            wb = excel.Workbooks.Open(old_path)
            # FileFormat 51 is .xlsx
            wb.SaveAs(new_path, FileFormat=51)
            wb.Close()
            print(f"Converted {file_name}.xls to .xlsx")

    excel.Quit()

    # 2. Process rows and columns (Ignore file '01')
    for i in range(2, 36):
        file_name = f"{i:02d}.xlsx"
        file_path = os.path.join(folder_path, file_name)

        if os.path.exists(file_path):
            # Read the file
            df = pd.read_excel(file_path, header=None)

            # Remove 1st column (Index 0)
            df.drop(df.columns[0], axis=1, inplace=True)

            # Remove 1st, 2nd, 3rd rows (Indices 0, 1, 2) and 5th row (Index 4)
            # We drop them by index. 
            df.drop([0, 1, 2, 4], axis=0, inplace=True)

            # Save the modified file
            df.to_excel(file_path, index=False, header=False)
            print(f"Processed data for {file_name}")

if __name__ == "__main__":
    process_excel_files()