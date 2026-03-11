import os
import re
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name)

def export_sheets_fast():
    # === CONFIGURATION ===
    output_folder = r"E:\Query\\"
    os.makedirs(output_folder, exist_ok=True)

    # === Select One Workbook ===
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select a workbook to process",
        filetypes=[("Excel Files", "*.xls*")]
    )
    if not file_path:
        print("No file selected. Exiting.")
        return

    # === Read all sheets at once ===
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

    for sheet_name, df in sheets.items():
        # Drop completely empty columns
        df = df.dropna(axis=1, how="all")

        # Deduplicate headers (keep first occurrence)
        headers = []
        new_cols = []
        for col in df.columns:
            col_str = str(col).strip()
            if col_str not in headers:
                headers.append(col_str)
                new_cols.append(col_str)
            else:
                new_cols.append(None)
        df.columns = new_cols
        df = df.loc[:, df.columns.notnull()]

        # Sanitize filename
        safe_name = sanitize_filename(sheet_name)

        # Save each sheet as separate workbook
        export_path = os.path.join(output_folder, f"{safe_name}.xlsx")
        # Use ExcelWriter to control sheet name
        with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=safe_name, index=False)

    print("✅ All sheets exported quickly with sheet name same as file name.")

if __name__ == "__main__":
    export_sheets_fast()

