import os
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def robocopy_files_from_excel():
    # Prompt user to select Excel file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel file with file copy list",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        print("No file selected. Exiting.")
        return

    # Read Excel file
    df = pd.read_excel(file_path)

    # Check required columns exist
    required_cols = ['Name', 'Folder Path', 'Destination Folder']
    if not all(col in df.columns for col in required_cols):
        print(f"Excel must contain columns: {required_cols}")
        return

    for index, row in df.iterrows():
        file_name = str(row['Name']).strip()
        src_folder = str(row['Folder Path']).strip()
        dest_folder = str(row['Destination Folder']).strip()
        src_path = os.path.join(src_folder, file_name)

        if not os.path.exists(src_path):
            print(f"Source file does not exist: {src_path}")
            continue

        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)

        command = ['robocopy', src_folder, dest_folder, file_name, '/MT:48', '/NFL', '/NDL', '/NJH', '/NJS']
        print(f"Copying '{file_name}' from '{src_folder}' to '{dest_folder}'...")
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode <= 7:
            print(f"Copied successfully.")
        else:
            print(f"Error copying '{file_name}': {result.stderr.strip()}")

if __name__ == "__main__":
    robocopy_files_from_excel()
