import os
import shutil
import pandas as pd


def robocopy_files_from_excel(file_path):
    # Check file exists
    if not os.path.exists(file_path):
        print(f"Excel file not found: {file_path}")
        return

    # Read Excel
    df = pd.read_excel(file_path)

    # Validate columns
    required_cols = ["Name", "Folder Path", "Destination Folder"]

    if not all(col in df.columns for col in required_cols):
        print(f"Excel must contain columns: {required_cols}")
        return

    # Process rows
    for _, row in df.iterrows():

        file_name = str(row["Name"]).strip()
        src_folder = str(row["Folder Path"]).strip()
        dest_folder = str(row["Destination Folder"]).strip()

        src_path = os.path.join(src_folder, file_name)

        if not os.path.isfile(src_path):
            print(f"Source file not found: {src_path}")
            continue

        os.makedirs(dest_folder, exist_ok=True)

        try:
            dest_path = os.path.join(dest_folder, file_name)

            shutil.copy2(src_path, dest_path)

            print(f"✅ Copied: {file_name}")

        except Exception as e:
            print(f"❌ Error copying {file_name}: {e}")


if __name__ == "__main__":

    # Example
    excel_file = "uploads/input.xlsx"

    robocopy_files_from_excel(excel_file)