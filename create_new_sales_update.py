import os
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

def find_latest_sales_update(directory):
    sales_update_files = list(directory.glob("*Sales Update*.xlsx"))
    if not sales_update_files:
        raise FileNotFoundError("No Sales Update file found.")
    return max(sales_update_files, key=os.path.getmtime)

def create_new_sales_update(import_path):
    # Find the latest Sales Update file
    latest_file = find_latest_sales_update(import_path)
    print(f"Found latest Sales Update file: {latest_file.name}")

    # Create new filename with today's date
    today_date = datetime.now().strftime("%m.%d")
    new_filename = f"Sales Update {today_date}.xlsx"
    new_file_path = import_path / new_filename

    # Load the workbook
    wb = load_workbook(filename=latest_file, data_only=False)

    # Save as new file
    wb.save(new_file_path)

    print(f"Created new Sales Update file: {new_filename}")

    return new_file_path

if __name__ == "__main__":
    import_path = Path("C:/Users/Shadow/projects/analyitics_automator")
    new_file = create_new_sales_update(import_path)
    print(f"New Sales Update file created at: {new_file}")