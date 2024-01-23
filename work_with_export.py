import csv
from openpyxl import Workbook
import os
import sys
from config.confidential_info import operating_file_path_instance
import subprocess

def find_csv_files(folder_path, prefix="NA", suffix=".csv"):
    return [f for f in os.listdir(folder_path) if f.startswith(prefix) and f.endswith(suffix)]


def read_csv_file(file_path):
    with open(file_path, 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)  # skip header row
        return list(reader)


def replace_values(data, replacements):
    return [[replacements.get(cell, cell) for cell in row] for row in data]


def create_workbook(data):
    wb = Workbook()
    ws = wb.active
    for row_data in data:
        ws.append(row_data)
    return wb


def save_workbook(wb, file_path):
    wb.save(file_path)


def process_csv_file(csv_file_path, replacements, processed_file_suffix="_processed.xlsx"):
    data = read_csv_file(csv_file_path)
    processed_data = replace_values(data, replacements)
    wb = create_workbook(processed_data)
    processed_file_path = os.path.splitext(
        csv_file_path)[0] + processed_file_suffix
    save_workbook(wb, processed_file_path)
    return processed_file_path


def main():
    downloads_folder_path = os.path.join(os.path.expanduser("~"), "Downloads")
    matching_files = find_csv_files(downloads_folder_path)

    if not matching_files:
        print("No matching CSV files found in Downloads folder. OR No COA's Posted, please try your query manually to confirm.")
        return  # Exit the function cleanly without sys.exit()

    downloaded_file_path = os.path.join(downloads_folder_path, matching_files[0])
    # this is equal to find and replace in excel
    replacements = {"NT": 0, "ND": 0, "PASS": 0, "<0.020": 0}

    try:
        processed_file_path = process_csv_file(downloaded_file_path, replacements)
        print(f"Data processed successfully. New workbook saved as: {processed_file_path}")

        # Use subprocess.run instead of os.system
        subprocess.run(['python', operating_file_path_instance.add_to_local], check=True, shell=True)
        # print("Chained script add to local execution completed.") #debug statement, not needed to run script

    except subprocess.CalledProcessError as e:
        print(f"An error occurred while running the add to local script: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()