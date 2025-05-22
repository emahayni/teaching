import json
import os
import zipfile

import pandas as pd


def load_config(config_path):
    """Load configuration from JSON file."""
    with open(config_path, "r") as f:
        return json.load(f)


def clean_output_folder(output_folder):
    """Remove all existing .xlsx files from the output folder."""
    if os.path.exists(output_folder):
        for file in os.listdir(output_folder):
            if file.endswith(".xlsx"):
                os.remove(os.path.join(output_folder, file))
        print(f"🧹 Cleaned output folder: {output_folder}")


def excel_cell_ref_to_indices(ref):
    col = ord(ref[0].upper()) - ord('A')
    row = int(ref[1:]) - 1
    return row, col


def zip_output_files(output_folder, zip_filepath):
    """Zip generated reports into a single archive."""
    with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(output_folder):
            zipf.write(os.path.join(output_folder, file), file)
    print(f'📦 Reports zipped successfully: {zip_filepath}')


def load_dataframe(filename, sheet_name=0):
    """Load marking data from a CSV or XLSX file into a DataFrame."""
    file_extension = os.path.splitext(filename)[-1].lower()
    if file_extension == ".csv":
        df = pd.read_csv(filename)
    elif file_extension in [".xls", ".xlsx"]:
        df = pd.read_excel(filename, sheet_name=sheet_name)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}")

    # Ensure column names are consistent
    df.rename(columns=str.strip, inplace=True)

    return df
