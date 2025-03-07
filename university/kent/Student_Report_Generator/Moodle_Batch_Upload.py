import json
import os

import pandas as pd


def load_config(config_file):
    """Load configuration from JSON file."""
    with open(config_file, "r") as f:
        return json.load(f)


def load_data(master_file, reference_file, sheet_name):
    """Load master and reference Excel sheets into DataFrames."""
    df_master = pd.read_excel(master_file, sheet_name=sheet_name)
    df_reference = pd.read_excel(reference_file)

    # Ensure column names are consistent
    df_master.rename(columns=str.strip, inplace=True)
    df_reference.rename(columns=str.strip, inplace=True)

    return df_master, df_reference


def preprocess_data(df_master, df_reference):
    """Preprocess data: normalize login values and identify missing students."""
    df_master["Login"] = df_master["Login"].str.lower()
    df_reference["Login"] = df_reference["Email address"].str.replace("@kent.ac.uk", "", regex=False)
    df_reference["Login"] = df_reference["Login"].str.lower()

    # Identify students in df_master but not in df_reference:
    missing_students = df_master[~df_master["Login"].isin(df_reference["Login"])]

    if not missing_students.empty:
        print("⚠️ The following students are in the master file but missing from the reference file:")
        print(missing_students["Login"].to_string(index=False))

    # Keep the common rows:
    df_reference = df_reference[df_reference["Login"].isin(df_master["Login"])].copy()

    # Create "Submission_ID" by extracting the second part of the "Identifier" column
    df_reference["Submission_ID"] = df_reference["Identifier"].apply(
        lambda x: str(x).split(" ")[1] if isinstance(x, str) and " " in x else ""
    )

    # Create "submission_file_name" using name + submission_id + module_name + assignment_name
    df_reference.loc[:, "submission_file_name"] = df_reference["Full name"] + "_" + \
                                                  df_reference["Submission_ID"] + "_assignsubmission_file_" + \
                                                  module_name + "_" + assignment_name + "_Feedback"

    # Daniel Iyare_117863_assignsubmission_file_COMPXXXX_A2_Feedback

    return df_master, df_reference


def update_reference_file(df_master, df_reference, output_file):
    """Update the reference file with grading and marking workflow state."""

    df_reference["Marking workflow state"] = "Released"
    df_reference["Grade"] = df_reference["Login"].map(df_master.set_index("Login")["Grade"])

    # Remove unnecessary columns
    df_reference.drop(columns=["Login", "Submission_ID"], inplace=True)

    """Save the updated reference file."""
    df_reference.to_excel(output_file, index=False)
    print(f"✅ Updated reference file saved as: {output_file}")


def main():
    """Main function to execute the workflow."""
    config_file = "config.json"
    config = load_config(config_file)

    master_file = config["master_file"]
    moodle_folder = config["moodle_folder"]
    reference_file = os.path.join(moodle_folder, config["reference_file"])
    output_reference_file = os.path.join(moodle_folder, config["output_reference_file"])
    sheet_name = config["sheet_name"]

    df_master, df_reference = load_data(master_file, reference_file, sheet_name)
    df_master, df_reference = preprocess_data(df_master, df_reference)
    update_reference_file(df_master, df_reference, output_reference_file)


if __name__ == "__main__":
    main()
