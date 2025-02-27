# Student Report Generator

## 📌 Overview

The Student Report Generator is a Python script designed to automate the generation of individual student reports from a master Excel sheet. The script reads student data, replaces placeholders in a report template, and saves the generated reports in a designated output folder.

## 🚀 Features

- Customizable Configuration: Uses a JSON config file for easy customization.

- Batch Processing: Supports processing a specific range of rows.

- Column Mapping: Dynamically maps master file columns to template cells.

- Formatted Output: Ensures text fits inside Excel cells by adjusting widths and heights.

- Command-Line Execution: Allows running the script with different configurations.

- Progress Tracking: Displays report count and progress updates.

## 📂 Project Structure

📦 Student_Report_Generator\

| Structure                            | Description                                                                                     |
|:-------------------------------------|-------------------------------------------------------------------------------------------------|
| │-- 📄 **reports_generater.py**      | # Main script                                                                                   |
| │-- 📄 **config.json**               | # Configuration file which contains the mapping between Master sheet and the individual report. |
| │-- 📋 **Master_Marking_Sheet.xlsx** | # The master marking sheet where each row is a student report.                                  |  
| │-- 📋 **Report_Template.xlsx**      | # This is the individual marking feedback report.                                               |  
| │-- 📂 reports/                      | # Output folder for generated reports.                                                          |

## ⚙️ Configuration

The script relies on a config.json file to manage its input and output settings. Below is an example configuration along with mapping between master marking sheet and the individual marking feedback report:

```sh
{
  "master_file": "master.xlsx",
  "template_file": "template.xlsx",
  "output_folder": "reports/",
  "assignment_name": "COMP8760",
  "sheet_name": "Sheet1",
  "key_column": "Login",
  "start_row": 1,
  "end_row": 10,
  "mapping": {
    "Forename": "B2",
    "Surname": "B3",
    "Q1": "C5",
    "Q2": "C6",
    "Q3": "C7"
  }
}
```

### Configuration Parameters

| Parameter       | Description                                               |
|-----------------|-----------------------------------------------------------|
| master_file     | Path to the master Excel file containing student data.    |
| template_file   | Path to the Excel report template.                        |
| output_folder   | Directory where reports will be saved.                    |
| assignment_name | Prefix for the report filenames.                          |
| sheet_name      | Sheet name in the master file.                            |
| key_column      | Column used for generating report filenames.              |
| start_row       | First row to process (for batch processing).              |
| end_row         | Last row to process (for batch processing).               |
| mapping         | Dictionary mapping column names to Excel cell references. |

## ▶️ Usage

Running the Script

To execute the script with the default configuration:

```sh
python reports_generater.py
```

To specify a different config file:

```sh
python reports_generater.py --config custom_config.json
```

### Expected Output

✅ Report saved: reports/COMP8760 - student123.xlsx\
✅ Report saved: reports/COMP8760 - student456.xlsx\
✅ All reports generated successfully. Total reports: 10

## 🤝 Contributing

Contributions are welcome! Feel free to open an issue or submit a pull request.

## 📜 License

This project is licensed under the MIT License.

🚀 Enjoy using the Student Report Generator! If you find this project useful, give it a ⭐ on GitHub!
