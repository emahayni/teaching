# Student Report Generator

## 📌 Overview
This Python script automates the generation of student reports from a master Excel sheet. It reads student data, replaces placeholders in a report template, and saves individual reports based on a predefined configuration.

## 🚀 Features
- Reads student data from a **master Excel file**.
- Uses a **template Excel file** to generate reports.
- Supports **custom column mappings** for dynamic report generation.
- Processes a **range of rows** to generate reports selectively.
- Automatically **adjusts column widths and row heights** to fit text.
- Outputs reports to a specified folder.
- Command-line support for specifying a custom config file.

## 🛠️ Installation
Ensure you have **Python 3.x** installed. Then, install the required dependencies:
```sh
pip install pandas openpyxl
