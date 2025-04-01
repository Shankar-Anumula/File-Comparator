# Excel File Comparison Tool

This Python script compares Excel files from a source directory against target files using a lookup schema. The comparison highlights differences between the source and target data, such as missing, extra, or mismatched columns and rows. The results are output to an Excel file, making it easy to review the discrepancies.

## Features

- Compares Excel files from a source directory with those from a target directory.
- Utilizes a lookup directory for schema definition (column names, data types, and primary keys).
- Detects missing and extra columns between source and target files.
- Highlights differences in values between matching columns.
- Generates a summary report and highlights mismatched records in the output.

## Prerequisites

Before using this script, you need to have Python and the following libraries installed:

- `pandas` for handling Excel data and generating reports.
- `openpyxl` for reading and writing Excel files and applying styles.
  
To install these dependencies, run the following command:

```bash
pip install pandas openpyxl
