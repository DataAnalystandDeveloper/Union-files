# Excel Data Consolidation Script
This Python script reads multiple Excel files from a specified folder, extracts data from a particular sheet ("Summary"), processes it by removing unnecessary rows, and consolidates it into a single Excel output file.
## Overview
The script reads all `.xlsx` files in a specified input directory, searches for a sheet named "Summary" in each file, and:
1. Identifies and removes rows before the "Location" keyword in Column B.
2. Removes the last two rows of the sheet.
3. Cleans up any empty columns or rows before exporting the consolidated data.
The cleaned data is saved to an Excel file without headers or index columns in the specified output directory.
## Features
- **Batch Processing**: Reads all `.xlsx` files in a folder.
- **Data Filtering**: Retains only relevant rows by locating the "Location" keyword.
- **Data Cleaning**: Removes empty columns and rows.
- **Consolidated Output**: Exports a single consolidated file to a specified directory.
