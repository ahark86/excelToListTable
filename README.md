# Excel to rST List-Table Converter

## Description

Input an Excel sheet and receive a .txt document with the Excel data formatted as an rST List-Table. The resulting List-Table is ready to be pasted into a .rst document.

## How to use

When running the script, you can pass in the required parameters as command line arguments for quickest results in the following order:

1. Absolute path to the Excel file
1. Name of the target worksheet within the chosen Excel doc
1. Number of columns in the target sheet
1. Number of rows in the target sheet
1. Absolute path for the desired output document

Example:

`python3 excelToListTable.py "/path/to/excelsheet" "My_Sheet" "5" "6" "/path/to/desired_output_file.txt"`

If you do not provide the command line arguments, you can enter them in prompts once the code is run.


