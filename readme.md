    # Excel Data Copy Script

This script copies data from one Excel file to a new Excel file using the `win32com` library in Python. It allows you to select specific columns to copy and paste them into a new workbook.

## Prerequisites

- Python installed on your system.
- The `os` and `win32com.client` modules are required. You can install `pywin32` package to access `win32com` functionalities.

```bash
pip install pywin32

Usage
Place the script in the same directory as your source Excel file, named content.xlsx.

Run the script.

You will be prompted to enter the column names you want to copy. Input the column names separated by commas.

The script will create a new Excel file named Hammadde.xlsx in the same directory, containing the selected columns from the source file.

If any errors occur, the script will display an error message and wait for 10 seconds before exiting.

Author
Furkan

Feel free to modify and improve this script according to your specific needs. If you encounter any issues, please refer to the error message displayed in case of a problem.

For any questions or assistance, you can reach out to furkancakir.dev.