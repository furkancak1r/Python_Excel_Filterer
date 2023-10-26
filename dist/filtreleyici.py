import os
import win32com.client as win32
import time

try:
    # Get the current working directory
    current_dir = os.getcwd()

    # Excel application initialization
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    # Open the Excel file
    input_file = os.path.join(current_dir, 'content.xlsx')
    workbook = excel.Workbooks.Open(input_file)

    # Copy and paste data
    worksheet = workbook.Worksheets(1)

    # Create a new Excel file
    new_workbook = excel.Workbooks.Add()
    new_worksheet = new_workbook.Worksheets(1)

    # Get user input for columns to copy
    columns_to_copy = input("Kopyalanacak sütun adlarını girin (virgülle ayrılmış): ").split(',')
    target_columns = []

    # Generate target column names alphabetically based on the number of columns to copy
    for i in range(len(columns_to_copy)):
        target_columns.append(chr(ord('A') + i))

    for i in range(len(columns_to_copy)):
        source_col_range = worksheet.Range(columns_to_copy[i] + ":" + columns_to_copy[i])
        source_col_range.Copy()
        
        target_col = target_columns[i]
        target_col_range = new_worksheet.Range(target_col + "1")
        target_col_range.PasteSpecial()

    # Save the new Excel file
    output_file = os.path.join(current_dir, 'Hammadde.xlsx')
    new_workbook.SaveAs(output_file)
    new_workbook.Close()

except Exception as e:
    print("Hata: " + str(e))
    print("content.xlsx bulunamadı!")
    time.sleep(10)  # Add a 10-second delay
