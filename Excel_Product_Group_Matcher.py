import os
import numpy as np
import win32com.client
#Ürün gruplarını stok koduna göre eşleştirerek ürün gruplarını aktarmak için kullanıldı
def process_excel(file_path):
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        yazilacak_sheet = workbook.Sheets("yazilacak")
        veriler_sheet = workbook.Sheets("veriler")
        excel.Visible = False

        # Convert Excel sheets to NumPy arrays
        yazilacak_data = np.array(yazilacak_sheet.UsedRange.Value)
        veriler_data = np.array(veriler_sheet.UsedRange.Value)

        # Get column indices
        yazilacak_col_index = 0  # Adjust the column index as needed

        # Create a dictionary for faster lookups
        veriler_dict = dict(zip(veriler_data[:, 0], veriler_data[:, 2]))

        # Process yazilacak sheet
        for i in range(1, len(yazilacak_data)):
            yazilacak_value = yazilacak_data[i, yazilacak_col_index]

            # Check if the value is in veriler_dict
            if yazilacak_value in veriler_dict:
                yazilacak_sheet.Cells(i + 1, 8).Value = veriler_dict[yazilacak_value]

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel.Quit()

# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\SSH\\Python_Excel_Product_Serial_Number_Creation\\exceller\\c.xlsx"

# Check if the file exists
if os.path.exists(file_path):
    # Process the Excel file
    process_excel(file_path)
