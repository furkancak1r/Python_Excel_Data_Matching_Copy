import os
import numpy as np
import openpyxl
#Ürün gruplarını stok koduna göre eşleştirerek ürün gruplarını aktarmak için kullanıldı
def process_excel(file_path):
    try:
        # Load the workbook and the worksheets
        workbook = openpyxl.load_workbook(file_path)
        yazilacak_sheet = workbook["yazilacak"]
        veriler_sheet = workbook["veriler"]

        # Convert Excel sheets to NumPy arrays
        yazilacak_data = np.array(list(yazilacak_sheet.values))
        veriler_data = np.array(list(veriler_sheet.values))

        # Get column indices
        yazilacak_col_index = 5  # Adjust the column index as needed

        # Create a dictionary for faster lookups
        veriler_dict = dict(zip(veriler_data[:, 0], veriler_data[:, 1]))
        # Process yazilacak sheet
        for i in range(1, len(yazilacak_data)):
            yazilacak_value = yazilacak_data[i][yazilacak_col_index]

            # Check if the value is in veriler_dict
            if yazilacak_value in veriler_dict:
                yazilacak_sheet.cell(row=i + 1, column=13).value = veriler_dict[yazilacak_value]

        # Save and close the workbook
        workbook.save(file_path)
        workbook.close()

    except Exception as e:
        print(f"An error occurred: {e}")

# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\SSH\\SSH Exceller\\ekipman kartı aktarım şablonu çalışması.xlsx"

# Check if the file exists
if os.path.exists(file_path):
    # Process the Excel file
    process_excel(file_path)
else:
    print(f"File not found: {file_path}")
