import os
import win32com.client
# Logo'dan fixidesk'e eşleşen fatura numarası ve tarihleri yazdırmak için kullanıldı
# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Python_Excel_Data_Matching_Copy\\old.xlsx"

# Check if the file exists
if os.path.exists(file_path):
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.Sheets("fixidesk")
        result_sheet = workbook.Sheets("logo")
        excel.Visible = False

        # Read data into dictionaries for fixidesk and logo columns
        fixidesk_data = {}
        logo_data = {}
        fixidesk_range = worksheet.Range("A:A")
        logo_range = result_sheet.Range("D:D")
        for i, fixidesk_value in enumerate(fixidesk_range.Value):
            fixidesk_data[fixidesk_value] = i + 1
        for i, logo_value in enumerate(logo_range.Value):
            logo_data[logo_value] = i + 1

        # Find matching values and update fixidesk columns
        for fixidesk_value, fixidesk_row in fixidesk_data.items():
            if fixidesk_value in logo_data:
                logo_row = logo_data[fixidesk_value]
                # Get the value from logo B and write it to fixidesk C
                worksheet.Cells(fixidesk_row, 3).Value = result_sheet.Cells(logo_row, 2).Value
                # Get the value from logo C and write it to fixidesk D
                worksheet.Cells(fixidesk_row, 4).Value = result_sheet.Cells(logo_row, 3).Value

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel.Quit()
