import os
import win32com.client

# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Fixidesk Excel Düzenleme\\old.xlsx"

# Check if the file exists
if os.path.exists(file_path):
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.Sheets("fixidesk")
        result_sheet = workbook.Sheets("logo")
        excel.Visible = False

        # Match fixidesk A columns cell and logo's D columns cell
        fixidesk_range = worksheet.Range("A:A")
        logo_range = result_sheet.Range("D:D")
        for fixidesk_cell in fixidesk_range:
            fixidesk_value = fixidesk_cell.Value
            for logo_cell in logo_range:
                logo_value = logo_cell.Value
                if fixidesk_value == logo_value:
                    # Get the value from logo B and write it to fixidesk C
                    fixidesk_cell.Offset(0, 2).Value = logo_cell.Offset(0, 1).Value
                    # Get the value from logo C and write it to fixidesk D
                    fixidesk_cell.Offset(0, 3).Value = logo_cell.Offset(0, 2).Value

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel.Quit()
        

# Rest of the code...
