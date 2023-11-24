# Excel Product Group Matcher

## Introduction
This Python script utilizes the `win32com.client` library to automate Microsoft Excel, specifically designed for matching and copying product group information based on stock codes between two sheets.

## Usage
1. Install the required library:
   ```bash
   pip install pywin32
Specify the file path of the Excel workbook in the file_path variable within the script.

Run the script. It performs the following actions:

Opens the Excel application.
Reads data from the "yazilacak" and "veriler" sheets.
Matches stock codes between the two sheets.
Updates the "yazilacak" sheet with corresponding product group information from the "veriler" sheet.
Saves and closes the workbook.
## Dependencies
Python 3.x
pywin32 library
## Example
### Example Input

# Specify the file path
```
file_path = "your_path"
```
Example Output
The "yazilacak" sheet is updated with product group information based on matching stock codes from the "veriler" sheet.

Notes
Ensure that the specified file path is correct.
Handle any exceptions that may occur during execution.
Feel free to customize the script according to your specific requirements.
This README.md file provides a brief overview of the script, its usage, dependencies, and an example input and output. Adjustments can be made based on your specific project details.