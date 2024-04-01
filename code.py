import openpyxl


filePath = ""

try:

    spreadsheet = openpyxl.load_workbook(filePath)
    print("File found")



except FileNotFoundError as error:
    print("No file found")