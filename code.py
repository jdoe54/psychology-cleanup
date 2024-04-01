import openpyxl
import config


filePath = config.FILE_PATH

try:

    spreadsheet = openpyxl.load_workbook(filePath)
    print("File found")

    original = spreadsheet["original"]
    clean = spreadsheet["clean"]

    for row in original.iter_rows(min_row=2, values_only=True):
        



except FileNotFoundError as error:
    print("No file found")