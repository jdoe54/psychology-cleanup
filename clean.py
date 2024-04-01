import openpyxl
import config


filePath = config.FILE_PATH
dataRange = config.DATA_RANGE

try:

    spreadsheet = openpyxl.load_workbook(filePath)
    print("File found")

    original = spreadsheet["original"]
    clean = spreadsheet["clean"]

    newSheet = spreadsheet.copy_worksheet(spreadsheet["original"])
    newSheet.title = "Copy"


    #for row in original.iter_rows(min_row=2, values_only=True):
    #    print(row["P"])

    spreadsheet.save(filePath)



except FileNotFoundError as error:
    print("No file found")