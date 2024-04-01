import openpyxl
import config


filePath = config.FILE_PATH
dataRange = config.DATA_RANGE
timeLimit = config.TIME_LIMIT

try:

    spreadsheet = openpyxl.load_workbook(filePath)
    print("File found")

    original = spreadsheet["original"]
    clean = spreadsheet["clean"]

    newClean = spreadsheet.copy_worksheet(spreadsheet["original"])
    newClean.title = "new clean"

    index = 3

    print(newClean[config.DATA_RANGE])

    for row in newClean.iter_rows(min_row = index, values_only=True):
        dataDuration = newClean[config.DURATION_COL + str(index)].value

        if (int(dataDuration) < timeLimit):
            print("Time is less than " + str(timeLimit))
        else:
            print("Acceptable Time!")
        index = index + 1

    spreadsheet.save(filePath)



except FileNotFoundError as error:
    print("No file found")