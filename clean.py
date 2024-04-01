import openpyxl
import config

filePath = config.FILE_PATH
timeLimit = config.TIME_LIMIT

try:

    # First, it finds the file in your computer. If it works, then it will say "File found"
    spreadsheet = openpyxl.load_workbook(filePath)
    print("File found")

    original = spreadsheet["original"]
    clean = spreadsheet["clean"]

    # It makes a copy of the old original sheet, and calls the new sheet "new clean"
    newClean = spreadsheet.copy_worksheet(spreadsheet["original"])
    newClean.title = "new clean"

    rowNumber = 3
    deleteIndex = 0
    maxRow = 0
    deleteList = []

    print("Starting cleanup")

    # It goes through each row in the new copy of the sheet
    for row in newClean.iter_rows(min_row = rowNumber, values_only=True):
        dataDuration = newClean[config.DURATION_COL + str(rowNumber)].value

    # If the current row's time duration is less than the time limit (ex. 900 seconds), then it adds the row to the delete list.
        if (int(dataDuration) < timeLimit):
            deleteList.append(rowNumber)
        
    
        dataColumns = newClean[config.DATA_RANGE_LOWER + str(rowNumber) + ":" + config.DATA_RANGE_UPPER + str(rowNumber)]

    # It finds empty cells here, if there is no value, then it adds to the delete list and stops searching in the row (break).
        for row in dataColumns:
            for cell in row:
                if cell.value == None:
                    if rowNumber not in deleteList:
                        deleteList.append(rowNumber)
                    break
                

        rowNumber = rowNumber + 1

    # This tracks the number of rows in the sheet.
        maxRow = rowNumber
    
    # Once complete, it starts to delete the rows, it subtracts by delete index since deleting rows shifts the row number.
    for num in deleteList:
        newClean.delete_rows(num - deleteIndex)
        deleteIndex = deleteIndex + 1
        print("Deleting row " + str(num))

    # This says how many data rows are left, and then saves the excel file in the end.
    print("There are " + str(maxRow) + " currently. " + str(len(deleteList)) + " rows will be removed. New row count is " + str(maxRow - len(deleteList)) + ".")
    print("Complete!")
    
    spreadsheet.save(filePath)



except FileNotFoundError as error:
    # If it doesn't find the file, it says "No file found"
    print("No file found")