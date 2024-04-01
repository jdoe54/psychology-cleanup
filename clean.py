import openpyxl
import config

filePath = config.FILE_PATH
timeLimit = config.TIME_LIMIT

try:

    spreadsheet = openpyxl.load_workbook(filePath)
    print("File found")

    original = spreadsheet["original"]
    clean = spreadsheet["clean"]

    newClean = spreadsheet.copy_worksheet(spreadsheet["original"])
    newClean.title = "new clean"

    rowNumber = 3
    deleteIndex = 0
    maxRow = 0
    deleteList = []

    print("Starting cleanup")

    for row in newClean.iter_rows(min_row = rowNumber, values_only=True):
        dataDuration = newClean[config.DURATION_COL + str(rowNumber)].value

        if (int(dataDuration) < timeLimit):
            deleteList.append(rowNumber)
        

        dataColumns = newClean[config.DATA_RANGE_LOWER + str(rowNumber) + ":" + config.DATA_RANGE_UPPER + str(rowNumber)]

        for row in dataColumns:
            for cell in row:
                if cell.value == None:
                    if rowNumber not in deleteList:
                        deleteList.append(rowNumber)
                    break
                

        rowNumber = rowNumber + 1

        maxRow = rowNumber
    
    
    for num in deleteList:
        newClean.delete_rows(num - deleteIndex)
        deleteIndex = deleteIndex + 1
        print("Deleting row " + str(num))

    print("There are " + str(maxRow) + " currently. " + str(len(deleteList)) + " rows will be removed. New row count is " + str(maxRow - len(deleteList)) + ".")
    print("Complete!")
    
    spreadsheet.save(filePath)



except FileNotFoundError as error:
    print("No file found")