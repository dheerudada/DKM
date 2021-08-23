#! Python 3
# - Copy and Paste Ranges using OpenPyXl library
import xlrd
import openpyxl


# Copy range of cells as a nested list
# Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(startCol, startRow, endCol, endRow, sheet, UL, LL, NF):
    rangeSelected = []
    # Loops through selected Rows
    for i in range(startRow, endRow + 1, 1):
        # Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol, endCol + 1, 1):
            norm=sheet.cell(row=i, column=j).value
            if isinstance(norm, (float, int)) is True:
                #print(type(norm))
                norm=norm*NF
                if norm>UL or norm<LL:
                    norm=""
            rowSelected.append(norm)
        # Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)

    return rangeSelected


# Paste range
# Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
    countRow = 0
    for i in range(startRow, endRow + 1, 1):
        countCol = 0
        for j in range(startCol, endCol + 1, 1):
            sheetReceiving.cell(column=i, row=j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1




