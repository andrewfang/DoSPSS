import xlrd
from xlutils.copy import copy
import sys



"""
curr_row = -1
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    index = -1
    for cell in row:
        index += 1
        if type(cell.value) == unicode and 'r square change' in cell.value.lower():
            resultSheet.write(rSquareRowModel1, rSquareColModel1, worksheet.row(curr_row + 1)[index].value)
            rSquareRowModel1 += 1
            resultSheet.write(rSquareRowModel2, rSquareColModel2, worksheet.row(curr_row + 2)[index].value)
            rSquareRowModel2 += 1
result.save("resultOutput.xls")
"""




if __name__ == '__main__':
    FILENAME = raw_input("File Name > ")
    while 'xls' not in FILENAME:
        print "Please specify a xls or xlsx file"
        FILENAME = raw_input("File Name > ")
        if FILENAME == 'exit':
            sys.exit(0)

    print "===Starting SPSS Analysis==="
    print "+Now processing", FILENAME

    workbook = xlrd.open_workbook(FILENAME)
    worksheet = workbook.sheet_by_index(0)
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    curr_row = -1
    result = copy(xlrd.open_workbook('result.xls'))
    resultSheet = result.get_sheet(0)

    done = False
    fieldOfInterest = None
    numModels = 1
    base = 0
    fieldsfound = 0

    row = 0
    col = -1


    while not done:
        fieldOfInterest = raw_input("Name of field to gather > ")
        fieldTitle = raw_input("Name to title the field > ")
        fieldOfInterest = fieldOfInterest
        horizShift = int(raw_input("Horizontal Shift > "))
        vertShift = int(raw_input("Vert Shift > "))

        col += 1
        row = 0

        resultSheet.write(row, col, fieldTitle)
        row += 1

        curr_row = -1
        while curr_row < num_rows:
            curr_row += 1
            therow = worksheet.row(curr_row)
            index = -1
            for cell in therow:
                index += 1
                if type(cell.value) == unicode and fieldOfInterest in cell.value:
                    fieldsfound += 1
                    resultSheet.write(row, col, worksheet.row(curr_row + vertShift)[index + horizShift].value)
                    row += 1
        
        result.save("resultOutput.xls")

        areyoudone = raw_input("Are you done? > ")
        if 'y' in areyoudone:
            done = True
    print "I found", fieldsfound, "relevant fields"
    print "===Thanks for using DoSPSS v1.0!==="
    sys.exit(0)

