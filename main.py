#!/usr/bin/python

## Program will create UPDATE, DELETE, or INSERT statements from one of:
##  - CSV File - unimplemented
##  - XLS File - unimplemented
##  - XLSX File - in progress

# CAVEATS:
#  - Program failed to parse file that I edited in LibreOffice. why?
#  - XLS file renamed to XLSX did not work with openpyxl


from Tkinter import *
from tkFileDialog import askopenfilename
from openpyxl.reader.excel import load_workbook
from openpyxl.cell import Cell
import sys

filename = ''

# directs to parsing functions for each file type
def execute(fname):
    if fname[-5:] == '.xlsx':
        load_xlsx(fname)
    else:
        sys.exit()

# we should divide this into two functions
#  - load_xlsx for making sure we can load the file and getting
#    total # of rows and columns so user can choose WHERE fields
#  - interpolate strings based on the data passed in

def load_xlsx(xlsxFile):
    pre = ""
    post = ""
    statement = ""
    update = "UPDATE "
    uSet = " SET "
    insert = "INSERT INTO "
    delete = "DELETE "
    where = " WHERE "
    fields = [] # populates from the first row in the spreadsheet
    
    wb = load_workbook(xlsxFile)
    sheetNames = wb.get_sheet_names() # sheetNames is Array of sheet names
    sheets = []

    # get the name of our sheets so we can refer to them as tables
    for i in range(0,len(sheetNames)):
        #DEBUG print sheetNames[i]
        #DEBUG print i
        sheets.append(wb.get_sheet_by_name(sheetNames[i]))

    # at this point we have all worksheets stored in the sheets Array

    # if operation = 1: #UPDATE
    for i in range(0,len(sheets)):
        for rIndex in range(0,len(sheets[i].rows)):
            statement = statement + update + sheetNames[i] + uSet
            for cIndex in range(0,len(sheets[i].columns)):
                if rIndex == 0: #first row has the fields
                    fields.append(sheets[i].cell(row=rIndex,column=cIndex).value)
                else:
                    statement = statement + str(fields[cIndex]) + " = '" + str(sheets[i].cell(row=rIndex,column=cIndex).value) + "'"
                if cIndex != (len(sheets[i].columns) - 1): # more columns so add a comma
                    statement = statement + ", "
                if cIndex == (len(sheets[i].columns) - 1): # last column so insert line break
                    # also add WHERE index value from user here
                    # {WHERE fields[userSuppliedIndex] = sheets[i].cell(row=rIndex,column=userSuppliedIndex).value}
                    statement = statement + "\n"
                    
    # else if operation = 2: #INSERT
    for i in range(0,len(sheets)):
        for rIndex in range(0,len(sheets[i].rows)):
            for cIndex in range(0,len(sheets[i].columns)):
                if rIndex == 0:
                    if cIndex != (len(sheets[i].columns) - 1): # first row, but not last column
                        pre = pre + str(sheets[i].cell(row=rIndex,column=cIndex).value) + ","
                    else: # must be the last column, close paren and add literals
                        pre = pre + str(sheets[i].cell(row=rIndex,column=cIndex).value) + ") VALUES ("
                if cIndex != (len(sheets[i].columns) - 1): # not last, add comma
                    post = post + "'" + str(sheets[i].cell(row=rIndex,column=cIndex).value) + "',"
                else: # last columns of this row
                    statement = statement + insert + sheetNames[i] + " (" + pre + post + "'" + str(sheets[i].cell(row=rIndex,column=cIndex).value) + "')\n"
                    post = "" # reset post every row

    # else if operation = 3: #DELETE


    

    statement = statement.replace("'NULL'", "NULL")
    print statement

    # DEBUG print sheets[0].cell('A1').value
    # DEBUG print sheets[0].cell(row=1,column=1).value
        
    # DEBUG print sheets[0]
    # DEBUG print filename + ' is an xlsx'

def getFile():
    filename = askopenfilename()
    fileLabel.grid(row=0,columnspan=4)
    Radiobutton(root, text="UPDATE", variable=operation, value=1).grid(row=1,column=1)
    Radiobutton(root, text="INSERT", variable=operation, value=2).grid(row=2,column=1)
    Radiobutton(root, text="DELETE", variable=operation, value=3).grid(row=3,column=1)
    go = Button(root, text="OK", command=execute(filename))
    go.grid(row=2,column=3)



root = Tk() # root widget, only 1 per program!
root.wm_title("Make a Statement")
f = Frame(root, height=100, width=300)
f.grid_propagate(0)
f.grid()

fileLabel = Label(root, text=filename)
loadFile = Button(root, text="Load File", command=getFile).grid(row=1,columnspan=1)


# WHEN WE WANT TO BULD EXIT:
# def callback():
#     sys.exit()

operation = 0 # Default 0 option is for SELECT


root.mainloop()


## main program

