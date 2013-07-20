#!/usr/bin/python

## Program will create UPDATE, DELETE, or INSERT statements from one of:
##  - CSV File - unimplemented
##  - XLS File - unimplemented
##  - XLSX File - in progress

# CAVEATS:
#  - Program failed to parse file that I edited in LibreOffice. why?
#  - XLS file renamed to XLSX did not work with openpyxl

# Launch Web Browser w/ Data in File
#  - http://anh.cs.luc.edu/python/hands-on/3.1/handsonHtml/webtemplates.html


from Tkinter import *
from tkFileDialog import askopenfilename
from openpyxl.reader.excel import load_workbook
from openpyxl.cell import Cell
import sys

filename = ''
operation = 0 # Default 0 option is for SELECT

def getFile():
    filename = askopenfilename()
    instructions.grid_forget()
    loadFile.grid_forget()
    raw = getData(filename) # get
    numTables = len(raw)
    numFields = []

    for i in range(0,len(raw)):
        Label(frame, text='Table: ' + raw[i][0]).grid(row=2,column=0)

    print str(len(raw))
    print str(raw[1])

    fileLabel = Label(frame, text="Loaded: .." + filename[-36:])
    fileLabel.grid(row=1,column=0,columnspan=2)
    Radiobutton(frame, text="UPDATE", variable=operation, value=1).grid(row=3,column=0)
    Radiobutton(frame, text="INSERT", variable=operation, value=2).grid(row=4,column=0)
    Radiobutton(frame, text="DELETE", variable=operation, value=3).grid(row=5,column=0)
    go = Button(frame, text="CONVERT", command=getData(filename))
    go.grid(row=4,column=1,columnspan=2)

# directs to parsing functions for each file type
def getData(filename):
    if filename[-5:] == '.xlsx':
        return load_xlsx(filename)
    else:
        sys.exit()

# creates fields/values tuple from an XLSX spreadsheet
def load_xlsx(xlsxFile):
    fields = [] # populates from the first row in the spreadsheet
    values = [] # all the values in sheet (not including first row)
    dataRow = [] # a single row of data in the spreadsheet
    table = [] # array of the tableName, then fields and finally values
    database = [] # the entire spreadsheet in multi-dimensional array; return value
    
    wb = load_workbook(xlsxFile)
    sheetNames = wb.get_sheet_names() # sheetNames is Array of sheet names
    sheets = []

    for i in range(0,len(sheetNames)):
        sheets.append(wb.get_sheet_by_name(sheetNames[i]))

    for i in range(0,len(sheets)):
        for rIndex in range(0,len(sheets[i].rows)):
            for cIndex in range(0,len(sheets[i].columns)):
                if rIndex == 0: 
                    fields.append(str(sheets[i].cell(row=rIndex,column=cIndex).value))
                else:
                    dataRow.append(str(sheets[i].cell(row=rIndex,column=cIndex).value))
                if cIndex == (len(sheets[i].columns) - 1):
                    values.append(dataRow)
                    dataRow = []
            if rIndex == (len(sheets[i].rows) - 1):
                while 'None' in fields: # clean up any BLANK fields
                    fields.remove('None')
                table.append(sheetNames[i])
                table.append(fields)
                table.append(values)
                fields = []
                values = []

        database.append(table)
        table = []
    


    return database #return tuple of fields and values


root = Tk() # root widget, only 1 per program!
root.wm_title("Make a Statement")
frame = Frame(root, height=140, width=300)
frame.grid_propagate(0)
frame.grid()

instructions = Label(frame, text='for M$ Excel or CSV file.')
instructions.grid(row=0,column=1)
loadFile = Button(frame, text="Browse", command=getFile)
loadFile.grid(row=0,column=0)



root.mainloop()


## main program

